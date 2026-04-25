from __future__ import annotations

import json
import queue
import random
import threading
import time
import ctypes
from dataclasses import dataclass
from pathlib import Path
from tkinter import END, DoubleVar, IntVar, StringVar, filedialog, messagebox

import customtkinter as ctk
import numpy as np
import soundcard as sc
from pynput.keyboard import Controller, Key
from pynput.mouse import Button as MouseButton


APP_DIR = Path(__file__).resolve().parent.parent
PROFILE_PATH = APP_DIR / "profile.json"
SAMPLE_RATE = 44_100
BLOCK_SIZE = 1024
FINGERPRINT_BINS = 64
FINGERPRINT_WINDOW = 2048
FINGERPRINT_HOP = 1024
FINGERPRINT_WINDOW_VECTOR = np.hanning(FINGERPRINT_WINDOW).astype(np.float32)
DETECTION_INTERVAL_SECONDS = 0.05
SCORE_EVENT_INTERVAL_SECONDS = 0.10
MAX_UI_EVENTS_PER_TICK = 200
MAX_LOG_LINES = 500

MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP = 0x0040


SPECIAL_KEYS = {
    "alt": Key.alt,
    "alt_l": Key.alt_l,
    "alt_r": Key.alt_r,
    "backspace": Key.backspace,
    "caps_lock": Key.caps_lock,
    "cmd": Key.cmd,
    "ctrl": Key.ctrl,
    "ctrl_l": Key.ctrl_l,
    "ctrl_r": Key.ctrl_r,
    "delete": Key.delete,
    "down": Key.down,
    "end": Key.end,
    "enter": Key.enter,
    "esc": Key.esc,
    "escape": Key.esc,
    "f1": Key.f1,
    "f2": Key.f2,
    "f3": Key.f3,
    "f4": Key.f4,
    "f5": Key.f5,
    "f6": Key.f6,
    "f7": Key.f7,
    "f8": Key.f8,
    "f9": Key.f9,
    "f10": Key.f10,
    "f11": Key.f11,
    "f12": Key.f12,
    "home": Key.home,
    "insert": Key.insert,
    "left": Key.left,
    "page_down": Key.page_down,
    "page_up": Key.page_up,
    "right": Key.right,
    "shift": Key.shift,
    "shift_l": Key.shift_l,
    "shift_r": Key.shift_r,
    "space": Key.space,
    "tab": Key.tab,
    "up": Key.up,
}


@dataclass
class MacroStep:
    action: str
    delay_min_ms: int
    delay_max_ms: int

    def delay_ms(self) -> int:
        if self.delay_max_ms <= self.delay_min_ms:
            return self.delay_min_ms
        return random.randint(self.delay_min_ms, self.delay_max_ms)


def mono(samples: np.ndarray) -> np.ndarray:
    if samples.ndim == 1:
        return samples.astype(np.float32)
    return samples.mean(axis=1).astype(np.float32)


def normalize(samples: np.ndarray) -> np.ndarray:
    samples = samples.astype(np.float32)
    samples = samples - float(np.mean(samples))
    peak = float(np.max(np.abs(samples))) if samples.size else 0.0
    if peak < 1e-6:
        return samples
    return samples / peak


def fingerprint(samples: np.ndarray) -> np.ndarray:
    samples = normalize(samples)
    if len(samples) < FINGERPRINT_WINDOW:
        return np.array([], dtype=np.float32)

    frames: list[np.ndarray] = []
    for start in range(0, len(samples) - FINGERPRINT_WINDOW + 1, FINGERPRINT_HOP):
        frame = samples[start : start + FINGERPRINT_WINDOW] * FINGERPRINT_WINDOW_VECTOR
        spectrum = np.abs(np.fft.rfft(frame))[1 : FINGERPRINT_BINS + 1]
        spectrum = np.log1p(spectrum).astype(np.float32)
        spectrum -= float(np.mean(spectrum))
        norm = float(np.linalg.norm(spectrum))
        if norm > 1e-6:
            spectrum /= norm
        frames.append(spectrum)

    if not frames:
        return np.array([], dtype=np.float32)
    result = np.stack(frames).astype(np.float32).reshape(-1)
    norm = float(np.linalg.norm(result))
    if norm > 1e-6:
        result /= norm
    return result


def fingerprint_similarity(left: np.ndarray, right: np.ndarray) -> float:
    if left.size == 0 or right.size == 0 or left.size != right.size:
        return 0.0
    return float(np.dot(left, right))


def parse_delay(value: str, line_number: int) -> tuple[int, int]:
    value = value.strip()
    if not value:
        return 0, 0

    if "-" in value:
        left, right = [part.strip() for part in value.split("-", maxsplit=1)]
        try:
            delay_min = int(left)
            delay_max = int(right)
        except ValueError as exc:
            raise ValueError(f"Line {line_number}: delay range must be integers") from exc
        if delay_min > delay_max:
            raise ValueError(f"Line {line_number}: delay range min must be <= max")
        if delay_min < 0:
            raise ValueError(f"Line {line_number}: delay must be >= 0")
        return delay_min, delay_max

    try:
        delay = int(value)
    except ValueError as exc:
        raise ValueError(f"Line {line_number}: delay must be an integer or min-max range") from exc
    if delay < 0:
        raise ValueError(f"Line {line_number}: delay must be >= 0")
    return delay, delay


def parse_macro(text: str) -> list[MacroStep]:
    steps: list[MacroStep] = []
    for line_number, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        parts = [part.strip() for part in line.split(",", maxsplit=1)]
        action = parts[0]
        delay_min_ms = 0
        delay_max_ms = 0
        if len(parts) == 2 and parts[1]:
            delay_min_ms, delay_max_ms = parse_delay(parts[1], line_number)
        if not action:
            raise ValueError(f"Line {line_number}: missing action")
        steps.append(MacroStep(action=action, delay_min_ms=delay_min_ms, delay_max_ms=delay_max_ms))

    if not steps:
        raise ValueError("Macro is empty")
    return steps


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


class MacroRunner:
    def __init__(self) -> None:
        self.keyboard = Controller()
        self.lock = threading.Lock()

    def run(self, steps: list[MacroStep], jitter_ms: int = 0) -> None:
        with self.lock:
            for step in steps:
                self.run_action(step.action, jitter_ms)
                delay_ms = step.delay_ms()
                if delay_ms:
                    time.sleep(delay_ms / 1000)

    def run_action(self, action: str, jitter_ms: int) -> None:
        normalized = action.strip().lower()
        if normalized in {"left_click", "click_left", "mouse_left"}:
            self.click(MouseButton.left, jitter_ms)
            return
        if normalized in {"right_click", "click_right", "mouse_right", "rclick"}:
            self.click(MouseButton.right, jitter_ms)
            return
        if normalized in {"middle_click", "click_middle", "mouse_middle"}:
            self.click(MouseButton.middle, jitter_ms)
            return

        self.press_combo(action, jitter_ms)

    def click(self, button: MouseButton, jitter_ms: int) -> None:
        down_flag, up_flag = self.mouse_flags(button)
        ctypes.windll.user32.mouse_event(down_flag, 0, 0, 0, 0)
        self.sleep_jitter(jitter_ms)
        ctypes.windll.user32.mouse_event(up_flag, 0, 0, 0, 0)
        self.sleep_jitter(jitter_ms)

    @staticmethod
    def mouse_flags(button: MouseButton) -> tuple[int, int]:
        if button == MouseButton.left:
            return MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP
        if button == MouseButton.right:
            return MOUSEEVENTF_RIGHTDOWN, MOUSEEVENTF_RIGHTUP
        if button == MouseButton.middle:
            return MOUSEEVENTF_MIDDLEDOWN, MOUSEEVENTF_MIDDLEUP
        raise ValueError(f"Unsupported mouse button: {button}")

    def press_combo(self, combo: str, jitter_ms: int) -> None:
        keys = [self.to_key(part.strip().lower()) for part in combo.split("+") if part.strip()]
        if not keys:
            return

        for key in keys:
            self.keyboard.press(key)
            self.sleep_jitter(jitter_ms)
        for key in reversed(keys):
            self.keyboard.release(key)
            self.sleep_jitter(jitter_ms)

    @staticmethod
    def sleep_jitter(jitter_ms: int) -> None:
        if jitter_ms <= 0:
            time.sleep(0.01)
            return
        time.sleep(random.randint(0, jitter_ms) / 1000)

    @staticmethod
    def to_key(value: str):
        if value in SPECIAL_KEYS:
            return SPECIAL_KEYS[value]
        if len(value) == 1:
            return value
        raise ValueError(f"Unsupported key: {value}")


class AudioDetector(threading.Thread):
    def __init__(
        self,
        device_name: str,
        sample: np.ndarray,
        detect_seconds: float,
        threshold: float,
        cooldown_s: float,
        event_queue: queue.Queue[tuple[str, object]],
        stop_event: threading.Event,
    ) -> None:
        super().__init__(daemon=True)
        self.device_name = device_name
        detect_frames = max(FINGERPRINT_WINDOW, int(SAMPLE_RATE * detect_seconds))
        self.sample = normalize(sample[: min(len(sample), detect_frames)])
        self.sample_fingerprint = fingerprint(self.sample)
        self.threshold = threshold
        self.cooldown_s = cooldown_s
        self.event_queue = event_queue
        self.stop_event = stop_event
        self.last_trigger = 0.0
        self.last_score_event = 0.0
        self.last_analysis = 0.0

    def run(self) -> None:
        try:
            device = next(device for device in sc.all_microphones(include_loopback=True) if device.name == self.device_name)
        except StopIteration:
            self.event_queue.put(("error", "Selected audio device was not found"))
            return

        sample_len = len(self.sample)
        if sample_len < BLOCK_SIZE:
            self.event_queue.put(("error", "Recorded sample is too short"))
            return

        rolling = np.zeros(sample_len + BLOCK_SIZE * 2, dtype=np.float32)

        try:
            with device.recorder(samplerate=SAMPLE_RATE, channels=2, blocksize=BLOCK_SIZE) as recorder:
                while not self.stop_event.is_set():
                    block = mono(recorder.record(numframes=BLOCK_SIZE))
                    block_len = len(block)
                    rolling[:-block_len] = rolling[block_len:]
                    rolling[-block_len:] = block

                    now = time.monotonic()
                    if now - self.last_analysis < DETECTION_INTERVAL_SECONDS:
                        continue
                    self.last_analysis = now

                    score = self.best_score(rolling, sample_len)
                    if now - self.last_score_event >= SCORE_EVENT_INTERVAL_SECONDS:
                        self.last_score_event = now
                        self.event_queue.put(("score", score))

                    if score >= self.threshold and now - self.last_trigger >= self.cooldown_s:
                        self.last_trigger = now
                        self.event_queue.put(("trigger", score))
        except Exception as exc:
            self.event_queue.put(("error", str(exc)))

    def best_score(self, rolling: np.ndarray, sample_len: int) -> float:
        best = 0.0
        offsets = (0, BLOCK_SIZE // 2, BLOCK_SIZE)
        for offset in offsets:
            end = len(rolling) - offset
            start = end - sample_len
            if start < 0:
                continue
            window = rolling[start:end]
            score = fingerprint_similarity(fingerprint(window), self.sample_fingerprint)
            if score > best:
                best = score
        return best


class AudioMacroApp:
    def __init__(self, root: ctk.CTk) -> None:
        self.root = root
        self.root.title("Audio Macro Trigger")
        self.root.geometry("1200x780")
        self.root.minsize(1080, 700)

        self.devices: list[sc.Microphone] = []
        self.sample = np.array([], dtype=np.float32)
        self.detector: AudioDetector | None = None
        self.stop_event = threading.Event()
        self.recording = False
        self.playing_sample = False
        self.macro_running = False
        self.log_messages: list[str] = []
        self.log_render_pending = False
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.runner = MacroRunner()
        self.admin = is_admin()

        self.device_var = StringVar()
        self.sample_seconds_var = DoubleVar(value=2.0)
        self.detect_seconds_var = DoubleVar(value=0.8)
        self.threshold_var = DoubleVar(value=0.35)
        self.cooldown_var = DoubleVar(value=2.0)
        self.jitter_var = IntVar(value=30)
        self.status_var = StringVar(value=f"Idle ({'Administrator' if self.admin else 'User'})")
        self.score_var = StringVar(value="Score: -")

        self.configure_style()
        self.build_ui()
        self.refresh_devices()
        self.load_profile()
        self.root.after(100, self.consume_events)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def configure_style(self) -> None:
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

    def build_ui(self) -> None:
        self.root.configure(fg_color="#eef2f6")

        outer = ctk.CTkFrame(self.root, fg_color="#eef2f6", corner_radius=0)
        outer.pack(fill="both", expand=True, padx=22, pady=20)
        outer.columnconfigure(0, weight=0)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(1, weight=1)

        header = ctk.CTkFrame(outer, fg_color="transparent")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 18))
        header.columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Audio Macro Trigger", font=ctk.CTkFont(size=24, weight="bold"), text_color="#111827").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(header, text="Listen for system audio, then run precise keyboard and mouse macros.", font=ctk.CTkFont(size=13), text_color="#5f6b7a").grid(row=1, column=0, sticky="w", pady=(2, 0))
        ctk.CTkLabel(header, textvariable=self.status_var, fg_color="#ffffff", text_color="#253040", corner_radius=10, padx=14, pady=8).grid(row=0, column=1, rowspan=2, sticky="e")

        sidebar = ctk.CTkScrollableFrame(
            outer,
            fg_color="transparent",
            width=360,
            height=640,
            scrollbar_button_color="#cbd5e1",
            scrollbar_button_hover_color="#94a3b8",
        )
        sidebar.grid(row=1, column=0, sticky="nsew", padx=(0, 16))
        sidebar.columnconfigure(0, weight=1)

        workspace = ctk.CTkFrame(outer, fg_color="transparent")
        workspace.grid(row=1, column=1, sticky="nsew")
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(1, weight=3)
        workspace.rowconfigure(2, weight=2)

        audio_group = self.create_card(sidebar, "Audio Source", "Choose the output device to monitor.")
        audio_group.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        audio_group.columnconfigure(0, weight=1)
        audio_group.columnconfigure(1, weight=0)
        self.device_combo = ctk.CTkComboBox(audio_group, variable=self.device_var, values=[], state="readonly", height=36, border_color="#cbd5e1", button_color="#2563eb", button_hover_color="#1d4ed8")
        self.device_combo.grid(row=2, column=0, sticky="ew", padx=(14, 8), pady=(0, 14))
        ctk.CTkButton(audio_group, text="Refresh", command=self.refresh_devices, width=92, height=36, fg_color="#334155", hover_color="#1f2937").grid(row=2, column=1, padx=(0, 14), pady=(0, 14))

        tuning_group = self.create_card(sidebar, "Detection", "Tune speed, accuracy and repeat behavior.")
        tuning_group.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        tuning_group.columnconfigure(0, weight=1)
        tuning_group.columnconfigure(1, weight=1)
        self.add_number_control(tuning_group, "Sample seconds", self.sample_seconds_var, 0.5, 10.0, 0.5, 2, 0)
        self.add_number_control(tuning_group, "Detect seconds", self.detect_seconds_var, 0.1, 5.0, 0.1, 2, 1)
        self.add_number_control(tuning_group, "Threshold", self.threshold_var, 0.1, 0.95, 0.05, 3, 0)
        self.add_number_control(tuning_group, "Cooldown seconds", self.cooldown_var, 0.0, 30.0, 0.5, 3, 1)
        self.add_number_control(tuning_group, "Action jitter ms", self.jitter_var, 0, 500, 5, 4, 0)

        sample_group = self.create_card(sidebar, "Sample & Control", "Capture, verify and test before listening.")
        sample_group.grid(row=2, column=0, sticky="ew")
        sample_group.columnconfigure(0, weight=1)
        sample_group.columnconfigure(1, weight=1)
        self.record_button = ctk.CTkButton(sample_group, text="Record sample", command=self.record_sample, height=38, fg_color="#2563eb", hover_color="#1d4ed8")
        self.record_button.grid(row=2, column=0, sticky="ew", padx=(14, 6), pady=(0, 10))
        self.play_sample_button = ctk.CTkButton(sample_group, text="Play sample", command=self.play_sample, state="disabled", height=38, fg_color="#0f766e", hover_color="#115e59")
        self.play_sample_button.grid(row=2, column=1, sticky="ew", padx=(6, 14), pady=(0, 10))
        ctk.CTkButton(sample_group, text="Import sample", command=self.import_sample, height=36, fg_color="#64748b", hover_color="#475569").grid(row=3, column=0, sticky="ew", padx=(14, 6), pady=(0, 10))
        ctk.CTkButton(sample_group, text="Save profile", command=self.save_profile, height=36, fg_color="#64748b", hover_color="#475569").grid(row=3, column=1, sticky="ew", padx=(6, 14), pady=(0, 10))
        ctk.CTkButton(sample_group, text="Test macro", command=self.test_macro, height=36, fg_color="#475569", hover_color="#334155").grid(row=4, column=0, sticky="ew", padx=(14, 6), pady=(0, 12))
        self.listen_button = ctk.CTkButton(sample_group, text="Start listening", command=self.toggle_listening, height=36, fg_color="#15803d", hover_color="#166534")
        self.listen_button.grid(row=4, column=1, sticky="ew", padx=(6, 14), pady=(0, 12))
        self.activity_bar = ctk.CTkProgressBar(sample_group, mode="indeterminate", height=8, corner_radius=4, progress_color="#2563eb")
        self.activity_bar.grid(row=5, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 14))
        self.activity_bar.set(0)
        if not self.admin:
            ctk.CTkLabel(sample_group, text="Not running as Administrator. Elevated apps may ignore macro keys.", wraplength=310, justify="left", text_color="#a16207").grid(row=6, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 14))

        score_group = self.create_card(workspace, "Live Detector", "Watch matching confidence while the target sound plays.")
        score_group.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        score_group.columnconfigure(0, weight=1)
        score_header = ctk.CTkFrame(score_group, fg_color="transparent")
        score_header.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 8))
        score_header.columnconfigure(0, weight=1)
        ctk.CTkLabel(score_header, textvariable=self.score_var, font=ctk.CTkFont(size=16, weight="bold"), text_color="#111827").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(score_header, text="Trigger when score >= Threshold", text_color="#66717f").grid(row=0, column=1, sticky="e")
        self.score_bar = ctk.CTkProgressBar(score_group, mode="determinate", height=12, corner_radius=8, progress_color="#16a34a")
        self.score_bar.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 14))
        self.score_bar.set(0)

        macro_group = self.create_card(workspace, "Macro Steps", "One action per line: action, delay_ms or action, min-max.")
        macro_group.grid(row=1, column=0, sticky="nsew", pady=(0, 12))
        macro_group.rowconfigure(2, weight=1)
        macro_group.columnconfigure(0, weight=1)
        self.macro_text = ctk.CTkTextbox(macro_group, wrap="none", font=ctk.CTkFont(family="Consolas", size=13), border_width=1, border_color="#cbd5e1", fg_color="#fbfcfd", text_color="#17202a")
        self.macro_text.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 14))
        self.macro_text.insert(END, "right_click, 900-1100\nright_click, 0\n")

        events_group = self.create_card(workspace, "Events", "Runtime log for detection, sample capture and macro execution.")
        events_group.grid(row=2, column=0, sticky="nsew")
        events_group.rowconfigure(2, weight=1)
        events_group.columnconfigure(0, weight=1)
        self.event_list = ctk.CTkTextbox(events_group, font=ctk.CTkFont(family="Consolas", size=12), border_width=1, border_color="#cbd5e1", fg_color="#fbfcfd", text_color="#253040")
        self.event_list.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 14))

    @staticmethod
    def create_card(parent, title: str, subtitle: str = "") -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color="#ffffff", corner_radius=12, border_width=1, border_color="#d9e2ec")
        card.columnconfigure(0, weight=1)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=15, weight="bold"), text_color="#111827").grid(row=0, column=0, columnspan=8, sticky="w", padx=14, pady=(14, 2))
        if subtitle:
            ctk.CTkLabel(card, text=subtitle, text_color="#66717f", wraplength=620, justify="left").grid(row=1, column=0, columnspan=8, sticky="w", padx=14, pady=(0, 12))
        return card

    def add_number_control(self, parent, label: str, variable, from_: float, to: float, increment: float, row: int, column: int) -> None:
        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.grid(row=row, column=column, sticky="ew", padx=14, pady=(0, 12))
        box.columnconfigure(1, weight=1)
        ctk.CTkLabel(box, text=label, text_color="#56616f", font=ctk.CTkFont(size=12)).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))
        ctk.CTkButton(box, text="-", width=28, height=30, fg_color="#e2e8f0", hover_color="#cbd5e1", text_color="#17202a", command=lambda: self.adjust_number(variable, from_, to, -increment)).grid(row=1, column=0, sticky="w")
        ctk.CTkEntry(box, textvariable=variable, width=74, height=30, justify="center", border_color="#cbd5e1").grid(row=1, column=1, sticky="ew", padx=5)
        ctk.CTkButton(box, text="+", width=28, height=30, fg_color="#e2e8f0", hover_color="#cbd5e1", text_color="#17202a", command=lambda: self.adjust_number(variable, from_, to, increment)).grid(row=1, column=2, sticky="e")

    @staticmethod
    def adjust_number(variable, from_: float, to: float, delta: float) -> None:
        try:
            current = float(variable.get())
        except Exception:
            current = from_
        next_value = min(to, max(from_, current + delta))
        if isinstance(variable, IntVar):
            variable.set(int(round(next_value)))
        else:
            variable.set(round(next_value, 3))

    def refresh_devices(self) -> None:
        self.devices = list(sc.all_microphones(include_loopback=True))
        names = [device.name for device in self.devices]
        self.device_combo.configure(values=names)
        loopback = next((name for name in names if "loopback" in name.lower()), names[0] if names else "")
        if loopback and not self.device_var.get():
            self.device_combo.set(loopback)
        self.log(f"Found {len(names)} audio input/loopback devices")

    def record_sample(self) -> None:
        if self.recording:
            return
        if not self.device_var.get():
            messagebox.showerror("Missing device", "Select an audio device first")
            return
        seconds = float(self.sample_seconds_var.get())
        self.recording = True
        self.record_button.configure(text="Recording...", state="disabled")
        self.play_sample_button.configure(state="disabled")
        self.activity_bar.start()
        self.status_var.set(f"Recording {seconds:.1f}s sample...")
        self.log(f"Recording sample for {seconds:.1f}s")
        threading.Thread(target=self._record_sample_worker, args=(seconds,), daemon=True).start()

    def _record_sample_worker(self, seconds: float) -> None:
        try:
            device = next(device for device in self.devices if device.name == self.device_var.get())
            with device.recorder(samplerate=SAMPLE_RATE, channels=2, blocksize=BLOCK_SIZE) as recorder:
                samples = recorder.record(numframes=int(SAMPLE_RATE * seconds))
            self.sample = mono(samples)
            self.events.put(("sample", len(self.sample)))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def play_sample(self) -> None:
        if self.sample.size == 0:
            messagebox.showerror("Missing sample", "Record or import a sample first")
            return
        if self.playing_sample:
            return
        self.playing_sample = True
        self.play_sample_button.configure(text="Playing...", state="disabled")
        self.activity_bar.start()
        self.status_var.set(f"Playing sample: {len(self.sample) / SAMPLE_RATE:.2f}s")
        self.log("Playing sample")
        threading.Thread(target=self._play_sample_worker, daemon=True).start()

    def _play_sample_worker(self) -> None:
        try:
            sample = self.sample.astype(np.float32)
            playback = np.column_stack([sample, sample])
            sc.default_speaker().play(playback, samplerate=SAMPLE_RATE)
            self.events.put(("playback_done", len(sample)))
        except Exception as exc:
            self.events.put(("error", str(exc)))

    def import_sample(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("NumPy sample", "*.npy")])
        if not path:
            return
        try:
            self.sample = np.load(path).astype(np.float32)
            self.log(f"Imported sample: {Path(path).name}")
            self.status_var.set(f"Sample loaded: {len(self.sample) / SAMPLE_RATE:.2f}s")
            self.set_sample_ready(True)
        except Exception as exc:
            messagebox.showerror("Import failed", str(exc))

    def toggle_listening(self) -> None:
        if self.detector and self.detector.is_alive():
            self.stop_listening()
        else:
            self.start_listening()

    def test_macro(self) -> None:
        try:
            steps = parse_macro(self.macro_text.get("1.0", END))
        except ValueError as exc:
            messagebox.showerror("Invalid macro", str(exc))
            return
        self.log("Manual macro test started")
        self.start_macro_thread(steps)

    def start_listening(self) -> None:
        if self.sample.size == 0:
            messagebox.showerror("Missing sample", "Record or import a sample first")
            return
        try:
            parse_macro(self.macro_text.get("1.0", END))
        except ValueError as exc:
            messagebox.showerror("Invalid macro", str(exc))
            return

        self.stop_event.clear()
        self.detector = AudioDetector(
            device_name=self.device_var.get(),
            sample=self.sample,
            detect_seconds=float(self.detect_seconds_var.get()),
            threshold=float(self.threshold_var.get()),
            cooldown_s=float(self.cooldown_var.get()),
            event_queue=self.events,
            stop_event=self.stop_event,
        )
        self.detector.start()
        self.listen_button.configure(text="Stop listening", fg_color="#dc2626", hover_color="#b91c1c")
        self.status_var.set("Listening")
        self.log("Listening started")

    def stop_listening(self) -> None:
        self.stop_event.set()
        self.listen_button.configure(text="Start listening", fg_color="#15803d", hover_color="#166534")
        self.status_var.set(f"Idle ({'Administrator' if self.admin else 'User'})")
        self.log("Listening stopped")

    def consume_events(self) -> None:
        processed = 0
        pending_score: float | None = None
        while processed < MAX_UI_EVENTS_PER_TICK:
            try:
                event, payload = self.events.get_nowait()
            except queue.Empty:
                break
            processed += 1

            if event == "score":
                pending_score = float(payload)
            elif event == "trigger":
                score = float(payload)
                self.log(f"Triggered at score {score:.3f}")
                try:
                    steps = parse_macro(self.macro_text.get("1.0", END))
                    self.start_macro_thread(steps)
                except ValueError as exc:
                    self.log(str(exc))
            elif event == "macro_done":
                self.log("Macro finished")
            elif event == "macro_idle":
                self.macro_running = False
            elif event == "sample":
                seconds = int(payload) / SAMPLE_RATE
                self.recording = False
                self.record_button.configure(text="Record sample", state="normal")
                self.set_sample_ready(True)
                self.activity_bar.stop()
                self.status_var.set(f"Sample recorded: {seconds:.2f}s")
                self.log(f"Recording finished: {seconds:.2f}s")
            elif event == "playback_done":
                seconds = int(payload) / SAMPLE_RATE
                self.playing_sample = False
                self.play_sample_button.configure(text="Play sample", state="normal")
                self.activity_bar.stop()
                self.status_var.set(f"Sample playback finished: {seconds:.2f}s")
                self.log("Sample playback finished")
            elif event == "error":
                if self.recording:
                    self.recording = False
                    self.record_button.configure(text="Record sample", state="normal")
                    self.set_sample_ready(self.sample.size > 0)
                    self.activity_bar.stop()
                if self.playing_sample:
                    self.playing_sample = False
                    self.play_sample_button.configure(text="Play sample", state="normal" if self.sample.size else "disabled")
                    self.activity_bar.stop()
                self.status_var.set("Error")
                self.log(f"Error: {payload}")

        if pending_score is not None:
            self.score_var.set(f"Score: {pending_score:.3f}")
            self.score_bar.set(min(1.0, max(0.0, pending_score)))

        next_delay = 20 if not self.events.empty() else 100
        self.root.after(next_delay, self.consume_events)

    def start_macro_thread(self, steps: list[MacroStep]) -> None:
        if self.macro_running:
            self.log("Macro already running; trigger skipped")
            return
        self.macro_running = True
        jitter_ms = int(self.jitter_var.get())
        threading.Thread(target=self._macro_worker, args=(steps, jitter_ms), daemon=True).start()

    def _macro_worker(self, steps: list[MacroStep], jitter_ms: int) -> None:
        try:
            self.runner.run(steps, jitter_ms)
            self.events.put(("macro_done", None))
        except Exception as exc:
            self.events.put(("error", f"Macro failed: {exc}"))
        finally:
            self.events.put(("macro_idle", None))

    def save_profile(self) -> None:
        sample_path = APP_DIR / "sample.npy"
        if self.sample.size:
            np.save(sample_path, self.sample)

        profile = {
            "device": self.device_var.get(),
            "sample_seconds": float(self.sample_seconds_var.get()),
            "detect_seconds": float(self.detect_seconds_var.get()),
            "threshold": float(self.threshold_var.get()),
            "cooldown": float(self.cooldown_var.get()),
            "jitter": int(self.jitter_var.get()),
            "macro": self.macro_text.get("1.0", END).strip(),
            "sample_path": str(sample_path.name) if self.sample.size else "",
        }
        PROFILE_PATH.write_text(json.dumps(profile, indent=2), encoding="utf-8")
        self.log("Profile saved")

    def load_profile(self) -> None:
        if not PROFILE_PATH.exists():
            return
        try:
            profile = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
            self.device_var.set(profile.get("device", self.device_var.get()))
            self.sample_seconds_var.set(float(profile.get("sample_seconds", self.sample_seconds_var.get())))
            self.detect_seconds_var.set(float(profile.get("detect_seconds", self.detect_seconds_var.get())))
            self.threshold_var.set(float(profile.get("threshold", self.threshold_var.get())))
            self.cooldown_var.set(float(profile.get("cooldown", self.cooldown_var.get())))
            self.jitter_var.set(int(profile.get("jitter", self.jitter_var.get())))
            self.macro_text.delete("1.0", END)
            self.macro_text.insert(END, profile.get("macro", ""))
            sample_path = APP_DIR / profile.get("sample_path", "")
            if sample_path.exists():
                self.sample = np.load(sample_path).astype(np.float32)
                self.status_var.set(f"Sample loaded: {len(self.sample) / SAMPLE_RATE:.2f}s")
                self.set_sample_ready(True)
            self.log("Profile loaded")
        except Exception as exc:
            self.log(f"Profile load failed: {exc}")

    def set_sample_ready(self, ready: bool) -> None:
        self.play_sample_button.configure(state="normal" if ready and not self.playing_sample else "disabled")

    def log(self, message: str) -> None:
        timestamp = time.strftime("%H:%M:%S")
        self.log_messages.append(f"{timestamp}  {message}")
        if len(self.log_messages) > MAX_LOG_LINES:
            self.log_messages = self.log_messages[-MAX_LOG_LINES:]
        if not self.log_render_pending:
            self.log_render_pending = True
            self.root.after(50, self.render_log)

    def render_log(self) -> None:
        self.log_render_pending = False
        self.event_list.delete("1.0", END)
        if self.log_messages:
            self.event_list.insert(END, "\n".join(self.log_messages) + "\n")
            self.event_list.see(END)

    def on_close(self) -> None:
        self.stop_event.set()
        self.root.destroy()


def main() -> None:
    root = ctk.CTk()
    AudioMacroApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
