from __future__ import annotations

import json
import queue
import random
import threading
import time
import ctypes
from dataclasses import dataclass
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, DoubleVar, Frame, IntVar, Listbox, StringVar, Text, Tk
from tkinter import filedialog, messagebox, ttk

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

    window = np.hanning(FINGERPRINT_WINDOW).astype(np.float32)
    frames: list[np.ndarray] = []
    for start in range(0, len(samples) - FINGERPRINT_WINDOW + 1, FINGERPRINT_HOP):
        frame = samples[start : start + FINGERPRINT_WINDOW] * window
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
                    rolling = np.concatenate([rolling[len(block) :], block])
                    score = self.best_score(rolling, sample_len)
                    self.event_queue.put(("score", score))

                    now = time.monotonic()
                    if score >= self.threshold and now - self.last_trigger >= self.cooldown_s:
                        self.last_trigger = now
                        self.event_queue.put(("trigger", score))
        except Exception as exc:
            self.event_queue.put(("error", str(exc)))

    def best_score(self, rolling: np.ndarray, sample_len: int) -> float:
        best = 0.0
        offsets = (0, BLOCK_SIZE // 2, BLOCK_SIZE, BLOCK_SIZE + BLOCK_SIZE // 2)
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
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Audio Macro Trigger")
        self.root.geometry("1080x720")
        self.root.minsize(980, 640)

        self.devices: list[sc.Microphone] = []
        self.sample = np.array([], dtype=np.float32)
        self.detector: AudioDetector | None = None
        self.stop_event = threading.Event()
        self.recording = False
        self.playing_sample = False
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
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#f6f7f9")
        style.configure("TLabelframe", background="#f6f7f9", borderwidth=1, relief="solid")
        style.configure("TLabelframe.Label", background="#f6f7f9", foreground="#20242a", font=("Segoe UI", 10, "bold"))
        style.configure("TLabel", background="#f6f7f9", foreground="#20242a", font=("Segoe UI", 9))
        style.configure("Muted.TLabel", background="#f6f7f9", foreground="#5f6670")
        style.configure("Status.TLabel", background="#eceff3", foreground="#20242a", padding=(8, 4))
        style.configure("TButton", font=("Segoe UI", 9), padding=(10, 5))
        style.configure("Primary.TButton", font=("Segoe UI", 9, "bold"), padding=(12, 6))

    def build_ui(self) -> None:
        self.root.configure(background="#f6f7f9")

        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill=BOTH, expand=True)
        outer.columnconfigure(0, weight=3)
        outer.columnconfigure(1, weight=2)
        outer.rowconfigure(3, weight=1)

        header = ttk.Frame(outer)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Audio Macro Trigger", font=("Segoe UI", 15, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(header, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=1, sticky="e")

        audio_group = ttk.LabelFrame(outer, text="Audio")
        audio_group.grid(row=1, column=0, sticky="ew", padx=(0, 10), pady=(0, 12))
        audio_group.columnconfigure(0, weight=1)
        audio_group.columnconfigure(1, weight=0)

        ttk.Label(audio_group, text="Audio device").grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 2))
        self.device_combo = ttk.Combobox(audio_group, textvariable=self.device_var, state="readonly")
        self.device_combo.grid(row=1, column=0, sticky="ew", padx=(10, 8), pady=(0, 10))
        ttk.Button(audio_group, text="Refresh", command=self.refresh_devices).grid(row=1, column=1, sticky="e", padx=(0, 10), pady=(0, 10))

        tuning_group = ttk.LabelFrame(outer, text="Detection")
        tuning_group.grid(row=1, column=1, sticky="ew", pady=(0, 12))
        for column in range(3):
            tuning_group.columnconfigure(column, weight=1)
        self.add_spinbox(tuning_group, "Sample seconds", self.sample_seconds_var, 0.5, 10.0, 0.5, 0, 0)
        self.add_spinbox(tuning_group, "Detect seconds", self.detect_seconds_var, 0.1, 5.0, 0.1, 0, 1)
        self.add_spinbox(tuning_group, "Threshold", self.threshold_var, 0.1, 0.95, 0.05, 0, 2)
        self.add_spinbox(tuning_group, "Cooldown seconds", self.cooldown_var, 0.0, 30.0, 0.5, 1, 0)
        self.add_spinbox(tuning_group, "Action jitter ms", self.jitter_var, 0, 500, 5, 1, 1)

        sample_group = ttk.LabelFrame(outer, text="Sample & Actions")
        sample_group.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        for column in range(6):
            sample_group.columnconfigure(column, weight=0)
        sample_group.columnconfigure(6, weight=1)
        self.record_button = ttk.Button(sample_group, text="Record sample", command=self.record_sample)
        self.record_button.grid(row=0, column=0, sticky="w", padx=(10, 6), pady=10)
        self.play_sample_button = ttk.Button(sample_group, text="Play sample", command=self.play_sample, state="disabled")
        self.play_sample_button.grid(row=0, column=1, sticky="w", padx=6, pady=10)
        ttk.Button(sample_group, text="Import sample", command=self.import_sample).grid(row=0, column=2, sticky="w", padx=6, pady=10)
        ttk.Button(sample_group, text="Save profile", command=self.save_profile).grid(row=0, column=3, sticky="w", padx=6, pady=10)
        ttk.Button(sample_group, text="Test macro", command=self.test_macro).grid(row=0, column=4, sticky="w", padx=6, pady=10)
        self.listen_button = ttk.Button(sample_group, text="Start listening", command=self.toggle_listening, style="Primary.TButton")
        self.listen_button.grid(row=0, column=6, sticky="e", padx=10, pady=10)
        self.activity_bar = ttk.Progressbar(sample_group, mode="indeterminate", length=160)
        self.activity_bar.grid(row=1, column=0, columnspan=7, sticky="ew", padx=10, pady=(0, 10))
        if not self.admin:
            ttk.Label(sample_group, text="Warning: not running as Administrator. Elevated apps may ignore macro keys.", style="Muted.TLabel").grid(row=2, column=0, columnspan=7, sticky="w", padx=10, pady=(0, 10))

        macro_group = ttk.LabelFrame(outer, text="Macro Steps")
        macro_group.grid(row=3, column=0, sticky="nsew", padx=(0, 10))
        macro_group.rowconfigure(1, weight=1)
        macro_group.columnconfigure(0, weight=1)
        ttk.Label(macro_group, text="One action per line: action, delay_ms or action, min-max", style="Muted.TLabel").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))
        self.macro_text = Text(macro_group, height=16, wrap="none")
        self.macro_text.configure(font=("Consolas", 10), relief="solid", borderwidth=1, padx=8, pady=8, undo=True)
        self.macro_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.macro_text.insert(END, "right_click, 900-1100\nright_click, 0\n")

        events_group = ttk.LabelFrame(outer, text="Events")
        events_group.grid(row=3, column=1, sticky="nsew")
        events_group.rowconfigure(1, weight=1)
        events_group.columnconfigure(0, weight=1)
        ttk.Label(events_group, textvariable=self.score_var, style="Muted.TLabel").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))
        self.event_list = Listbox(events_group, height=16)
        self.event_list.configure(font=("Consolas", 9), relief="solid", borderwidth=1, activestyle="none")
        self.event_list.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

    @staticmethod
    def add_spinbox(parent: Frame, label: str, variable, from_: float, to: float, increment: float, row: int, column: int) -> None:
        box = ttk.Frame(parent)
        box.grid(row=row, column=column, sticky="ew", padx=10, pady=(10, 8))
        box.columnconfigure(0, weight=1)
        ttk.Label(box, text=label).grid(row=0, column=0, sticky="w")
        ttk.Spinbox(box, textvariable=variable, from_=from_, to=to, increment=increment, width=12).grid(row=1, column=0, sticky="ew", pady=(4, 0))

    def refresh_devices(self) -> None:
        self.devices = list(sc.all_microphones(include_loopback=True))
        names = [device.name for device in self.devices]
        self.device_combo["values"] = names
        loopback = next((name for name in names if "loopback" in name.lower()), names[0] if names else "")
        if loopback and not self.device_var.get():
            self.device_var.set(loopback)
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
        self.activity_bar.start(12)
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
        self.activity_bar.start(12)
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
        self.listen_button.configure(text="Stop listening")
        self.status_var.set("Listening")
        self.log("Listening started")

    def stop_listening(self) -> None:
        self.stop_event.set()
        self.listen_button.configure(text="Start listening")
        self.status_var.set(f"Idle ({'Administrator' if self.admin else 'User'})")
        self.log("Listening stopped")

    def consume_events(self) -> None:
        while True:
            try:
                event, payload = self.events.get_nowait()
            except queue.Empty:
                break

            if event == "score":
                self.score_var.set(f"Score: {float(payload):.3f}")
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

        self.root.after(100, self.consume_events)

    def start_macro_thread(self, steps: list[MacroStep]) -> None:
        jitter_ms = int(self.jitter_var.get())
        threading.Thread(target=self._macro_worker, args=(steps, jitter_ms), daemon=True).start()

    def _macro_worker(self, steps: list[MacroStep], jitter_ms: int) -> None:
        try:
            self.runner.run(steps, jitter_ms)
            self.events.put(("macro_done", None))
        except Exception as exc:
            self.events.put(("error", f"Macro failed: {exc}"))

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
        self.event_list.insert(END, f"{timestamp}  {message}")
        self.event_list.yview_moveto(1)

    def on_close(self) -> None:
        self.stop_event.set()
        self.root.destroy()


def main() -> None:
    root = Tk()
    AudioMacroApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
