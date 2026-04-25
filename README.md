# Audio Macro Trigger

Audio Macro Trigger is a Windows desktop app that listens to system audio and runs keyboard or mouse macros when a recorded sound sample is detected.

It is designed for workflows where a specific app or game sound should trigger a short action sequence, such as right-clicking twice with randomized timing.

## Features

- Modern CustomTkinter dashboard with a live detection score meter.
- Long-running detection optimizations: throttled fingerprint matching, throttled UI score updates, and bounded event logs.
- Record sound samples from system audio with WASAPI loopback.
- Play recorded samples back before using them for detection.
- Detect sounds with a spectrum fingerprint, making detection less sensitive to volume changes.
- Trigger keyboard and mouse actions after detection.
- Use fixed delays or randomized delay ranges between actions.
- Add optional action jitter while the macro is running.
- Relaunch as Administrator through `run.ps1`.
- Save local settings in `profile.json`.

## Quick Start

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
.\run.ps1
```

`run.ps1` opens a UAC prompt and runs the app as Administrator. Use this when the target app or game is also running as Administrator.

To run without elevation:

```powershell
python .\src\audio_macro_app.py
```

## Basic Workflow

1. Click **Refresh** and choose the audio device to monitor.
2. Play the sound you want to detect.
3. Click **Record sample**.
4. Wait for `Recording finished` in **Events**.
5. Click **Play sample** to verify the captured audio.
6. Configure **Macro Steps**.
7. Click **Test macro** to verify the input sequence without audio detection.
8. Click **Start listening**.
9. Play the target sound again and watch **Score** and **Events**.

## Macro Syntax

Each macro line uses this format:

```text
action, delay_ms
```

`delay_ms` is the delay after the action. It can be a fixed value:

```text
right_click, 1000
```

or a randomized range:

```text
right_click, 800-1200
```

Example: right-click once, wait roughly one second, then right-click again:

```text
right_click, 900-1100
right_click, 0
```

Supported mouse actions:

```text
left_click
right_click
middle_click
```

Keyboard examples:

```text
f1, 100
space, 250
ctrl+shift+a, 500
enter, 0
```

## Detection Settings

| Setting | Purpose | Suggested start |
| --- | --- | --- |
| `Sample seconds` | How long the app records the sample. | `2.0-3.0` |
| `Detect seconds` | How much of the beginning of the sample is used for matching. Lower values trigger faster. | `0.5-0.8` |
| `Threshold` | Minimum score required to trigger the macro. | `0.30-0.40` |
| `Cooldown seconds` | Minimum time between repeated triggers. | `2.0` |
| `Action jitter ms` | Random internal delay during clicks and key presses. | `30` |

If detection is too slow, lower **Detect seconds**.

If the app does not trigger, lower **Threshold**.

If the app triggers too often, raise **Threshold** or **Detect seconds**.

## Troubleshooting

### The sample sounds correct, but the macro does not trigger

- Click **Start listening** and watch **Score** while the target sound plays.
- Lower **Threshold** to `0.25-0.35`.
- Lower **Detect seconds** only if the score rises late.
- Record a shorter and cleaner sample with less background audio.

### Events shows `Triggered`, but nothing happens

- Click **Test macro**.
- If **Test macro** works on the desktop but not in a game, the game may be blocking simulated input.
- Run the app through `.\run.ps1` and accept the UAC prompt.
- Some anti-cheat systems block normal simulated input.

### The macro triggers too often

- Increase **Cooldown seconds**.
- Increase **Threshold**.
- Record a sample that contains only the unique part of the sound.

### The app feels slow after running for a long time

- Keep **Cooldown seconds** high enough so macros do not queue up repeatedly.
- Avoid very long samples unless needed.
- Use **Detect seconds** around `0.5-0.8` for lower CPU use and faster detection.
- The event log is capped to recent entries, so long sessions should stay responsive.

## Project Structure

```text
.
|-- src/
|   `-- audio_macro_app.py
|-- requirements.txt
|-- run.ps1
|-- README.md
`-- .gitignore
```

Runtime files such as `profile.json` and `sample.npy` are ignored by Git because they are local to your machine and audio device setup.

## Notes

- This app sends standard system input. It does not bypass anti-cheat or input protection.
- For elevated target apps, run this app as Administrator.
- Detection quality depends heavily on the sample. A short, clean, unique sound works best.

## Tech Stack

- Python
- CustomTkinter
- SoundCard
- NumPy
- pynput
