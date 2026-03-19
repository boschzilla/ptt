# Push-to-Talk Voice-to-Text

Hold a key to record from your microphone, release to transcribe and paste into the active window. Built for dictating into Claude Code.

## Setup

```
pip install -r requirements.txt
```

## Usage

```
python ptt.py
```

A small always-on-top window appears with these controls:

### Hotkey
Click the hotkey button, then press any key to set it as your push-to-talk trigger. Default: **Right Ctrl**.

### Model
- **Fast (base.en)** — ~0.5s transcription, lower accuracy
- **Accurate (small.en)** — ~1-2s transcription, better accuracy

Both models load at startup so switching is instant.

### Options
- **Send Enter after paste** — automatically submits after pasting (useful for sending to Claude Code)
- **Switch to Claude window and back** — focuses the Claude Code window before pasting, then returns focus to your previous window (finds any window with "Claude" in the title)

### Audio feedback
- High beep (800Hz) — recording started
- Low beep (400Hz) — recording stopped, transcribing

## Dependencies

- `keyboard` — global hotkey hooks
- `PyAudio` — microphone capture
- `faster-whisper` — local Whisper transcription (base.en + small.en)
- `pyperclip` — clipboard paste
- `numpy` — audio processing

## Credits

Inspired by Josh.
