"""
Push-to-Talk voice-to-text with UI.
Hold a hotkey to record, release to transcribe and paste.
"""

import ctypes
import io
import json
import os
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk
import wave
import winsound

import keyboard
import pyaudio
import pyperclip
from faster_whisper import WhisperModel

# --- Config ---
SAMPLE_RATE = 16000
CHANNELS = 1
CHUNK = 1024
WHISPER_MODELS = {"Fast (base.en)": "base.en", "Accurate (small.en)": "small.en"}
DEFAULT_MODEL_LABEL = "Fast (base.en)"
MIN_DURATION_SEC = 0.3
DEFAULT_HOTKEY = "right ctrl"
CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def load_config() -> dict:
    try:
        with open(CONFIG_PATH, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_config(cfg: dict):
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=2)


user32 = ctypes.windll.user32


def find_claude_window():
    """Find the window titled exactly 'CLAUDE'."""
    hwnd = user32.FindWindowW(None, "CLAUDE")
    return hwnd if hwnd else None


def get_foreground_window():
    return user32.GetForegroundWindow()


import subprocess

SW_RESTORE = 9


def set_foreground_window(hwnd):
    """Force window to foreground via PowerShell."""
    user32.ShowWindow(hwnd, SW_RESTORE)
    # PowerShell's AppActivate reliably steals focus
    subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         f'(New-Object -ComObject WScript.Shell).AppActivate({hwnd})'],
        capture_output=True, timeout=3
    )


class PushToTalk:
    def __init__(self):
        self._models = {}
        self._active_model = DEFAULT_MODEL_LABEL
        self.audio = None
        self.recording = False
        self._lock = threading.Lock()
        self._stop_event = threading.Event()
        self._frames = []
        self._hotkey = DEFAULT_HOTKEY
        self._hook = None

        # UI options
        self.send_enter = False
        self.switch_window = False

        # UI callback
        self._on_status = None

    def load_models(self):
        for label, model_id in WHISPER_MODELS.items():
            self._set_status(f"Loading {label}...")
            self._models[label] = WhisperModel(model_id, device="cpu", compute_type="int8")
        self.audio = pyaudio.PyAudio()
        self._register_hotkey()
        self._set_status("Ready")

    def set_model(self, label: str):
        self._active_model = label

    @property
    def whisper(self):
        return self._models.get(self._active_model)

    def _set_status(self, text: str):
        if self._on_status:
            self._on_status(text)

    def _register_hotkey(self):
        self._unregister_hotkey()

        def handler(event):
            if event.name == self._hotkey:
                if event.event_type == "down":
                    self._on_press(event)
                elif event.event_type == "up":
                    self._on_release(event)

        self._hook = keyboard.hook(handler, suppress=False)

    def _unregister_hotkey(self):
        if self._hook:
            keyboard.unhook(self._hook)
            self._hook = None

    def set_hotkey(self, key: str):
        self._hotkey = key
        if self.whisper:
            self._register_hotkey()

    def _on_press(self, _event):
        with self._lock:
            if self.recording:
                return
            self.recording = True
        self._frames = []
        self._stop_event.clear()
        t = threading.Thread(target=self._record_loop, daemon=True)
        t.start()
        winsound.Beep(800, 100)
        self._set_status("Recording...")

    def _on_release(self, _event):
        with self._lock:
            if not self.recording:
                return
            self.recording = False
        self._stop_event.set()

    def _record_loop(self):
        stream = self.audio.open(
            format=pyaudio.paInt16,
            channels=CHANNELS,
            rate=SAMPLE_RATE,
            input=True,
            frames_per_buffer=CHUNK,
        )
        try:
            while not self._stop_event.is_set():
                data = stream.read(CHUNK, exception_on_overflow=False)
                self._frames.append(data)
        finally:
            stream.stop_stream()
            stream.close()

        winsound.Beep(400, 100)

        audio_data = b"".join(self._frames)
        min_bytes = int(SAMPLE_RATE * 2 * MIN_DURATION_SEC)
        if len(audio_data) < min_bytes:
            self._set_status("Too short, skipped")
            return

        self._set_status("Transcribing...")
        text = self._transcribe(audio_data)
        if text:
            self._deliver(text)
            self._set_status(f"OK: {text[:60]}")
        else:
            self._set_status("No speech detected")

    def _transcribe(self, audio_data: bytes) -> str:
        buf = io.BytesIO()
        with wave.open(buf, "wb") as wf:
            wf.setnchannels(CHANNELS)
            wf.setsampwidth(2)
            wf.setframerate(SAMPLE_RATE)
            wf.writeframes(audio_data)
        buf.seek(0)
        segments, _ = self.whisper.transcribe(buf, language="en", beam_size=1)
        return " ".join(seg.text for seg in segments).strip()

    def _deliver(self, text: str):
        """Paste text into the target window, with or without focus."""
        WM_PASTE = 0x0302
        WM_KEYDOWN = 0x0100
        WM_KEYUP = 0x0101
        VK_RETURN = 0x0D

        old_clip = ""
        try:
            old_clip = pyperclip.paste()
        except Exception:
            pass
        pyperclip.copy(text)
        time.sleep(0.05)

        if self.switch_window:
            claude_hwnd = find_claude_window()
            if claude_hwnd:
                # Send paste directly to the Claude window — no focus needed
                user32.SendMessageW(claude_hwnd, WM_PASTE, 0, 0)
                if self.send_enter:
                    time.sleep(0.05)
                    user32.PostMessageW(claude_hwnd, WM_KEYDOWN, VK_RETURN, 0)
                    user32.PostMessageW(claude_hwnd, WM_KEYUP, VK_RETURN, 0)
            else:
                # Fallback: paste into active window
                keyboard.send("ctrl+v")
                time.sleep(0.15)
                if self.send_enter:
                    keyboard.send("enter")
        else:
            keyboard.send("ctrl+v")
            time.sleep(0.15)
            if self.send_enter:
                keyboard.send("enter")

        time.sleep(0.1)
        try:
            pyperclip.copy(old_clip)
        except Exception:
            pass

    def shutdown(self):
        self._unregister_hotkey()
        if self.audio:
            self.audio.terminate()


class PttApp:
    def __init__(self):
        self.ptt = PushToTalk()
        self._cfg = load_config()
        self.root = tk.Tk()
        self.root.title("Push-to-Talk")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
        self._apply_config()

        # Thread-safe status updates
        self._pending_status = None
        self.ptt._on_status = self._queue_status

        # Load model in background
        threading.Thread(target=self._init_model, daemon=True).start()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}
        frame = ttk.Frame(self.root, padding=12)
        frame.pack()

        # Status
        self._status_var = tk.StringVar(value="Loading...")
        status_label = ttk.Label(frame, textvariable=self._status_var, font=("Segoe UI", 11, "bold"))
        status_label.pack(**pad)

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Hotkey selector
        hotkey_frame = ttk.Frame(frame)
        hotkey_frame.pack(fill="x", **pad)
        ttk.Label(hotkey_frame, text="Hotkey:").pack(side="left")
        self._hotkey_var = tk.StringVar(value=DEFAULT_HOTKEY)
        self._hotkey_btn = ttk.Button(hotkey_frame, textvariable=self._hotkey_var,
                                       command=self._start_hotkey_capture, width=20)
        self._hotkey_btn.pack(side="right")

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Model selector
        model_frame = ttk.Frame(frame)
        model_frame.pack(fill="x", **pad)
        ttk.Label(model_frame, text="Model:").pack(side="left")
        self._model_var = tk.StringVar(value=DEFAULT_MODEL_LABEL)
        for label in WHISPER_MODELS:
            rb = ttk.Radiobutton(model_frame, text=label, variable=self._model_var,
                                  value=label, command=self._on_model_change)
            rb.pack(side="right", padx=(8, 0))

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Checkboxes
        self._enter_var = tk.BooleanVar(value=False)
        enter_cb = ttk.Checkbutton(frame, text="Send Enter after paste",
                                    variable=self._enter_var, command=self._on_enter_toggle)
        enter_cb.pack(anchor="w", **pad)

        self._switch_var = tk.BooleanVar(value=False)
        switch_cb = ttk.Checkbutton(frame, text="Switch to Claude window and back",
                                     variable=self._switch_var, command=self._on_switch_toggle)
        switch_cb.pack(anchor="w", **pad)

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Reload button
        ttk.Button(frame, text="Reload", command=self._reload_app).pack(fill="x", **pad)

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Last transcription
        ttk.Label(frame, text="Last:").pack(anchor="w", padx=8)
        self._last_var = tk.StringVar(value="")
        last_label = ttk.Label(frame, textvariable=self._last_var, wraplength=280,
                                foreground="gray")
        last_label.pack(anchor="w", padx=8, pady=(0, 8))

    def _apply_config(self):
        hotkey = self._cfg.get("hotkey", DEFAULT_HOTKEY)
        model = self._cfg.get("model", DEFAULT_MODEL_LABEL)
        send_enter = self._cfg.get("send_enter", False)
        switch_window = self._cfg.get("switch_window", False)

        self._hotkey_var.set(hotkey)
        self.ptt._hotkey = hotkey
        self._model_var.set(model)
        self.ptt._active_model = model
        self._enter_var.set(send_enter)
        self.ptt.send_enter = send_enter
        self._switch_var.set(switch_window)
        self.ptt.switch_window = switch_window

    def _reload_app(self):
        """Kill this process and restart it to pick up code changes."""
        self._save_config()
        self.ptt.shutdown()
        self.root.destroy()
        os.execv(sys.executable, [sys.executable] + sys.argv)

    def _save_config(self):
        self._cfg = {
            "hotkey": self._hotkey_var.get(),
            "model": self._model_var.get(),
            "send_enter": self._enter_var.get(),
            "switch_window": self._switch_var.get(),
        }
        save_config(self._cfg)

    def _init_model(self):
        self.ptt.load_models()

    def _queue_status(self, text: str):
        """Thread-safe status update via tkinter after()."""
        self._pending_status = text
        try:
            self.root.after(0, self._apply_status)
        except Exception:
            pass

    def _apply_status(self):
        if self._pending_status is not None:
            status = self._pending_status
            self._pending_status = None
            if status.startswith("OK: "):
                self._last_var.set(status[4:])
                self._status_var.set("Ready")
            else:
                self._status_var.set(status)

    def _on_model_change(self):
        self.ptt.set_model(self._model_var.get())
        self._save_config()

    def _on_enter_toggle(self):
        self.ptt.send_enter = self._enter_var.get()
        self._save_config()

    def _on_switch_toggle(self):
        self.ptt.switch_window = self._switch_var.get()
        self._save_config()

    def _start_hotkey_capture(self):
        self._hotkey_btn.configure(state="disabled")
        self._hotkey_var.set("Press a key...")
        self.ptt._unregister_hotkey()

        def wait_for_key():
            time.sleep(0.3)  # ignore the click that triggered this
            key = keyboard.read_key(suppress=False)
            self.root.after(0, lambda: self._finish_hotkey_capture(key))

        threading.Thread(target=wait_for_key, daemon=True).start()

    def _finish_hotkey_capture(self, key: str):
        self._hotkey_var.set(key)
        self._hotkey_btn.configure(state="normal")
        self.ptt.set_hotkey(key)
        self._save_config()

    def _on_close(self):
        self.ptt.shutdown()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = PttApp()
    app.run()
