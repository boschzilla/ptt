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
import numpy as np
import sounddevice as sd
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


def get_foreground_window():
    return user32.GetForegroundWindow()


SW_RESTORE = 9


IsIconic = user32.IsIconic  # returns nonzero if window is minimized


GetCurrentThreadId = ctypes.windll.kernel32.GetCurrentThreadId
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
AttachThreadInput = user32.AttachThreadInput


SPI_GETFOREGROUNDLOCKTIMEOUT = 0x2000
SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001
SPIF_SENDCHANGE = 0x0002


def set_foreground_window(hwnd):
    """Reliably steal focus by temporarily zeroing the foreground lock timeout."""
    if IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)
    fg_hwnd = user32.GetForegroundWindow()
    fg_tid = GetWindowThreadProcessId(fg_hwnd, None)
    my_tid = GetCurrentThreadId()
    attached = False
    if fg_tid and fg_tid != my_tid:
        AttachThreadInput(my_tid, fg_tid, True)
        attached = True
    # Temporarily disable the foreground lock so SetForegroundWindow works
    # from a background thread — no Alt tap needed (which can trigger Task Switcher).
    old_timeout = ctypes.c_uint(0)
    user32.SystemParametersInfoW(
        SPI_GETFOREGROUNDLOCKTIMEOUT, 0, ctypes.byref(old_timeout), 0
    )
    user32.SystemParametersInfoW(
        SPI_SETFOREGROUNDLOCKTIMEOUT, 0, None, SPIF_SENDCHANGE
    )
    user32.SetForegroundWindow(hwnd)
    user32.BringWindowToTop(hwnd)
    user32.SystemParametersInfoW(
        SPI_SETFOREGROUNDLOCKTIMEOUT, 0, ctypes.c_void_p(old_timeout.value), SPIF_SENDCHANGE
    )
    if attached:
        AttachThreadInput(my_tid, fg_tid, False)


class PushToTalk:
    def __init__(self):
        self._models = {}
        self._active_model = DEFAULT_MODEL_LABEL
        self._stream = None
        self.recording = False
        self._lock = threading.Lock()
        self._stop_event = threading.Event()
        self._frames = []
        self._hotkey = DEFAULT_HOTKEY
        self._hook = None

        # UI options
        self.send_enter = False
        self.switch_window = False

        # Window that was active when hotkey was pressed
        self._target_hwnd = None

        # Last successful transcription for repeat
        self._last_text = None

        # UI callback
        self._on_status = None

    def load_models(self):
        for label, model_id in WHISPER_MODELS.items():
            self._set_status(f"Loading {label}...")
            self._models[label] = WhisperModel(
                model_id, device="cpu", compute_type="int8"
            )
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
        # Capture the window that's currently focused — this is where text will go
        self._target_hwnd = get_foreground_window()
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
        def _callback(indata, frames, time_info, status):
            self._frames.append(indata.copy())

        self._stream = sd.InputStream(
            samplerate=SAMPLE_RATE,
            channels=CHANNELS,
            dtype="int16",
            blocksize=CHUNK,
            callback=_callback,
        )
        self._stream.start()
        try:
            while not self._stop_event.is_set():
                time.sleep(0.01)
        finally:
            self._stream.stop()
            self._stream.close()
            self._stream = None

        winsound.Beep(400, 100)

        audio_data = np.concatenate(self._frames).tobytes()
        min_bytes = int(SAMPLE_RATE * 2 * MIN_DURATION_SEC)
        if len(audio_data) < min_bytes:
            self._set_status("Too short, skipped")
            return

        self._set_status("Transcribing...")
        text = self._transcribe(audio_data)
        if text:
            self._last_text = text
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

    @staticmethod
    def _get_window_title(hwnd) -> str:
        buf = ctypes.create_unicode_buffer(512)
        user32.GetWindowTextW(hwnd, buf, 512)
        return buf.value

    @staticmethod
    def _get_window_class(hwnd) -> str:
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, buf, 256)
        return buf.value

    def _is_mintty(self, hwnd) -> bool:
        return self._get_window_class(hwnd) == "mintty"

    def _is_wezterm(self, hwnd) -> bool:
        return self._get_window_class(hwnd) == "org.wezfurlong.wezterm"

    def _is_electron(self, hwnd) -> bool:
        return self._get_window_class(hwnd) == "Chrome_WidgetWin_1"

    def _is_cursor(self, hwnd) -> bool:
        return self._is_electron(hwnd) and "Cursor" in self._get_window_title(hwnd)

    def _deliver(self, text: str):
        """Restore focus to the window captured at press-time, then paste."""
        hwnd = self._target_hwnd
        ptt_hwnd = user32.FindWindowW(None, "Push-to-Talk")
        is_cursor = hwnd and self._is_cursor(hwnd)
        if hwnd and hwnd != ptt_hwnd:
            set_foreground_window(hwnd)
            if is_cursor:
                time.sleep(0.5)
            elif self._is_electron(hwnd):
                time.sleep(0.35)
            else:
                time.sleep(0.15)

        if hwnd and self._is_mintty(hwnd):
            # mintty: type characters directly (ctrl+v and shift+insert don't work)
            keyboard.write(text, delay=0.01)
        else:
            # Clipboard paste — WezTerm uses ctrl+shift+v, everything else ctrl+v
            if hwnd and (self._is_wezterm(hwnd) or is_cursor):
                paste_key = "ctrl+shift+v"
            else:
                paste_key = "ctrl+v"
            old_clip = ""
            try:
                old_clip = pyperclip.paste()
            except Exception:
                pass
            pyperclip.copy(text)
            # Verify clipboard took the value before pasting
            time.sleep(0.05)
            for _ in range(5):
                try:
                    if pyperclip.paste() == text:
                        break
                except Exception:
                    pass
                time.sleep(0.02)
            keyboard.send(paste_key)
            # Wait for app to process the paste before restoring clipboard
            if is_cursor:
                time.sleep(0.7)
            elif hwnd and self._is_electron(hwnd):
                time.sleep(0.5)
            else:
                time.sleep(0.2)
            try:
                pyperclip.copy(old_clip)
            except Exception:
                pass

        if self.send_enter:
            time.sleep(0.05)
            keyboard.send("enter")

        # Optionally return focus to the PTT window
        if self.switch_window:
            ptt_hwnd = user32.FindWindowW(None, "Push-to-Talk")
            if ptt_hwnd:
                time.sleep(0.05)
                set_foreground_window(ptt_hwnd)

    def repeat(self, countdown_cb=None):
        """Re-deliver the last transcription after a 3-second countdown."""
        if not self._last_text:
            self._set_status("Nothing to repeat")
            return
        text = self._last_text
        for i in range(3, 0, -1):
            self._set_status(f"Repeat in {i}...")
            if countdown_cb:
                countdown_cb(i)
            time.sleep(1)
        self._target_hwnd = get_foreground_window()
        self._deliver(text)
        self._set_status(f"OK: {text[:60]}")

    def shutdown(self):
        self._unregister_hotkey()
        if self._stream:
            self._stream.stop()
            self._stream.close()


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
        status_label = ttk.Label(
            frame, textvariable=self._status_var, font=("Segoe UI", 11, "bold")
        )
        status_label.pack(**pad)

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Hotkey selector
        hotkey_frame = ttk.Frame(frame)
        hotkey_frame.pack(fill="x", **pad)
        ttk.Label(hotkey_frame, text="Hotkey:").pack(side="left")
        self._hotkey_var = tk.StringVar(value=DEFAULT_HOTKEY)
        self._hotkey_btn = ttk.Button(
            hotkey_frame,
            textvariable=self._hotkey_var,
            command=self._start_hotkey_capture,
            width=20,
        )
        self._hotkey_btn.pack(side="right")

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Model selector
        model_frame = ttk.Frame(frame)
        model_frame.pack(fill="x", **pad)
        ttk.Label(model_frame, text="Model:").pack(side="left")
        self._model_var = tk.StringVar(value=DEFAULT_MODEL_LABEL)
        for label in WHISPER_MODELS:
            rb = ttk.Radiobutton(
                model_frame,
                text=label,
                variable=self._model_var,
                value=label,
                command=self._on_model_change,
            )
            rb.pack(side="right", padx=(8, 0))

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Checkboxes
        self._enter_var = tk.BooleanVar(value=False)
        enter_cb = ttk.Checkbutton(
            frame,
            text="Send Enter after paste",
            variable=self._enter_var,
            command=self._on_enter_toggle,
        )
        enter_cb.pack(anchor="w", **pad)

        self._switch_var = tk.BooleanVar(value=False)
        switch_cb = ttk.Checkbutton(
            frame,
            text="Return focus here after paste",
            variable=self._switch_var,
            command=self._on_switch_toggle,
        )
        switch_cb.pack(anchor="w", **pad)

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Buttons row
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", **pad)
        ttk.Button(btn_frame, text="Reload", command=self._reload_app).pack(
            side="left", expand=True, fill="x", padx=(0, 4)
        )
        self._repeat_btn = ttk.Button(
            btn_frame, text="Repeat", command=self._on_repeat
        )
        self._repeat_btn.pack(side="right", expand=True, fill="x", padx=(4, 0))

        ttk.Separator(frame, orient="horizontal").pack(fill="x", **pad)

        # Last transcription
        ttk.Label(frame, text="Last:").pack(anchor="w", padx=8)
        self._last_var = tk.StringVar(value="")
        last_label = ttk.Label(
            frame, textvariable=self._last_var, wraplength=280, foreground="gray"
        )
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

    def _on_repeat(self):
        self._repeat_btn.configure(state="disabled")
        threading.Thread(target=self._repeat_thread, daemon=True).start()

    def _repeat_thread(self):
        self.ptt.repeat()
        try:
            self.root.after(0, lambda: self._repeat_btn.configure(state="normal"))
        except Exception:
            pass

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
