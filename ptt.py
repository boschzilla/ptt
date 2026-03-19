"""
Push-to-Talk voice-to-text with UI.
Hold a hotkey to record, release to transcribe and paste.
"""

import ctypes
import io
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

user32 = ctypes.windll.user32


def find_claude_window():
    """Find a window with 'Claude' in the title."""
    result = []

    def enum_cb(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(256)
            user32.GetWindowTextW(hwnd, buf, 256)
            title = buf.value
            if title and "claude" in title.lower():
                result.append(hwnd)
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
    user32.EnumWindows(WNDENUMPROC(enum_cb), 0)
    return result[0] if result else None


def get_foreground_window():
    return user32.GetForegroundWindow()


def set_foreground_window(hwnd):
    """Bring window to foreground with Alt-key trick."""
    user32.keybd_event(0x12, 0, 0, 0)       # Alt down
    user32.SetForegroundWindow(hwnd)
    user32.keybd_event(0x12, 0, 0x0002, 0)  # Alt up


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
        self._hooks = []

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
        self._hooks.append(keyboard.on_press_key(self._hotkey, self._on_press, suppress=False))
        self._hooks.append(keyboard.on_release_key(self._hotkey, self._on_release, suppress=False))

    def _unregister_hotkey(self):
        for hook in self._hooks:
            keyboard.unhook(hook)
        self._hooks.clear()

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
        """Paste text, optionally switching windows and sending Enter."""
        prev_hwnd = None

        if self.switch_window:
            prev_hwnd = get_foreground_window()
            claude_hwnd = find_claude_window()
            if claude_hwnd:
                set_foreground_window(claude_hwnd)
                time.sleep(0.3)

        # Paste via clipboard
        old_clip = ""
        try:
            old_clip = pyperclip.paste()
        except Exception:
            pass
        pyperclip.copy(text)
        time.sleep(0.05)
        keyboard.send("ctrl+v")
        time.sleep(0.15)

        if self.send_enter:
            keyboard.send("enter")
            time.sleep(0.1)

        try:
            pyperclip.copy(old_clip)
        except Exception:
            pass

        if self.switch_window and prev_hwnd:
            time.sleep(0.2)
            set_foreground_window(prev_hwnd)

    def shutdown(self):
        self._unregister_hotkey()
        if self.audio:
            self.audio.terminate()


class PttApp:
    def __init__(self):
        self.ptt = PushToTalk()
        self.root = tk.Tk()
        self.root.title("Push-to-Talk")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()

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

        # Last transcription
        ttk.Label(frame, text="Last:").pack(anchor="w", padx=8)
        self._last_var = tk.StringVar(value="")
        last_label = ttk.Label(frame, textvariable=self._last_var, wraplength=280,
                                foreground="gray")
        last_label.pack(anchor="w", padx=8, pady=(0, 8))

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

    def _on_enter_toggle(self):
        self.ptt.send_enter = self._enter_var.get()

    def _on_switch_toggle(self):
        self.ptt.switch_window = self._switch_var.get()

    def _start_hotkey_capture(self):
        self._hotkey_btn.configure(state="disabled")
        self._hotkey_var.set("Press a key...")
        self.ptt._unregister_hotkey()

        def on_key(event):
            key = event.name
            keyboard.unhook(hook)
            self.root.after(0, lambda: self._finish_hotkey_capture(key))

        hook = keyboard.on_press(on_key, suppress=False)

    def _finish_hotkey_capture(self, key: str):
        self._hotkey_var.set(key)
        self._hotkey_btn.configure(state="normal")
        self.ptt.set_hotkey(key)

    def _on_close(self):
        self.ptt.shutdown()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = PttApp()
    app.run()
