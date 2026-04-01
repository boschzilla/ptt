"""
Iterate over delivery methods for mintty (openclaw-tui).
Run this while openclaw-tui is open and watch which method types into it.

Usage:
    python test_deliver.py
"""
import ctypes
import time
import sys

import keyboard
import pyperclip

user32 = ctypes.windll.user32

TEXT = "hello from ptt test"

KEYEVENTF_KEYUP = 0x0002
VK_MENU   = 0x12
VK_SHIFT  = 0x10
VK_INSERT = 0x2D
WM_CHAR   = 0x0102
WM_PASTE  = 0x0302
SW_RESTORE = 9


def find_mintty():
    found = []
    def cb(hwnd, _):
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, buf, 256)
        if buf.value == "mintty":
            title = ctypes.create_unicode_buffer(256)
            user32.GetWindowTextW(hwnd, title, 256)
            found.append((hwnd, title.value))
        return True
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_size_t, ctypes.c_size_t)
    user32.EnumWindows(WNDENUMPROC(cb), 0)
    return found


def focus(hwnd):
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)
    user32.keybd_event(VK_MENU, 0, 0, 0)
    user32.SetForegroundWindow(hwnd)
    user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
    time.sleep(0.25)
    actual = user32.GetForegroundWindow()
    print(f"  focus ok: {actual == hwnd}  (wanted {hwnd}, got {actual})")


def method_shift_insert(hwnd):
    pyperclip.copy(TEXT)
    time.sleep(0.05)
    focus(hwnd)
    keyboard.send("shift+insert")
    time.sleep(0.3)


def method_keybd_event(hwnd):
    pyperclip.copy(TEXT)
    time.sleep(0.05)
    focus(hwnd)
    user32.keybd_event(VK_SHIFT, 0, 0, 0)
    time.sleep(0.02)
    user32.keybd_event(VK_INSERT, 0, 0, 0)
    time.sleep(0.02)
    user32.keybd_event(VK_INSERT, 0, KEYEVENTF_KEYUP, 0)
    time.sleep(0.02)
    user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
    time.sleep(0.3)


def method_keyboard_write(hwnd):
    focus(hwnd)
    keyboard.write(TEXT, delay=0.01)
    time.sleep(0.3)


def method_wm_char(hwnd):
    """Send WM_CHAR directly — no focus needed."""
    for ch in TEXT:
        user32.PostMessageW(hwnd, WM_CHAR, ord(ch), 1)
        time.sleep(0.005)
    time.sleep(0.3)


def method_wm_paste(hwnd):
    """Send WM_PASTE directly — no focus needed."""
    pyperclip.copy(TEXT)
    time.sleep(0.05)
    user32.SendMessageW(hwnd, WM_PASTE, 0, 0)
    time.sleep(0.3)


METHODS = [
    ("shift+insert via keyboard lib", method_shift_insert),
    ("shift+insert via keybd_event",  method_keybd_event),
    ("keyboard.write char-by-char",   method_keyboard_write),
    ("WM_CHAR direct (no focus)",     method_wm_char),
    ("WM_PASTE direct (no focus)",    method_wm_paste),
]


def main():
    windows = find_mintty()
    if not windows:
        print("No mintty window found — open openclaw-tui first.")
        sys.exit(1)

    hwnd, title = windows[0]
    print(f"Found mintty: hwnd={hwnd}  title={title!r}\n")

    for name, fn in METHODS:
        input(f"Press Enter to try: {name!r}  (watch the mintty window)")
        print(f"  -> running {name}")
        fn(hwnd)
        print()


if __name__ == "__main__":
    main()
