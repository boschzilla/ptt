"""
Test sending text to a window without stealing focus.

1. Lists visible windows — pick one by number
2. Move your focus somewhere else
3. Script waits 3 seconds then fires each method in turn
"""
import ctypes
import time

user32 = ctypes.windll.user32

WM_CHAR    = 0x0102
WM_KEYDOWN = 0x0100
WM_KEYUP   = 0x0101
WM_PASTE   = 0x0302
VK_RETURN  = 0x0D

TEXT = "hello no focus"


# ── Window enumeration ────────────────────────────────────────────────────────

def list_windows():
    windows = []

    def cb(hwnd, _):
        if not user32.IsWindowVisible(hwnd):
            return True
        title_buf = ctypes.create_unicode_buffer(256)
        user32.GetWindowTextW(hwnd, title_buf, 256)
        title = title_buf.value.strip()
        if not title:
            return True
        cls_buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, cls_buf, 256)
        windows.append((hwnd, title, cls_buf.value))
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_size_t, ctypes.c_size_t)
    user32.EnumWindows(WNDENUMPROC(cb), 0)
    return windows


# ── Delivery methods (no focus) ───────────────────────────────────────────────

def try_wm_char(hwnd):
    for ch in TEXT:
        user32.PostMessageW(hwnd, WM_CHAR, ord(ch), 1)
        time.sleep(0.005)


def try_wm_keydown(hwnd):
    for ch in TEXT:
        vk = ctypes.windll.user32.VkKeyScanW(ord(ch)) & 0xFF
        user32.PostMessageW(hwnd, WM_KEYDOWN, vk, 0)
        user32.PostMessageW(hwnd, WM_KEYUP,   vk, 0)
        time.sleep(0.005)


def try_wm_paste(hwnd):
    import pyperclip
    pyperclip.copy(TEXT)
    time.sleep(0.05)
    user32.SendMessageW(hwnd, WM_PASTE, 0, 0)


METHODS = [
    ("WM_CHAR (PostMessage per char)",      try_wm_char),
    ("WM_KEYDOWN/UP (PostMessage per key)", try_wm_keydown),
    ("WM_PASTE (SendMessage)",              try_wm_paste),
]


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    windows = list_windows()

    print("\nVisible windows:\n")
    for i, (hwnd, title, cls) in enumerate(windows):
        print(f"  [{i:2d}]  {title[:60]:<60}  ({cls})")

    print()
    choice = int(input("Pick a window number: "))
    hwnd, title, cls = windows[choice]
    print(f"\nTarget: {title!r}  hwnd={hwnd}  class={cls}")
    print("\nNow move your focus somewhere else.")
    print("Waiting 3 seconds, then testing each method...\n")
    time.sleep(3)

    for name, fn in METHODS:
        current = user32.GetForegroundWindow()
        print(f"[{name}]")
        print(f"  foreground hwnd now: {current}  (target is {hwnd}, match={current == hwnd})")
        fn(hwnd)
        time.sleep(0.5)
        input("  Did text appear? (Enter to continue)")
        print()


if __name__ == "__main__":
    main()
