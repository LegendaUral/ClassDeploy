"""
Полноэкранный оверлей блокировки экрана.
Запускается как отдельный процесс из remote.lock_screen() когда агент в сессии 0.
"""
import sys
import os
import threading
import time
import tkinter as tk
from ctypes import *


def _block_taskmgr(stop_signal: str):
    import subprocess
    while True:
        if stop_signal and os.path.exists(stop_signal):
            return
        try:
            r = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq taskmgr.exe"],
                capture_output=True, timeout=2,
            )
            if b"taskmgr.exe" in (r.stdout or b""):
                subprocess.run(
                    ["taskkill", "/IM", "taskmgr.exe", "/F"],
                    capture_output=True, timeout=2,
                )
        except Exception:
            pass
        time.sleep(0.5)


def main():
    message = sys.argv[1] if len(sys.argv) > 1 else "Экран заблокирован"
    stop_signal = sys.argv[2] if len(sys.argv) > 2 else ""

    threading.Thread(target=_block_taskmgr, args=(stop_signal,), daemon=True).start()

    try:
        windll.user32.BlockInput(True)
    except Exception:
        pass

    root = tk.Tk()
    root.title("ClassDeploy Lock")
    root.attributes("-fullscreen", True)
    root.attributes("-topmost", True)
    root.overrideredirect(True)
    root.configure(bg="black")

    def block(e=None):
        return "break"

    for seq in ("<Escape>", "<Alt-F4>", "<Control-c>", "<Control-w>",
                "<Tab>", "<Alt_L>", "<Alt_R>", "<Control-Shift-Escape>", "<Key>"):
        root.bind_all(seq, block)
    root.protocol("WM_DELETE_WINDOW", lambda: None)

    tk.Label(
        root, text=message,
        bg="black", fg="white",
        font=("Segoe UI", 48, "bold"),
        wraplength=root.winfo_screenwidth() - 100,
        justify="center",
    ).place(relx=0.5, rely=0.38, anchor="center")

    tk.Label(
        root, text="Экран заблокирован учителем",
        bg="black", fg="#555",
        font=("Segoe UI", 18),
    ).place(relx=0.5, rely=0.54, anchor="center")

    def close(*_):
        try:
            root.destroy()
        except Exception:
            pass

    def grab():
        if stop_signal and os.path.exists(stop_signal):
            close()
            return
        try:
            root.attributes("-topmost", True)
            root.focus_force()
            root.lift()
            root.grab_set_global()
        except Exception:
            pass
        root.after(250, grab)

    root.after(100, grab)
    try:
        root.mainloop()
    finally:
        try:
            windll.user32.BlockInput(False)
        except Exception:
            pass
        if stop_signal:
            try:
                if os.path.exists(stop_signal):
                    os.unlink(stop_signal)
            except Exception:
                pass


if __name__ == "__main__":
    main()
