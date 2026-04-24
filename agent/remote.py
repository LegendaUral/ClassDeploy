"""
Удалённое управление: блокировка экрана, ввод мыши/клавиатуры,
сообщения, управление питанием, скример.
"""
from __future__ import annotations
import os
import sys
import time
import base64
import ctypes
from ctypes import *
import logging
import subprocess
import tempfile
import threading
from pathlib import Path
from typing import Optional
import uuid
import shlex

log = logging.getLogger("agent.remote")

try:
    from shared.config import AGENT_TEMP_DIR
except Exception:
    AGENT_TEMP_DIR = tempfile.gettempdir()


_screamer_lock = threading.Lock()
_screamer_close: Optional[callable] = None
_temp_lock = threading.Lock()
_temp_files: set[str] = set()


def _register_temp_file(path: str):
    with _temp_lock:
        _temp_files.add(path)


def _cleanup_temp_file(path: str):
    if not path:
        return
    _safe_unlink(path)
    with _temp_lock:
        _temp_files.discard(path)


def _create_temp_script(prefix: str, suffix: str, code: str) -> str:
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=suffix, delete=False,
        encoding="utf-8", prefix=prefix, dir=AGENT_TEMP_DIR if "AGENT_TEMP_DIR" in globals() else None,
    )
    tmp.write(code)
    tmp.close()
    _register_temp_file(tmp.name)
    return tmp.name


def _close_screamer():
    global _screamer_close
    closer = None
    with _screamer_lock:
        closer = _screamer_close
        _screamer_close = None
    if closer:
        try:
            closer()
        except Exception:
            pass

# ════════════════════════════════════════════════════════
#  Удалённый ввод (мышь + клавиатура)
# ════════════════════════════════════════════════════════

class RemoteInput:
    def __init__(self):
        self._ok = False
        try:
            import win32api, win32con
            self._api = win32api
            self._con = win32con
            self._ok = True
        except ImportError:
            log.warning("pywin32 не установлен — удалённый ввод недоступен")

    def _mouse_flags(self, button: str):
        return {
            "left":   (self._con.MOUSEEVENTF_LEFTDOWN,   self._con.MOUSEEVENTF_LEFTUP),
            "right":  (self._con.MOUSEEVENTF_RIGHTDOWN,  self._con.MOUSEEVENTF_RIGHTUP),
            "middle": (self._con.MOUSEEVENTF_MIDDLEDOWN, self._con.MOUSEEVENTF_MIDDLEUP),
        }.get(button, (self._con.MOUSEEVENTF_LEFTDOWN, self._con.MOUSEEVENTF_LEFTUP))

    # ── Мышь ──

    def move_mouse(self, x: int, y: int, sw: int, sh: int):
        if not self._ok:
            return
        try:
            u32 = ctypes.windll.user32
            rw = u32.GetSystemMetrics(0)
            rh = u32.GetSystemMetrics(1)
            self._api.SetCursorPos((int(x * rw / max(1, sw)),
                                    int(y * rh / max(1, sh))))
        except Exception as e:
            log.debug("move_mouse: %s", e)

    def mouse_button(self, x: int, y: int, sw: int, sh: int,
                     button: str = "left", down: bool = True):
        if not self._ok:
            return
        try:
            self.move_mouse(x, y, sw, sh)
            dn, up = self._mouse_flags(button)
            self._api.mouse_event(dn if down else up, 0, 0, 0, 0)
        except Exception as e:
            log.debug("mouse_button: %s", e)

    def click(self, x: int, y: int, sw: int, sh: int,
               button: str = "left", double: bool = False):
        if not self._ok:
            return
        try:
            self.move_mouse(x, y, sw, sh)
            dn, up = self._mouse_flags(button)
            times = 2 if double else 1
            for _ in range(times):
                self._api.mouse_event(dn, 0, 0, 0, 0)
                time.sleep(0.02)
                self._api.mouse_event(up, 0, 0, 0, 0)
                if double:
                    time.sleep(0.05)
        except Exception as e:
            log.debug("click: %s", e)

    def scroll(self, x: int, y: int, sw: int, sh: int, delta: int):
        if not self._ok:
            return
        try:
            self.move_mouse(x, y, sw, sh)
            self._api.mouse_event(self._con.MOUSEEVENTF_WHEEL, 0, 0, delta, 0)
        except Exception as e:
            log.debug("scroll: %s", e)

    # ── Клавиатура ──

    def key(self, vk: int, down: bool = True):
        if not self._ok:
            return
        try:
            flag = 0 if down else self._con.KEYEVENTF_KEYUP
            self._api.keybd_event(vk, 0, flag, 0)
        except Exception as e:
            log.debug("key: %s", e)

    def key_combination(self, *vk_codes: int):
        """Нажать комбинацию клавиш (все down, затем все up в обратном порядке)."""
        for vk in vk_codes:
            self.key(vk, True)
            time.sleep(0.02)
        for vk in reversed(vk_codes):
            self.key(vk, False)
            time.sleep(0.02)

    def type_text(self, text: str):
        if not self._ok:
            return
        try:
            for ch in text:
                self._send_unicode(ch)
                time.sleep(0.01)
        except Exception as e:
            log.debug("type_text: %s", e)

    def _send_unicode(self, ch: str):
        INPUT_KEYBOARD    = 1
        KEYEVENTF_UNICODE = 0x0004
        KEYEVENTF_KEYUP   = 0x0002

        class KEYBDINPUT(ctypes.Structure):
            _fields_ = [
                ("wVk",        ctypes.c_ushort),
                ("wScan",      ctypes.c_ushort),
                ("dwFlags",    ctypes.c_ulong),
                ("time",       ctypes.c_ulong),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
            ]

        class INPUT_I(ctypes.Union):
            _fields_ = [("ki", KEYBDINPUT), ("padding", ctypes.c_byte * 24)]

        class INPUT(ctypes.Structure):
            _fields_ = [("type", ctypes.c_ulong), ("ii", INPUT_I)]

        code = ord(ch)
        dn = INPUT(type=INPUT_KEYBOARD,
                   ii=INPUT_I(ki=KEYBDINPUT(0, code, KEYEVENTF_UNICODE, 0, None)))
        up = INPUT(type=INPUT_KEYBOARD,
                   ii=INPUT_I(ki=KEYBDINPUT(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, None)))
        ctypes.windll.user32.SendInput(1, ctypes.byref(dn), ctypes.sizeof(INPUT))
        ctypes.windll.user32.SendInput(1, ctypes.byref(up), ctypes.sizeof(INPUT))


# ════════════════════════════════════════════════════════
#  Определение сессии Windows
# ════════════════════════════════════════════════════════

def _session_id() -> int:
    try:
        sid = ctypes.c_ulong(0)
        ctypes.windll.kernel32.ProcessIdToSessionId(
            ctypes.windll.kernel32.GetCurrentProcessId(), ctypes.byref(sid))
        return sid.value
    except Exception:
        return -1

def _is_session_zero() -> bool:
    return _session_id() == 0


# ════════════════════════════════════════════════════════
#  Управление чувствительностью мыши
# ════════════════════════════════════════════════════════

def set_mouse_sensitivity(speed: int) -> bool:
    """
    Изменить скорость мыши в Windows.
    speed: 1-20 (1 = медленно, 10 = по умолчанию, 20 = быстро)
    """
    speed = max(1, min(20, int(speed or 10)))
    try:
        import winreg
        key_path = r"Control Panel\Mouse"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, "MouseSensitivity", 0, winreg.REG_SZ, str(speed))
        log.info("Чувствительность мыши установлена: %d", speed)
        return True
    except Exception as e:
        log.warning("set_mouse_sensitivity: %s", e)
        return False


def get_mouse_sensitivity() -> int:
    """Получить текущую чувствительность мыши."""
    try:
        import winreg
        key_path = r"Control Panel\Mouse"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            value, _ = winreg.QueryValueEx(key, "MouseSensitivity")
            return int(value)
    except Exception:
        return 10  # Значение по умолчанию


# ════════════════════════════════════════════════════════
#  Инверсия WASD (переключение режима инверсии)
# ════════════════════════════════════════════════════════

_wasd_inversion_enabled = False
_wasd_inversion_thread: Optional[threading.Thread] = None
_wasd_inversion_stop = threading.Event()

# Коды виртуальных клавиш для WASD
_WASD_KEYS = {
    0x57: "W",  # W
    0x41: "A",  # A  
    0x53: "S",  # S
    0x44: "D",  # D
}

_WASD_INVERSION_MAP = {
    0x57: 0x53,  # W -> S
    0x53: 0x57,  # S -> W
    0x41: 0x44,  # A -> D
    0x44: 0x41,  # D -> A
}


def invert_wasd_key(vk: int) -> int:
    """Инвертировать клавишу WASD или вернуть оригинальную."""
    if _wasd_inversion_enabled and vk in _WASD_INVERSION_MAP:
        return _WASD_INVERSION_MAP[vk]
    return vk


def _wasd_inversion_loop(ri: RemoteInput):
    """
    Поток для отслеживания состояния инверсии WASD.
    Примечание: для полной реализации требуется низкоуровневый keyboard hook.
    Текущая реализация отслеживает состояние и может быть интегрирована
    с INPUT_EVENT сообщениями на уровне server/gui.
    """
    log.info("WASD инверсия: состояние активно")
    while not _wasd_inversion_stop.is_set():
        try:
            time.sleep(0.1)
        except Exception:
            pass
    log.info("WASD инверсия: состояние деактивировано")


def set_wasd_inversion(enabled: bool) -> bool:
    """
    Включить/отключить инверсию WASD.
    Когда включено, W→S, S→W, A→D, D→A
    """
    global _wasd_inversion_enabled, _wasd_inversion_thread
    
    try:
        if enabled and not _wasd_inversion_enabled:
            _wasd_inversion_enabled = True
            _wasd_inversion_stop.clear()
            ri = RemoteInput()
            _wasd_inversion_thread = threading.Thread(
                target=_wasd_inversion_loop, args=(ri,),
                daemon=True, name="WASDInversionThread"
            )
            _wasd_inversion_thread.start()
            log.info("WASD инверсия включена")
            return True
        elif not enabled and _wasd_inversion_enabled:
            _wasd_inversion_enabled = False
            _wasd_inversion_stop.set()
            if _wasd_inversion_thread:
                _wasd_inversion_thread.join(timeout=2)
            log.info("WASD инверсия отключена")
            return True
        return _wasd_inversion_enabled == enabled
    except Exception as e:
        log.warning("set_wasd_inversion: %s", e)
        return False


def toggle_wasd_inversion() -> bool:
    """Переключить режим инверсии WASD."""
    global _wasd_inversion_enabled
    return set_wasd_inversion(not _wasd_inversion_enabled)


# ════════════════════════════════════════════════════════
#  Запуск программы в сессии пользователя
# ════════════════════════════════════════════════════════

def _command_line(args: list[str]) -> str:
    return subprocess.list2cmdline([str(a) for a in args])


def _run_user_session(cmd: list[str] | str) -> bool:
    """Запустить команду в интерактивной сессии пользователя."""
    if isinstance(cmd, str):
        args = [cmd]
    else:
        args = [str(a) for a in cmd if str(a)]
    if not args:
        return False
    if not _is_session_zero():
        try:
            subprocess.Popen(args if len(args) > 1 else args[0], shell=False)
            return True
        except Exception as e:
            log.error("Popen failed: %s", e)
            return False
    return _schtasks_run(args)


def _schtasks_run(cmd: list[str] | str) -> bool:
    task = f"CDA_{int(time.time() * 1000) % 9_999_999}"
    if isinstance(cmd, str):
        args = [cmd]
        command_line = cmd
    else:
        args = [str(a) for a in cmd if str(a)]
        command_line = _command_line(args)
    try:
        subprocess.run(
            ["schtasks", "/create", "/tn", task,
             "/tr", command_line, "/sc", "once", "/st", "23:59",
             "/RL", "HIGHEST", "/RU", "SYSTEM", "/IT", "/F"],
            check=True, capture_output=True, timeout=15,
        )
        subprocess.run(
            ["schtasks", "/run", "/tn", task, "/I"],
            check=True, capture_output=True, timeout=15,
        )
        threading.Timer(20.0, lambda: subprocess.run(
            ["schtasks", "/delete", "/tn", task, "/F"],
            capture_output=True, timeout=10,
        )).start()
        return True
    except Exception as e:
        log.error("schtasks failed: %s", e)
        return False


def _find_python() -> Optional[str]:
    import shutil
    for name in ("python", "python3", "py"):
        p = shutil.which(name)
        if p:
            return p
    for p in (
        r"C:\Python313\python.exe", r"C:\Python312\python.exe",
        r"C:\Python311\python.exe", r"C:\Python310\python.exe",
    ):
        if os.path.exists(p):
            return p
    return None


# ════════════════════════════════════════════════════════
#  Блокировка экрана (inline overlay)
# ════════════════════════════════════════════════════════

_lock_mutex  = threading.Lock()
_lock_event: Optional[threading.Event] = None
_lock_thread: Optional[threading.Thread] = None


def _set_block_input(enabled: bool):
    try:
        windll.user32.BlockInput(bool(enabled))
    except Exception as e:
        log.debug("BlockInput(%s): %s", enabled, e)


def _overlay_thread(message: str, stop_evt: threading.Event):
    """Tkinter overlay в отдельном потоке (не в сессии 0)."""
    try:
        import tkinter as tk
        _set_block_input(True)
        root = tk.Tk()
        root.title("ClassDeploy Lock")
        root.attributes("-fullscreen", True)
        root.attributes("-topmost", True)
        root.overrideredirect(True)
        root.configure(bg="black")

        def block(e=None):
            return "break"

        for seq in ("<Escape>", "<Alt-F4>", "<Control-c>", "<Control-w>",
                    "<Tab>", "<Alt_L>", "<Alt_R>", "<Control-Shift-Escape>",
                    "<Key>"):
            root.bind_all(seq, block)
        root.protocol("WM_DELETE_WINDOW", lambda: None)

        tk.Label(
            root, text=message, bg="black", fg="white",
            font=("Segoe UI", 48, "bold"),
            wraplength=root.winfo_screenwidth() - 100,
            justify="center",
        ).place(relx=0.5, rely=0.38, anchor="center")

        tk.Label(
            root, text="Экран заблокирован учителем",
            bg="black", fg="#555",
            font=("Segoe UI", 18),
        ).place(relx=0.5, rely=0.54, anchor="center")

        def grab():
            if stop_evt.is_set():
                try:
                    root.destroy()
                except Exception:
                    pass
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
            _set_block_input(False)
    except Exception as e:
        _set_block_input(False)
        log.error("overlay_thread: %s", e)


def lock_screen(message: str = "Внимание! Смотрите на доску"):
    global _lock_event, _lock_thread
    with _lock_mutex:
        if _lock_thread and _lock_thread.is_alive():
            return  # уже заблокировано

        stop_evt = threading.Event()
        _lock_event = stop_evt

        if _is_session_zero():
            # Служба: запускаем overlay через Python в сессии пользователя
            _lock_via_session(message, stop_evt)
        else:
            t = threading.Thread(
                target=_overlay_thread, args=(message, stop_evt),
                daemon=True, name="OverlayThread",
            )
            t.start()
            _lock_thread = t
        log.info("Экран заблокирован")


def _lock_via_session(message: str, stop_evt: threading.Event):
    """Запуск overlay через сессию пользователя (для службы)."""
    python = _find_python()
    overlay_py = Path(__file__).parent / "overlay.py"
    signal_path = os.path.join(tempfile.gettempdir(), f"cda_unlock_{uuid.uuid4().hex}.signal")
    setattr(stop_evt, '_signal', signal_path)

    if python and overlay_py.exists():
        cmd = [python, str(overlay_py), message, signal_path]
    elif python:
        tmp_path = _create_temp_script('cda_lock_', '.py', _build_overlay_script(message, signal_path))
        setattr(stop_evt, '_tmp', tmp_path)
        cmd = [python, tmp_path, message, signal_path]
    else:
        ps_path = _create_temp_script('cda_lock_', '.ps1', _build_overlay_powershell(message, signal_path))
        setattr(stop_evt, '_tmp', ps_path)
        cmd = ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', ps_path]
    _run_user_session(cmd)

def unlock_screen():
    global _lock_event, _lock_thread
    _set_block_input(False)
    thread_to_join = None
    with _lock_mutex:
        if _lock_event:
            _lock_event.set()
            signal_path = getattr(_lock_event, '_signal', None)
            if signal_path:
                try:
                    Path(signal_path).write_text('unlock', encoding='utf-8')
                except Exception as e:
                    log.warning('Не удалось подать сигнал разблокировки: %s', e)
            tmp = getattr(_lock_event, '_tmp', None)
            if tmp:
                threading.Timer(10.0, lambda p=tmp: _cleanup_temp_file(p)).start()
            _lock_event = None
        if _lock_thread and _lock_thread.is_alive():
            thread_to_join = _lock_thread
        _lock_thread = None

    if thread_to_join and thread_to_join is not threading.current_thread():
        thread_to_join.join(timeout=2)

    for proc_name in ('python.exe', 'pythonw.exe', 'powershell.exe'):
        try:
            subprocess.run(
                ['taskkill', '/F', '/FI', f'IMAGENAME eq {proc_name}', '/FI', 'WINDOWTITLE eq ClassDeploy Lock*'],
                capture_output=True, timeout=5,
            )
        except Exception as e:
            log.debug('unlock taskkill %s: %s', proc_name, e)
    log.info("Экран разблокирован")


def _build_overlay_script(message: str, signal_path: str) -> str:
    safe = message.replace("\\", "\\\\").replace("'", "\\'").replace('"', '\\"')
    signal = signal_path.replace("\\", "\\\\").replace("'", "\\'")
    return f"""
import tkinter as tk
import os
from ctypes import *

msg = '{safe}'
signal_path = '{signal}'
windll.user32.BlockInput(True)
root = tk.Tk()
root.title("ClassDeploy Lock")
root.attributes("-fullscreen", True)
root.attributes("-topmost", True)
root.overrideredirect(True)
root.configure(bg="black")
for s in ("<Escape>","<Alt-F4>","<Control-c>","<Tab>","<Control-Shift-Escape>","<Key>"):
    root.bind_all(s, lambda e: "break")
root.protocol("WM_DELETE_WINDOW", lambda: None)
tk.Label(root, text=msg, bg="black", fg="white",
    font=("Segoe UI",48,"bold"),
    wraplength=root.winfo_screenwidth()-100,
    justify="center").place(relx=0.5, rely=0.38, anchor="center")
tk.Label(root, text="Экран заблокирован учителем",
    bg="black", fg="#555",
    font=("Segoe UI",18)).place(relx=0.5, rely=0.54, anchor="center")

def close(*_):
    try:
        root.destroy()
    except Exception:
        pass

def grab():
    if os.path.exists(signal_path):
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
    try: windll.user32.BlockInput(False)
    except Exception: pass
    try:
        if os.path.exists(signal_path):
            os.unlink(signal_path)
    except Exception:
        pass
"""

def _build_overlay_powershell(message: str, signal_path: str) -> str:
    safe = message.replace("'", "''")
    signal = signal_path.replace("'", "''")
    return f"""
Add-Type -AssemblyName System.Windows.Forms,System.Drawing
$f = New-Object Windows.Forms.Form
$f.Text = 'ClassDeploy Lock'
$f.FormBorderStyle = 'None'
$f.WindowState = 'Maximized'
$f.TopMost = $true
$f.BackColor = [Drawing.Color]::Black
$l = New-Object Windows.Forms.Label
$l.Text = '{safe}'
$l.ForeColor = [Drawing.Color]::White
$l.Font = New-Object Drawing.Font('Segoe UI',36,[Drawing.FontStyle]::Bold)
$l.AutoSize = $true
$f.Controls.Add($l)
$f.Add_Shown({{
    $l.Location = [Drawing.Point](($f.Width-$l.Width)/2,($f.Height-$l.Height)/2)
    $f.Activate()
}})
$timer = New-Object Windows.Forms.Timer
$timer.Interval = 300
$timer.Add_Tick({{
    if (Test-Path '{signal}') {{
        Remove-Item '{signal}' -ErrorAction SilentlyContinue
        $f.Close()
    }} else {{
        $f.TopMost = $true
        $f.Activate()
    }}
}})
$timer.Start()
[Windows.Forms.Application]::Run($f)
"""


# ════════════════════════════════════════════════════════
#  Скример
# ════════════════════════════════════════════════════════

def show_screamer(image_b64: str = ""):
    """
    Показать скример — полноэкранное изображение поверх всего.
    image_b64 — base64 JPEG/PNG. Если пусто — используется встроенная картинка.
    Скример закрывается любым кликом мыши или через 10 секунд.
    """
    _close_screamer()

    def _do_screamer():
        try:
            import tkinter as tk
            from PIL import Image, ImageTk
            import io

            img_data: Optional[bytes] = None
            if image_b64:
                try:
                    img_data = base64.b64decode(image_b64)
                except Exception:
                    pass

            root = tk.Tk()
            root.title("ClassDeploy Screamer")
            root.attributes("-fullscreen", True)
            root.attributes("-topmost", True)
            root.overrideredirect(True)
            root.configure(bg="black")

            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()

            if img_data:
                try:
                    pil_img = Image.open(io.BytesIO(img_data))
                    pil_img = pil_img.resize((sw, sh), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(pil_img)
                    lbl = tk.Label(root, image=photo, bg="black")
                    lbl.image = photo  # держим ссылку
                    lbl.place(x=0, y=0, width=sw, height=sh)
                except Exception:
                    img_data = None

            if not img_data:
                # Встроенная картинка: красный фон + текст
                root.configure(bg="#CC0000")
                tk.Label(
                    root, text="⚠",
                    bg="#CC0000", fg="white",
                    font=("Segoe UI", 200),
                ).place(relx=0.5, rely=0.3, anchor="center")
                tk.Label(
                    root, text="ВНИМАНИЕ!",
                    bg="#CC0000", fg="white",
                    font=("Segoe UI", 72, "bold"),
                ).place(relx=0.5, rely=0.65, anchor="center")

            def close(*_):
                global _screamer_close
                try:
                    root.destroy()
                except Exception:
                    pass
                finally:
                    with _screamer_lock:
                        if _screamer_close is close:
                            _screamer_close = None

            with _screamer_lock:
                globals()["_screamer_close"] = close

            root.bind("<Button-1>", close)
            root.bind("<Button-3>", close)
            root.bind("<Key>", close)
            root.after(10_000, close)  # автозакрытие через 10 сек

            try:
                root.focus_force()
            except Exception:
                pass
            root.mainloop()
        except Exception as e:
            log.error("screamer: %s", e)

    if _is_session_zero():
        # Служба: через Python в сессии пользователя
        python = _find_python()
        if not python:
            log.warning("Скример: Python не найден")
            return
        # Пишем данные во временный файл
        code = f"""
import tkinter as tk, base64, io
IMAGE_B64 = {repr(image_b64)}
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

root = tk.Tk()
root.title("ClassDeploy Screamer")
root.attributes("-fullscreen", True)
root.attributes("-topmost", True)
root.overrideredirect(True)
sw = root.winfo_screenwidth()
sh = root.winfo_screenheight()
shown = False
if HAS_PIL and IMAGE_B64:
    try:
        img_data = base64.b64decode(IMAGE_B64)
        pil_img = Image.open(io.BytesIO(img_data)).resize((sw, sh), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(pil_img)
        tk.Label(root, image=photo, bg="black").place(x=0, y=0, width=sw, height=sh)
        root.configure(bg="black")
        shown = True
    except Exception:
        pass
if not shown:
    root.configure(bg="#CC0000")
    tk.Label(root, text="\\u26a0", bg="#CC0000", fg="white", font=("Segoe UI",200)).place(relx=0.5, rely=0.3, anchor="center")
    tk.Label(root, text="\\u0412\\u041d\\u0418\\u041c\\u0410\\u041d\\u0418\\u0415!", bg="#CC0000", fg="white", font=("Segoe UI",72,"bold")).place(relx=0.5, rely=0.65, anchor="center")
def close(*_):
    try: root.destroy()
    except: pass
root.bind("<Button-1>", close)
root.bind("<Key>", close)
root.after(10000, close)
try: root.focus_force()
except: pass
root.mainloop()
"""
        tmp = tempfile.NamedTemporaryFile(
            mode="w", suffix=".py", delete=False,
            encoding="utf-8", prefix="cda_scream_",
        )
        tmp.write(code)
        tmp.close()
        _schtasks_run([python, tmp.name])
        threading.Timer(15.0, lambda: _safe_unlink(tmp.name)).start()
    else:
        threading.Thread(target=_do_screamer, daemon=True, name="Screamer").start()


def _safe_unlink(path: str):
    for _ in range(5):
        try:
            os.unlink(path)
            return
        except FileNotFoundError:
            return
        except Exception:
            time.sleep(0.2)
    log.debug('unlink failed: %s', path)


# ════════════════════════════════════════════════════════
#  Сообщение ученику
# ════════════════════════════════════════════════════════

def show_message(title: str, text: str):
    log.info("Сообщение ученику: %s / %s", title, text)

    def _show():
        try:
            import tkinter as tk
            from tkinter import messagebox
            r = tk.Tk()
            r.withdraw()
            r.attributes("-topmost", True)
            messagebox.showinfo(title, text)
            r.destroy()
        except Exception as e:
            log.debug("show_message: %s", e)

    if not _is_session_zero():
        threading.Thread(target=_show, daemon=True).start()
        return

    python = _find_python()
    if not python:
        try:
            subprocess.run(["msg", "*", f"{title}: {text}"], capture_output=True, timeout=10)
        except Exception:
            pass
        return

    safe_t = title.replace("'", "\\'")
    safe_x = text.replace("'", "\\'")
    py = (
        "import tkinter as tk\n"
        "from tkinter import messagebox\n"
        "r = tk.Tk()\n"
        "r.withdraw()\n"
        "r.attributes('-topmost', True)\n"
        f"messagebox.showinfo('{safe_t}', '{safe_x}')\n"
        "r.destroy()\n"
    )
    tmp_path = _create_temp_script('cda_msg_', '.py', py)
    _run_user_session([python, tmp_path])
    threading.Timer(20.0, lambda p=tmp_path: _cleanup_temp_file(p)).start()

def run_program(path: str, args: str = "") -> bool:
    path = (path or "").strip()
    if not path or any(ch in path for ch in ('\r', '\n', '\x00')):
        return False
    args = (args or "").strip()
    try:
        parsed_args = shlex.split(args, posix=False) if args else []
    except Exception as e:
        log.warning('run_program args: %s', e)
        return False
    cmd = [path, *parsed_args]
    ok = _run_user_session(cmd)
    log.info("run_program %s: %s", "ok" if ok else "fail", cmd)
    return ok

def open_vscode() -> bool:
    exe = _find_vscode()
    if exe:
        return run_program(exe, "--new-window")
    return _run_user_session(['cmd', '/c', 'start', '', 'code'])


def _find_vscode() -> Optional[str]:
    pf   = os.environ.get("ProgramFiles", r"C:\Program Files")
    pfx  = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    la   = os.environ.get("LOCALAPPDATA", "")
    candidates = [
        os.path.join(pf,  "Microsoft VS Code", "Code.exe"),
        os.path.join(pfx, "Microsoft VS Code", "Code.exe"),
    ]
    if la:
        candidates += [
            os.path.join(la, "Programs", "Microsoft VS Code", "Code.exe"),
            os.path.join(la, "Programs", "Microsoft VS Code Insiders", "Code - Insiders.exe"),
        ]
    for user_dir in Path(r"C:\Users").iterdir():
        if user_dir.is_dir():
            candidates.append(str(user_dir / "AppData" / "Local" / "Programs" / "Microsoft VS Code" / "Code.exe"))
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return None


def power(action: str):
    log.info("Power: %s", action)
    cmds = {
        "shutdown": ["shutdown", "/s", "/t", "5", "/f"],
        "reboot":   ["shutdown", "/r", "/t", "5", "/f"],
        "logoff":   ["shutdown", "/l", "/f"],
    }
    if action == "lock":
        ctypes.windll.user32.LockWorkStation()
    elif action in cmds:
        subprocess.Popen(cmds[action])
    else:
        log.warning("Неизвестная power команда: %s", action)
