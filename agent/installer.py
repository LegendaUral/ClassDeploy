"""
Установщик: MSI/EXE/ZIP без вмешательства пользователя.
Автокликер с CoInitialize в потоке и надёжным поиском окон.
"""
from __future__ import annotations
import os
import sys
import re
import time
import ctypes
import zipfile
import logging
import subprocess
import threading
from pathlib import Path
from typing import Callable, Tuple, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))
from shared.config import SILENT_FLAGS, SILENT_PROBE_TIMEOUT, INSTALL_TIMEOUT, PORTABLE_DIR
from shared import protocol as P

log = logging.getLogger("agent.installer")

StatusCallback = Callable[[str, str], None]

_active_pid: Optional[int] = None
_active_pid_lock = threading.Lock()


def _coinit():
    """CoInitialize для текущего потока (COM нужен pywinauto и pycaw)."""
    try:
        ctypes.windll.ole32.CoInitializeEx(None, 0)
    except Exception:
        pass


def _load_pywinauto():
    """Загружает pywinauto лениво. Возвращает модуль или None."""
    try:
        import pywinauto
        return pywinauto
    except ImportError:
        log.warning("pywinauto не установлен — автокликер недоступен")
        return None
    except Exception as e:
        log.warning("pywinauto ошибка загрузки: %s", e)
        return None


# ── Публичный API ──────────────────────────────────────────────────

def install(file_path: str, on_status: StatusCallback,
            custom_flags: str = "") -> Tuple[bool, str]:
    fp = Path(file_path)
    if not fp.exists():
        return False, f"Файл не найден: {file_path}"

    force_autoclick = False
    flags = (custom_flags or "").strip()
    if "--force-autoclick" in flags:
        force_autoclick = True
        flags = flags.replace("--force-autoclick", "").strip()

    ext = fp.suffix.lower()
    log.info("Установка: %s (ext=%s force_ac=%s)", fp.name, ext, force_autoclick)
    on_status(P.ST_INSTALLING, f"Установка {fp.name}")

    try:
        if ext == ".msi":
            return _install_msi(fp, flags, on_status)
        if ext == ".exe":
            if force_autoclick:
                flags = (flags + " --force-autoclick").strip()
            return _install_exe(fp, flags, on_status)
        if ext in (".zip", ".7z"):
            return _install_portable(fp, on_status)
        return False, f"Неподдерживаемое расширение: {ext}"
    except Exception as e:
        log.exception("Критическая ошибка установки")
        return False, f"Исключение: {e}"


def uninstall(program_name: str, on_status: StatusCallback) -> Tuple[bool, str]:
    on_status(P.ST_INSTALLING, f"Удаляю {program_name}")
    # PowerShell сначала
    ps = (
        f"Get-Package -Name '*{program_name}*' -ErrorAction SilentlyContinue | "
        f"Uninstall-Package -Force -ErrorAction SilentlyContinue"
    )
    rc, out = _run(["powershell", "-NoProfile", "-Command", ps], INSTALL_TIMEOUT)
    if rc == 0:
        return True, "Удалено через PowerShell"
    # WMIC fallback
    rc2, out2 = _run(
        ["wmic", "product", "where", f"name like '%{program_name}%'",
         "call", "uninstall", "/nointeractive"],
        INSTALL_TIMEOUT,
    )
    if rc2 == 0 and "ReturnValue = 0" in out2:
        return True, "Удалено через WMIC"
    return False, f"Не удалось удалить (rc={rc2})"


def installer_click(button_text: str) -> Tuple[bool, str]:
    """Ручной клик по кнопке в любом открытом окне установщика."""
    # CoInitialize обязателен для pywinauto в потоке агента
    _coinit()
    pwa = _load_pywinauto()
    if not pwa:
        return False, "pywinauto недоступен"

    try:
        from pywinauto import Desktop
        target = _norm(button_text or "")
        pid    = _get_active_pid()

        if pid and _proc_alive(pid):
            windows = Desktop(backend="uia").windows(process=pid, visible_only=True)
        else:
            windows = Desktop(backend="uia").windows(visible_only=True)

        for w in windows:
            try:
                for btn in w.descendants(control_type="Button"):
                    try:
                        if not (btn.is_enabled() and btn.is_visible()):
                            continue
                        txt = _norm(btn.window_text())
                        if target and target not in txt:
                            continue
                        btn.click_input()
                        log.info("installer_click: '%s'", btn.window_text())
                        return True, f"Клик: {btn.window_text()}"
                    except Exception:
                        continue
            except Exception:
                continue
    except Exception as e:
        log.debug("installer_click error: %s", e)

    return False, f"Кнопка не найдена: '{button_text}'"


# ── MSI ───────────────────────────────────────────────────────────

def _install_msi(fp: Path, custom_flags: str, cb: StatusCallback) -> Tuple[bool, str]:
    flags = custom_flags or "/quiet /norestart"
    cmd   = ["msiexec", "/i", str(fp)] + flags.split()
    log.info("MSI: %s", " ".join(cmd))
    cb(P.ST_INSTALLING, "msiexec /quiet")
    rc, out = _run(cmd, INSTALL_TIMEOUT)
    if rc in (0, 3010):
        return True, f"MSI установлен (код {rc})"
    return False, f"msiexec код {rc}: {out[-400:]}"


# ── EXE ──────────────────────────────────────────────────────────

def _install_exe(fp: Path, flags: str, cb: StatusCallback) -> Tuple[bool, str]:
    force_ac    = "--force-autoclick" in flags
    clean_flags = flags.replace("--force-autoclick", "").strip()

    if clean_flags and not force_ac:
        return _try_flag(fp, clean_flags, cb)

    tried = []
    if not force_ac:
        for flag in SILENT_FLAGS:
            cb(P.ST_INSTALLING, f"Флаг: {flag}")
            ok, msg = _try_flag(fp, flag, cb)
            if ok:
                return True, f"Флаг {flag}: {msg}"
            tried.append(f"{flag}: {msg[:60]}")
    else:
        cb(P.ST_AUTOCLICK, "Force-autoclick: пропускаю тихие флаги")

    log.warning("Тихие флаги не сработали → автокликер")
    cb(P.ST_AUTOCLICK, "Запускаю автокликер…")
    ok, msg = _autoclicker(fp, cb)
    if ok:
        return True, f"Автокликер: {msg}"
    return False, "Все методы не сработали:\n" + "\n".join(tried) + f"\nАвтокликер: {msg}"


def _try_flag(fp: Path, flag: str, cb: StatusCallback) -> Tuple[bool, str]:
    try:
        proc = subprocess.Popen(
            [str(fp)] + flag.split(),
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            creationflags=_cflags(),
        )
    except Exception as e:
        return False, f"Не запустить: {e}"

    deadline  = time.time() + SILENT_PROBE_TIMEOUT
    has_win   = False
    while time.time() < deadline:
        if proc.poll() is not None:
            break
        if _has_window(proc.pid):
            has_win = True
            break
        time.sleep(0.4)

    if has_win:
        _kill_tree(proc.pid)
        return False, "GUI окно появилось"

    try:
        rc = proc.wait(timeout=INSTALL_TIMEOUT)
    except subprocess.TimeoutExpired:
        _kill_tree(proc.pid)
        return False, "таймаут"

    return (True, f"код {rc}") if rc == 0 else (False, f"код {rc}")


def _autoclicker(fp: Path, cb: StatusCallback) -> Tuple[bool, str]:
    # CoInitialize ОБЯЗАТЕЛЕН — pywinauto использует COM
    _coinit()

    pwa = _load_pywinauto()
    if not pwa:
        return False, "pywinauto недоступен"

    try:
        proc = subprocess.Popen([str(fp)], creationflags=_cflags())
    except Exception as e:
        return False, f"Не запустить: {e}"

    _set_active_pid(proc.pid)

    PRIMARY = [
        r"^(next|далее|continue|продолжить).*$",
        r"^(i agree|agree|accept|согласен|принять).*$",
        r"^(install|установить|install now|установить сейчас).*$",
        r"^(yes|да|ok|ок).*$",
    ]
    FINISH = [
        r"^(finish|готово|close|закрыть|done|завершить).*$",
        r".*(launch|run|запустить).*$",
    ]
    SKIP = [
        r"^(cancel|отмена|no|нет|decline|отклонить).*$",
        r".*(назад|back).*$",
    ]
    EULA = [
        r".*(i agree|accept|license|eula|соглашаюсь|принимаю|лиценз).*$",
    ]

    try:
        from pywinauto import Desktop
    except Exception as e:
        _clear_active_pid(proc.pid)
        return False, f"Desktop import: {e}"

    # Ждём появления первого окна (до 30 сек)
    log.info("Автокликер: жду окно установщика (pid=%d)", proc.pid)
    wait_deadline = time.time() + 30
    while time.time() < wait_deadline:
        if proc.poll() is not None:
            _clear_active_pid(proc.pid)
            return False, f"Установщик закрылся до появления окна (код {proc.returncode})"
        if _has_window(proc.pid):
            log.info("Автокликер: окно появилось")
            break
        time.sleep(0.5)

    deadline   = time.time() + INSTALL_TIMEOUT
    last_click = time.time()
    clicked    = False
    idle_warn  = 0.0

    while time.time() < deadline:
        if proc.poll() is not None:
            rc = proc.returncode
            _clear_active_pid(proc.pid)
            if clicked:
                return True, f"Завершился (код {rc})"
            # Мог завершиться тихо (код 0) даже без кликов — это тоже успех
            if rc == 0:
                return True, f"Завершился без кликов (код 0)"
            return False, f"Закрылся без кликов (код {rc})"

        try:
            wins = Desktop(backend="uia").windows(process=proc.pid, visible_only=True)
        except Exception as e:
            log.debug("Desktop.windows: %s", e)
            time.sleep(1)
            continue

        for w in wins:
            try:
                # 1. Чекбоксы EULA
                try:
                    for chk in w.descendants(control_type="CheckBox"):
                        try:
                            txt = _norm(chk.window_text())
                            if txt and _match(txt, EULA):
                                state = chk.get_toggle_state()
                                if state == 0:
                                    chk.click_input()
                                    clicked    = True
                                    last_click = time.time()
                                    cb(P.ST_AUTOCLICK, f"Чекбокс: {txt[:40]}")
                                    time.sleep(0.4)
                        except Exception:
                            continue
                except Exception:
                    pass

                # 2. Кнопки — сначала финальные, потом основные
                try:
                    buttons = w.descendants(control_type="Button")
                except Exception:
                    continue

                hit = _click_buttons(buttons, FINISH,  SKIP, cb, finish=True)
                if not hit:
                    hit = _click_buttons(buttons, PRIMARY, SKIP, cb, finish=False)

                if hit:
                    clicked    = True
                    last_click = time.time()
                    time.sleep(1.5)  # ждём следующий экран
                    break

            except Exception as e:
                log.debug("autoclicker window iter: %s", e)
                continue

        now = time.time()
        if clicked and now - last_click > 40 and now - idle_warn > 25:
            idle_warn = now
            cb(P.ST_AUTOCLICK, "Жду следующий экран установщика…")

        # Зависание > 90 сек без клика после первого клика — аварийный выход
        if clicked and now - last_click > 90:
            _kill_tree(proc.pid)
            _clear_active_pid(proc.pid)
            return False, "Установщик завис (нет активности > 90 сек)"

        time.sleep(0.6)

    _kill_tree(proc.pid)
    _clear_active_pid(proc.pid)
    return False, "Таймаут автокликера"


def _click_buttons(buttons, include: list, skip: list,
                   cb: StatusCallback, finish: bool) -> bool:
    for btn in buttons:
        try:
            if not btn.is_visible():
                continue
            if not btn.is_enabled():
                continue
            txt = _norm(btn.window_text())
            if not txt:
                continue
            if _match(txt, skip):
                continue
            if not _match(txt, include):
                continue
            btn.click_input()
            cb(P.ST_AUTOCLICK, f"{'✅' if finish else '▶'} Кнопка: {txt[:40]}")
            if finish:
                time.sleep(2.0)
            return True
        except Exception:
            continue
    return False


# ── Portable (ZIP) ───────────────────────────────────────────────

def _install_portable(fp: Path, cb: StatusCallback) -> Tuple[bool, str]:
    name   = fp.stem
    target = Path(PORTABLE_DIR) / name
    target.mkdir(parents=True, exist_ok=True)
    cb(P.ST_INSTALLING, f"Распаковка в {target}")
    try:
        with zipfile.ZipFile(fp) as z:
            _safe_extract(z, target)
    except Exception as e:
        return False, f"Ошибка распаковки: {e}"

    exes = list(target.rglob("*.exe"))
    if exes:
        main_exe = min(exes, key=lambda p: len(p.parts))
        _shortcut(main_exe, name)
        return True, f"Распаковано в {target}, ярлык для {main_exe.name}"
    return True, f"Распаковано в {target}"


def _shortcut(target_exe: Path, name: str):
    try:
        import win32com.client
        desk = Path(os.environ.get("PUBLIC", r"C:\Users\Public")) / "Desktop"
        desk.mkdir(exist_ok=True)
        lnk  = desk / f"{name}.lnk"
        s    = win32com.client.Dispatch("WScript.Shell").CreateShortcut(str(lnk))
        s.TargetPath      = str(target_exe)
        s.WorkingDirectory = str(target_exe.parent)
        s.Save()
    except Exception as e:
        log.warning("Ярлык не создан: %s", e)


def _safe_extract(zf: zipfile.ZipFile, target: Path):
    base = target.resolve()
    for m in zf.infolist():
        dst = (base / m.filename).resolve()
        if not str(dst).startswith(str(base)):
            raise RuntimeError(f"Небезопасный путь: {m.filename}")
    zf.extractall(base)


# ── Утилиты ──────────────────────────────────────────────────────

def _cflags() -> int:
    if sys.platform == "win32":
        return subprocess.CREATE_NO_WINDOW
    return 0


def _run(cmd: list, timeout: int) -> Tuple[int, str]:
    try:
        r = subprocess.run(
            cmd, capture_output=True, timeout=timeout,
            creationflags=_cflags(),
        )
        out = (r.stdout or b"").decode("utf-8", errors="replace")
        err = (r.stderr or b"").decode("utf-8", errors="replace")
        return r.returncode, out + err
    except subprocess.TimeoutExpired:
        return -1, "таймаут"
    except FileNotFoundError as e:
        return -1, f"не найдено: {e}"
    except Exception as e:
        return -1, str(e)


def _has_window(pid: int) -> bool:
    if sys.platform != "win32":
        return False
    try:
        from ctypes import wintypes
        u32      = ctypes.windll.user32
        ENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        found    = [False]

        def _cb(hwnd, _):
            if not u32.IsWindowVisible(hwnd):
                return True
            pid_val = wintypes.DWORD()
            u32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid_val))
            if pid_val.value == pid and u32.GetWindowTextLengthW(hwnd) > 0:
                found[0] = True
                return False
            return True

        u32.EnumWindows(ENUMPROC(_cb), 0)
        return found[0]
    except Exception:
        return False


def _kill_tree(pid: int):
    try:
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            capture_output=True, timeout=10, creationflags=_cflags(),
        )
    except Exception:
        pass


def _proc_alive(pid: int) -> bool:
    try:
        r = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}"],
            capture_output=True, timeout=5, creationflags=_cflags(),
        )
        return str(pid) in (r.stdout or b"").decode("utf-8", errors="replace")
    except Exception:
        return False


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip().lower())


def _match(text: str, patterns: list) -> bool:
    return any(re.match(p, text, re.IGNORECASE) for p in patterns)


def _set_active_pid(pid: int):
    global _active_pid
    with _active_pid_lock:
        _active_pid = pid


def _get_active_pid() -> Optional[int]:
    with _active_pid_lock:
        return _active_pid


def _clear_active_pid(pid: Optional[int] = None):
    global _active_pid
    with _active_pid_lock:
        if pid is None or _active_pid == pid:
            _active_pid = None
