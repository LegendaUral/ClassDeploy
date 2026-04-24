"""
Агент ClassDeploy.
- CoInitialize в нужных потоках (COM для pycaw/pywinauto)
- Глобальный перехват исключений — агент никогда не падает насовсем
- Watchdog: если event loop завис — перезапускает его
- Автопереподключение при потере сети (бесконечный retry с backoff)
- Все команды изолированы: ошибка в одной не трогает остальные
"""
from __future__ import annotations
import os
import sys
import ssl
import base64
import socket
import asyncio
import logging
import platform
import threading
import hashlib
import time
import subprocess
import ctypes
import traceback
import uuid
import getpass
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    import websockets
except ImportError:
    print("Установите: pip install websockets")
    sys.exit(1)

from shared import protocol as P
from shared.config import (
    SERVER_PORT, HEARTBEAT_INTERVAL,
    AGENT_DATA_DIR, AGENT_TEMP_DIR, AGENT_LOG_DIR,
    SCREEN_FPS, SCREEN_JPEG_QUALITY, SCREEN_MAX_WIDTH, WS_MAX_MESSAGE_SIZE,
)
from agent.screen import ScreenStreamer
from agent import remote
from agent.sound import SoundControl

# ── Логирование ─────────────────────────────────────────────────
def _session_id() -> int:
    try:
        sid = ctypes.c_ulong(0)
        ctypes.windll.kernel32.ProcessIdToSessionId(
            ctypes.windll.kernel32.GetCurrentProcessId(), ctypes.byref(sid)
        )
        return int(sid.value)
    except Exception:
        return -1


def _active_console_session_id() -> int:
    try:
        return int(ctypes.windll.kernel32.WTSGetActiveConsoleSessionId())
    except Exception:
        return -1


def _session_is_active() -> bool:
    active = _active_console_session_id()
    current = _session_id()
    return active < 0 or current < 0 or active == current


def _safe_name(value: str, default: str = "user") -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in (value or "").strip())
    return cleaned or default


def _log_candidates() -> list[Path]:
    username = _safe_name(os.environ.get("USERNAME") or getpass.getuser() or "user")
    local_appdata = os.environ.get("LOCALAPPDATA", "").strip()
    temp_dir = tempfile.gettempdir()
    return [
        Path(AGENT_LOG_DIR) / username / "agent.log",
        Path(local_appdata) / "ClassDeploy" / "logs" / "agent.log" if local_appdata else None,
        Path(temp_dir) / "ClassDeploy" / username / "agent.log",
        Path(AGENT_LOG_DIR) / f"agent_session_{max(_session_id(), 0)}.log",
    ]


def _build_log_handlers() -> tuple[list[logging.Handler], str | None]:
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    chosen: str | None = None
    for candidate in _log_candidates():
        if candidate is None:
            continue
        try:
            candidate.parent.mkdir(parents=True, exist_ok=True)
            with open(candidate, "a", encoding="utf-8"):
                pass
            handlers.insert(0, logging.FileHandler(candidate, encoding="utf-8"))
            chosen = str(candidate)
            break
        except Exception:
            continue
    return handlers, chosen


_log_handlers, ACTIVE_LOG_FILE = _build_log_handlers()
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=_log_handlers,
    force=True,
)
log = logging.getLogger("agent")
if ACTIVE_LOG_FILE:
    log.info("Лог агента: %s", ACTIVE_LOG_FILE)
else:
    log.warning("Не удалось открыть файловый лог, остаётся только вывод в консоль")


# ── Адреса сервера ───────────────────────────────────────────────

def _get_servers() -> list[str]:
    addrs: list[str] = []
    env = os.environ.get("CLASS_DEPLOY_SERVER", "").strip()
    if env:
        for p in env.replace(",", ";").split(";"):
            p = p.strip()
            if p:
                addrs.append(p)

    cfg = Path(AGENT_LOG_DIR).parent / "server.txt"
    if cfg.exists():
        for line in cfg.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                addrs.append(line)

    if not addrs:
        addrs = ["127.0.0.1"]

    return list(dict.fromkeys(addrs))


# ── CoInitialize для основного потока ────────────────────────────

def _coinit():
    try:
        ctypes.windll.ole32.CoInitializeEx(None, 0)
    except Exception:
        pass



def _get_agent_id() -> str:
    try:
        os.makedirs(AGENT_DATA_DIR, exist_ok=True)
        p = Path(AGENT_DATA_DIR) / "agent_id.txt"
        if p.exists():
            v = p.read_text(encoding="utf-8").strip()
            if v:
                return v
        v = uuid.uuid4().hex[:8]
        p.write_text(v, encoding="utf-8")
        return v
    except Exception:
        return uuid.uuid4().hex[:8]


# ── Агент ────────────────────────────────────────────────────────

class Agent:
    def __init__(self):
        self.hostname = platform.node()
        self.agent_id = _get_agent_id()
        self.os_info  = f"{platform.system()} {platform.release()}"
        self.servers  = _get_servers()
        self._last_server_refresh = 0.0

        self._ws   = None
        self._loop: asyncio.AbstractEventLoop | None = None

        self._streamer:  ScreenStreamer | None = None
        self._screen_inflight = None  # backpressure

        self._incoming: dict[str, dict] = {}
        self._remote_input = remote.RemoteInput()
        self._sound        = SoundControl()

        self._blocked_apps: set[str] = set()
        self._block_stop   = threading.Event()
        self._block_thread: threading.Thread | None = None

    # ── Главный цикл — никогда не падает ─────────────────────────

    async def run_forever(self):
        """Бесконечный цикл переподключения с экспоненциальным backoff."""
        delay = 3  # начальная задержка между попытками
        inactive_logged = False

        while True:
            try:
                if (time.time() - self._last_server_refresh) > 5:
                    self.servers = _get_servers()
                    self._last_server_refresh = time.time()
            except Exception:
                pass

            if not _session_is_active():
                if not inactive_logged:
                    log.info("Сессия %s не активна, агент ждёт входа в активную учётную запись", _session_id())
                    inactive_logged = True
                self._stop_stream()
                self._drop_incoming()
                await asyncio.sleep(2)
                continue

            inactive_logged = False
            connected_once = False
            for host in self.servers:
                try:
                    await self._connect(host)
                    connected_once = True
                    delay = 3  # сброс задержки после успешного соединения
                    break
                except Exception as e:
                    log.warning("Не удалось подключиться к %s: %s", host, e)
                    self._drop_incoming()
                    await asyncio.sleep(1)

            if not connected_once:
                log.info("Нет связи с сервером. Повтор через %d сек…", delay)
                await asyncio.sleep(delay)
                delay = min(delay * 2, 60)  # экспоненциальный backoff, макс 60 сек
            else:
                delay = 3

    async def _session_watchdog(self, ws):
        while True:
            await asyncio.sleep(1.0)
            if not _session_is_active():
                log.info("Текущая сессия %s стала неактивной, соединение закрывается", _session_id())
                self._stop_stream()
                try:
                    await ws.close()
                except Exception:
                    pass
                return

    async def _connect(self, host: str):
        url = f"ws://{host}:{SERVER_PORT}"
        log.info("Подключаюсь: %s", url)

        try:
            async with websockets.connect(
                url,
                max_size=WS_MAX_MESSAGE_SIZE,
                ping_interval=20,
                ping_timeout=40,
                open_timeout=15,
            ) as ws:
                self._ws   = ws
                self._loop = asyncio.get_running_loop()

                # Представляемся серверу
                await self._send(ws, P.Message(type=P.HELLO, payload={
                    "hostname": self.hostname,
                    "agent_id": self.agent_id,
                    "os":       self.os_info,
                    "ip":       _local_ip(),
                    "proto":    2,
                    "caps": [
                        P.SCREEN_START, P.SCREEN_STOP, P.INPUT_EVENT,
                        P.LOCK_SCREEN, P.UNLOCK_SCREEN, P.POWER, P.MESSAGE_BOX,
                        P.PROCESS_LIST, P.KILL_PROCESS, P.PUSH_FILE,
                        P.INSTALLER_CLICK, P.SOUND_CONTROL, P.PLAY_AUDIO, P.RUN_PROGRAM,
                        P.BLOCK_DOMAIN, P.UNBLOCK_DOMAIN, P.BLOCK_APP, P.UNBLOCK_APP,
                        P.SCREAMER, P.MOUSE_SENSITIVITY, P.WASD_INVERSION, P.SPEAK_TEXT,
                    ],
                }))

                # Ждём WELCOME
                try:
                    raw = await asyncio.wait_for(ws.recv(), timeout=15)
                    ack = P.Message.from_json(raw)
                    if ack.type != P.WELCOME:
                        raise Exception(f"Неожиданный ответ: {ack.type}")
                except asyncio.TimeoutError:
                    raise Exception("Таймаут ожидания WELCOME")

                log.info("Подключён к %s as %s (%s)", host, self.hostname, self.agent_id)

                try:
                    await asyncio.gather(
                        self._heartbeat(ws),
                        self._listen(ws),
                        self._session_watchdog(ws),
                    )
                finally:
                    self._stop_stream()
                    self._ws   = None
                    self._loop = None

        except websockets.ConnectionClosed as e:
            log.info("Соединение с %s закрыто: %s", host, e)
        except OSError as e:
            # Нет сети — не логируем как ошибку, просто переподключимся
            log.debug("Сеть недоступна (%s): %s", host, e)
            raise
        except Exception:
            raise

    # ── Сердцебиение ─────────────────────────────────────────────

    async def _heartbeat(self, ws):
        while True:
            try:
                await self._send(ws, P.Message(type=P.HEARTBEAT))
            except Exception:
                return
            await asyncio.sleep(HEARTBEAT_INTERVAL)

    # ── Приём команд ─────────────────────────────────────────────

    async def _listen(self, ws):
        async for raw in ws:
            try:
                msg = P.Message.from_json(raw)
                await self._dispatch(ws, msg)
            except Exception as e:
                log.exception("Ошибка обработки сообщения: %s", e)

    async def _dispatch(self, ws, msg: P.Message):
        """Каждая команда полностью изолирована — ошибка не рвёт соединение."""
        t = (msg.type or "").strip()
        # Алиасы camelCase для совместимости
        t = {
            "screenStart":   P.SCREEN_START,  "screenStop":    P.SCREEN_STOP,
            "lockScreen":    P.LOCK_SCREEN,    "unlockScreen":  P.UNLOCK_SCREEN,
            "runProgram":    P.RUN_PROGRAM,    "soundControl":  P.SOUND_CONTROL,
            "blockApp":      P.BLOCK_APP,      "unblockApp":    P.UNBLOCK_APP,
            "blockDomain":   P.BLOCK_DOMAIN,   "unblockDomain": P.UNBLOCK_DOMAIN,
        }.get(t, t)

        try:
            if   t == P.PING:            pass
            elif t == P.FILE_START:      self._file_start(msg)
            elif t == P.FILE_CHUNK:      self._file_chunk(msg)
            elif t == P.FILE_END:        await self._file_end(ws, msg)
            elif t == P.INSTALL:         await self._do_install(ws, msg)
            elif t == P.UNINSTALL:       await self._do_uninstall(ws, msg)
            elif t == P.PUSH_FILE:       await self._do_push_file(ws, msg)
            elif t == P.SCREEN_START:    self._start_stream()
            elif t == P.SCREEN_STOP:     self._stop_stream()
            elif t == P.INPUT_EVENT:     self._handle_input(msg)
            elif t == P.LOCK_SCREEN:
                text = msg.payload.get("message", "Внимание учителю")
                threading.Thread(target=remote.lock_screen, args=(text,),
                                 daemon=True).start()
            elif t == P.UNLOCK_SCREEN:
                threading.Thread(target=remote.unlock_screen, daemon=True).start()
            elif t == P.SCREAMER:
                img = msg.payload.get("image_b64", "")
                threading.Thread(target=remote.show_screamer, args=(img,),
                                 daemon=True).start()
            elif t == P.POWER:
                action = msg.payload.get("action", "shutdown")
                threading.Thread(target=remote.power, args=(action,),
                                 daemon=True).start()
            elif t == P.MESSAGE_BOX:
                threading.Thread(
                    target=remote.show_message,
                    args=(msg.payload.get("title", "Учитель"),
                          msg.payload.get("text", "")),
                    daemon=True,
                ).start()
            elif t == P.PROCESS_LIST:    await self._do_process_list(ws, msg)
            elif t == P.KILL_PROCESS:    await self._do_kill_process(ws, msg)
            elif t == P.INSTALLER_CLICK: await self._do_installer_click(ws, msg)
            elif t == P.SOUND_CONTROL:   self._handle_sound(msg)
            elif t == P.PLAY_AUDIO:      await self._do_play_audio(ws, msg)
            elif t == P.RUN_PROGRAM:     self._handle_run(msg)
            elif t == P.BLOCK_DOMAIN:    self._block_domains(msg)
            elif t == P.UNBLOCK_DOMAIN:  self._unblock_domains(msg)
            elif t == P.BLOCK_APP:       self._block_apps(msg)
            elif t == P.UNBLOCK_APP:     self._unblock_apps(msg)
            elif t == P.MOUSE_SENSITIVITY: self._handle_mouse_sensitivity(msg)
            elif t == P.WASD_INVERSION:  self._handle_wasd_inversion(msg)
            elif t == P.SPEAK_TEXT:      self._handle_speak_text(msg)
            elif t == P.SHUTDOWN:
                log.info("SHUTDOWN от сервера")
                sys.exit(0)
            else:
                log.debug("Неизвестный тип: %s", t)

        except Exception as e:
            # Любая ошибка команды логируется, НЕ роняет агента
            log.exception("Ошибка команды %s: %s", t, e)

    # ── Трансляция экрана ────────────────────────────────────────

    def _start_stream(self):
        if self._streamer:
            return

        def send_frame(b64: str, w: int, h: int):
            ws   = self._ws
            loop = self._loop
            if not (ws and loop):
                return
            try:
                prev = self._screen_inflight
                if prev is not None and not prev.done():
                    return  # backpressure

                fut = asyncio.run_coroutine_threadsafe(
                    ws.send(P.Message(
                        type=P.SCREEN_FRAME,
                        payload={"data": b64, "w": w, "h": h},
                    ).to_json()),
                    loop,
                )
                self._screen_inflight = fut

                def _on_done(f):
                    if not f.cancelled() and f.exception() and self._streamer:
                        self._stop_stream()

                fut.add_done_callback(_on_done)
            except Exception:
                self._stop_stream()

        self._streamer = ScreenStreamer(
            send_frame,
            fps=SCREEN_FPS,
            quality=SCREEN_JPEG_QUALITY,
            max_width=SCREEN_MAX_WIDTH,
        )
        self._streamer.start()

    def _stop_stream(self):
        if self._streamer:
            try:
                self._streamer.stop()
            except Exception:
                pass
            self._streamer = None

    # ── Ввод ─────────────────────────────────────────────────────

    def _handle_input(self, msg: P.Message):
        p    = msg.payload
        kind = p.get("kind", "")
        sw   = int(p.get("screen_w", 1920))
        sh   = int(p.get("screen_h", 1080))
        ri   = self._remote_input
        try:
            if   kind == "mouse_move":   ri.move_mouse(p["x"], p["y"], sw, sh)
            elif kind == "mouse_click":  ri.click(p["x"], p["y"], sw, sh,
                                                   button=p.get("button", "left"),
                                                   double=bool(p.get("double", False)))
            elif kind == "mouse_button": ri.mouse_button(
                                                   p["x"], p["y"], sw, sh,
                                                   button=p.get("button", "left"),
                                                   down=bool(p.get("down", True)))
            elif kind == "scroll":       ri.scroll(p["x"], p["y"], sw, sh,
                                                    int(p.get("delta", 0)))
            elif kind == "key":          ri.key(int(p["vk"]), bool(p.get("down", True)))
            elif kind == "type":        ri.type_text(p.get("text", ""))
            elif kind == "combo":
                vks = [int(v) for v in p.get("vks", [])]
                if vks:
                    ri.key_combination(*vks)
        except Exception as e:
            log.debug("input %s: %s", kind, e)

    # ── Приём файлов ─────────────────────────────────────────────

    def _file_start(self, msg: P.Message):
        os.makedirs(AGENT_TEMP_DIR, exist_ok=True)
        filename = msg.payload.get("filename", "file")
        size     = int(msg.payload.get("size", 0))
        path     = Path(AGENT_TEMP_DIR) / f"{msg.job_id}_{filename}"
        try:
            fp = open(path, "wb")
        except Exception as e:
            log.error("FILE_START open: %s", e)
            return
        self._incoming[msg.job_id] = {
            "path": path, "size": size, "got": 0,
            "sha256": str(msg.payload.get("sha256", "")).strip().lower(),
            "fp": fp,
        }
        log.info("Приём файла: %s (%d байт)", filename, size)

    def _file_chunk(self, msg: P.Message):
        info = self._incoming.get(msg.job_id)
        if not info:
            return
        try:
            data = base64.b64decode(msg.payload["data"])
            info["fp"].write(data)
            info["got"] += len(data)
        except Exception as e:
            log.debug("FILE_CHUNK: %s", e)

    async def _file_end(self, ws, msg: P.Message):
        info = self._incoming.get(msg.job_id)
        if not info:
            return
        try:
            info["fp"].close()
        except Exception:
            pass

        expected = int(info.get("size", 0))
        got      = int(info.get("got", 0))
        if expected and expected != got:
            self._incoming.pop(msg.job_id, None)
            _try_remove(info["path"])
            await self._send(ws, P.Message(
                type=P.RESULT, job_id=msg.job_id,
                payload={"ok": False, "message": f"Размер: {got} / {expected}"},
            ))
            return

        sha_exp = info.get("sha256", "")
        if sha_exp:
            sha_got = _sha256(str(info["path"]))
            if sha_got != sha_exp:
                self._incoming.pop(msg.job_id, None)
                _try_remove(info["path"])
                await self._send(ws, P.Message(
                    type=P.RESULT, job_id=msg.job_id,
                    payload={"ok": False, "message": "SHA256 не совпал"},
                ))
                return

        log.info("Файл принят: %s (%d байт)", info["path"], got)
        await self._send(ws, P.Message(
            type=P.STATUS, job_id=msg.job_id,
            payload={"status": P.ST_RECEIVING, "message": "Файл получен"},
        ))

    # ── Установка / удаление ─────────────────────────────────────

    async def _do_install(self, ws, msg: P.Message):
        from agent import installer
        info = self._incoming.get(msg.job_id)
        if not info:
            await self._send(ws, P.Message(
                type=P.RESULT, job_id=msg.job_id,
                payload={"ok": False, "message": "Файл не передан"},
            ))
            return

        flags = msg.payload.get("flags", "")
        if msg.payload.get("force_autoclick"):
            flags = (flags + " --force-autoclick").strip()

        loop = asyncio.get_running_loop()

        def worker():
            # CoInitialize в потоке установки (нужен pywinauto)
            _coinit()

            def on_st(status, message):
                asyncio.run_coroutine_threadsafe(
                    self._send(ws, P.Message(
                        type=P.STATUS, job_id=msg.job_id,
                        payload={"status": status, "message": message},
                    )), loop,
                )

            try:
                ok, result = installer.install(str(info["path"]), on_st, flags)
            except Exception as e:
                ok, result = False, f"Исключение: {e}"

            asyncio.run_coroutine_threadsafe(
                self._send(ws, P.Message(
                    type=P.RESULT, job_id=msg.job_id,
                    payload={"ok": ok, "message": result},
                )), loop,
            )
            _try_remove(info["path"])
            self._incoming.pop(msg.job_id, None)

        threading.Thread(target=worker, daemon=True, name="InstallWorker").start()

    async def _do_uninstall(self, ws, msg: P.Message):
        from agent import installer
        name = msg.payload.get("program_name", "")
        loop = asyncio.get_running_loop()

        def worker():
            _coinit()

            def on_st(status, message):
                asyncio.run_coroutine_threadsafe(
                    self._send(ws, P.Message(
                        type=P.STATUS, job_id=msg.job_id,
                        payload={"status": status, "message": message},
                    )), loop,
                )
            try:
                ok, result = installer.uninstall(name, on_st)
            except Exception as e:
                ok, result = False, str(e)

            asyncio.run_coroutine_threadsafe(
                self._send(ws, P.Message(
                    type=P.RESULT, job_id=msg.job_id,
                    payload={"ok": ok, "message": result},
                )), loop,
            )

        threading.Thread(target=worker, daemon=True).start()

    async def _do_push_file(self, ws, msg: P.Message):
        import shutil
        info = self._incoming.get(msg.job_id)
        if not info:
            await self._send(ws, P.Message(
                type=P.RESULT, job_id=msg.job_id,
                payload={"ok": False, "message": "Файл не передан"},
            ))
            return

        src        = str(info["path"])
        target_dir = os.path.expandvars(msg.payload.get("target_dir", ""))
        filename   = msg.payload.get("filename", os.path.basename(src))
        overwrite  = bool(msg.payload.get("overwrite", True))

        try:
            os.makedirs(target_dir, exist_ok=True)
            dst = os.path.join(target_dir, filename)
            if os.path.exists(dst) and not overwrite:
                await self._send(ws, P.Message(
                    type=P.RESULT, job_id=msg.job_id,
                    payload={"ok": False, "message": f"Уже существует: {dst}"},
                ))
                return
            shutil.copy2(src, dst)
            log.info("PUSH_FILE: %s → %s", src, dst)
            await self._send(ws, P.Message(
                type=P.RESULT, job_id=msg.job_id,
                payload={"ok": True, "message": f"Скопировано: {dst}"},
            ))
        except Exception as e:
            log.exception("PUSH_FILE: %s", e)
            await self._send(ws, P.Message(
                type=P.RESULT, job_id=msg.job_id,
                payload={"ok": False, "message": str(e)},
            ))
        finally:
            _try_remove(src)
            self._incoming.pop(msg.job_id, None)

    async def _do_play_audio(self, ws, msg: P.Message):
        import shutil

        info = self._incoming.get(msg.job_id)
        if not info:
            await self._send(ws, P.Message(
                type=P.RESULT, job_id=msg.job_id,
                payload={"ok": False, "message": "Аудиофайл не передан"},
            ))
            return

        src = Path(str(info["path"]))
        dst = Path(AGENT_TEMP_DIR) / f"broadcast_audio_{msg.job_id}{src.suffix}"
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            ok = self._sound.play_file(str(dst), msg.payload.get("volume"), cleanup_path=str(dst))
            await self._send(ws, P.Message(
                type=P.RESULT, job_id=msg.job_id,
                payload={"ok": bool(ok), "message": "Аудио запущено" if ok else "Не удалось запустить аудио"},
            ))
        except Exception as e:
            await self._send(ws, P.Message(
                type=P.RESULT, job_id=msg.job_id,
                payload={"ok": False, "message": str(e)},
            ))
        finally:
            _try_remove(src)
            self._incoming.pop(msg.job_id, None)

    # ── Процессы ─────────────────────────────────────────────────

    async def _do_process_list(self, ws, msg: P.Message):
        try:
            procs = await asyncio.get_running_loop().run_in_executor(
                None, _list_procs
            )
            await self._send(ws, P.Message(
                type=P.PROCESS_LIST_RESULT, job_id=msg.job_id,
                payload={"ok": True, "processes": procs},
            ))
        except Exception as e:
            log.exception("PROCESS_LIST: %s", e)
            await self._send(ws, P.Message(
                type=P.PROCESS_LIST_RESULT, job_id=msg.job_id,
                payload={"ok": False, "message": str(e)},
            ))

    async def _do_kill_process(self, ws, msg: P.Message):
        pid = int(msg.payload.get("pid", 0))
        try:
            ok, result = _kill_proc(pid)
            await self._send(ws, P.Message(
                type=P.KILL_PROCESS_RESULT, job_id=msg.job_id,
                payload={"ok": ok, "message": result, "pid": pid},
            ))
        except Exception as e:
            await self._send(ws, P.Message(
                type=P.KILL_PROCESS_RESULT, job_id=msg.job_id,
                payload={"ok": False, "message": str(e), "pid": pid},
            ))

    async def _do_installer_click(self, ws, msg: P.Message):
        from agent import installer
        btn_text = str(msg.payload.get("button_text", "")).strip()
        try:
            ok, result = await asyncio.get_running_loop().run_in_executor(
                None, installer.installer_click, btn_text
            )
        except Exception as e:
            ok, result = False, str(e)
        await self._send(ws, P.Message(
            type=P.STATUS, job_id=msg.job_id,
            payload={
                "status":  P.ST_AUTOCLICK if ok else P.ST_ERROR,
                "message": result,
            },
        ))

    # ── Звук ─────────────────────────────────────────────────────

    def _handle_sound(self, msg: P.Message):
        action = msg.payload.get("action", "")
        try:
            if action == "mute":
                self._sound.mute()
            elif action == "unmute":
                self._sound.unmute()
            elif action == "set_volume":
                self._sound.set_volume(float(msg.payload.get("volume", 0.5)))
            elif action == "stop_audio":
                self._sound.stop_playback()
            elif action == "play_pending":
                info = self._incoming.get(msg.job_id)
                if not info:
                    log.warning("play_pending: файл не найден для job %s", msg.job_id)
                    return
                path = str(info.get("path") or "")
                ok = self._sound.play_file(path, msg.payload.get("volume"))
                log.info("play_pending %s: %s", msg.job_id, "ok" if ok else "fail")
            elif action == "play_file":
                path = str(msg.payload.get("path") or "")
                ok = self._sound.play_file(path, msg.payload.get("volume"))
                log.info("play_file %s: %s", path, "ok" if ok else "fail")
            else:
                log.debug("Неизвестный sound action: %s", action)
        except Exception as e:
            log.debug("sound %s: %s", action, e)

    # ── Запуск программ ──────────────────────────────────────────

    def _handle_run(self, msg: P.Message):
        kind = (msg.payload.get("kind") or "").strip().lower()
        path = msg.payload.get("path", "")
        args = msg.payload.get("args", "")
        try:
            if kind == "vscode":
                remote.open_vscode()
            else:
                remote.run_program(path, args)
        except Exception as e:
            log.warning("run_program: %s", e)

    # ── Блокировки ───────────────────────────────────────────────

    def _block_domains(self, msg: P.Message):
        domains = msg.payload.get("domains", [])
        hosts   = r"C:\Windows\System32\drivers\etc\hosts"
        try:
            existing = Path(hosts).read_text(encoding="utf-8")
            lines_to_add = []
            for d in domains:
                entry = f"127.0.0.1 {d}"
                if entry not in existing:
                    lines_to_add.append(entry)
            if lines_to_add:
                with open(hosts, "a", encoding="utf-8") as f:
                    f.write("\n" + "\n".join(lines_to_add) + "\n")
        except Exception as e:
            log.debug("block_domains: %s", e)

    def _unblock_domains(self, msg: P.Message):
        domains = set(msg.payload.get("domains", []))
        hosts   = r"C:\Windows\System32\drivers\etc\hosts"
        try:
            lines = Path(hosts).read_text(encoding="utf-8").splitlines(keepends=True)
            with open(hosts, "w", encoding="utf-8") as f:
                for line in lines:
                    if not any(f"127.0.0.1 {d}" in line for d in domains):
                        f.write(line)
        except Exception as e:
            log.debug("unblock_domains: %s", e)

    def _block_apps(self, msg: P.Message):
        apps = msg.payload.get("apps", [])
        self._blocked_apps.update(apps)
        if not (self._block_thread and self._block_thread.is_alive()):
            self._block_stop.clear()
            self._block_thread = threading.Thread(
                target=self._app_killer, daemon=True, name="AppBlocker"
            )
            self._block_thread.start()

    def _unblock_apps(self, msg: P.Message):
        self._blocked_apps.difference_update(msg.payload.get("apps", []))
        if not self._blocked_apps:
            self._block_stop.set()

    # ── Чувствительность мыши ───────────────────────────────────

    def _handle_mouse_sensitivity(self, msg: P.Message):
        try:
            speed = int(msg.payload.get("speed", 10))
        except (ValueError, TypeError):
            speed = 10
        try:
            ok = remote.set_mouse_sensitivity(speed)
            if ok:
                log.info("Чувствительность мыши установлена: %d", speed)
            else:
                log.warning("Ошибка при установке чувствительности мыши")
        except Exception as e:
            log.debug("mouse_sensitivity: %s", e)

    # ── Инверсия WASD ───────────────────────────────────────────

    def _handle_wasd_inversion(self, msg: P.Message):
        action = msg.payload.get("action", "toggle")  # toggle, enable, disable
        try:
            if action == "enable":
                ok = remote.set_wasd_inversion(True)
            elif action == "disable":
                ok = remote.set_wasd_inversion(False)
            else:  # toggle
                ok = remote.toggle_wasd_inversion()
            log.info("WASD инверсия (%s): %s", action, "ok" if ok else "fail")
        except Exception as e:
            log.debug("wasd_inversion %s: %s", action, e)

    # ── Озвучка текста ──────────────────────────────────────────

    def _handle_speak_text(self, msg: P.Message):
        text = str(msg.payload.get("text", "")).strip()
        try:
            volume = float(msg.payload.get("volume", 1.0))
        except (ValueError, TypeError):
            volume = 1.0
        try:
            rate = int(msg.payload.get("rate", -2))
        except (ValueError, TypeError):
            rate = -2
        try:
            if text:
                ok = self._sound.speak_text(text, volume, rate)
                log.info("speak_text: %s", "ok" if ok else "fail")
            else:
                log.debug("speak_text: empty text")
        except Exception as e:
            log.debug("speak_text: %s", e)

    def _app_killer(self):
        while not self._block_stop.wait(4):
            for app in list(self._blocked_apps):
                try:
                    subprocess.run(
                        ["taskkill", "/f", "/im", app],
                        capture_output=True, timeout=5,
                    )
                except Exception:
                    pass

    # ── Утилиты ──────────────────────────────────────────────────

    async def _send(self, ws, msg: P.Message):
        try:
            await ws.send(msg.to_json())
        except Exception as e:
            log.debug("send: %s", e)
            raise

    def _drop_incoming(self):
        for info in list(self._incoming.values()):
            try:
                fp = info.get("fp")
                if fp:
                    fp.close()
            except Exception:
                pass
            _try_remove(info.get("path"))
        self._incoming.clear()


# ── Системные утилиты ────────────────────────────────────────────

def _local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "?"


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _try_remove(path):
    if path:
        try:
            os.remove(str(path))
        except Exception:
            pass


def _list_procs() -> list[dict]:
    import json
    # MainModule.FileName требует прав — добавляем try/catch на уровне PS
    ps = (
        "Get-Process | ForEach-Object { "
        "  $p = $_; "
        "  $path = ''; "
        "  try { $path = $p.MainModule.FileName } catch {}; "
        "  [PSCustomObject]@{ "
        "    Pid=$p.Id; Name=$p.ProcessName; "
        "    SessionId=$p.SessionId; "
        "    MemKB=[math]::Round($p.WorkingSet64/1KB); "
        "    Path=$path "
        "  } "
        "} | ConvertTo-Json -Compress"
    )
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps],
            capture_output=True, timeout=25,
        )
        raw = (r.stdout or b"").decode("utf-8", errors="replace").strip()
        if not raw:
            return []
        data = json.loads(raw)
        if isinstance(data, dict):
            data = [data]
    except Exception as e:
        log.warning("_list_procs: %s", e)
        return []

    result = []
    for p in data:
        try:
            pid = int(p.get("Pid") or 0)
            if pid <= 0:
                continue
            mem_kb = p.get("MemKB", 0)
            mem = f"{int(mem_kb):,} KB".replace(",", " ") if mem_kb else ""
            result.append({
                "pid":     pid,
                "name":    str(p.get("Name") or ""),
                "session": str(p.get("SessionId") or ""),
                "memory":  mem,
                "path":    str(p.get("Path") or ""),
            })
        except Exception:
            continue
    return result


def _kill_proc(pid: int) -> tuple[bool, str]:
    try:
        r = subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)], capture_output=True,
        )
        out = (r.stdout or b"").decode("cp866", errors="replace").strip()
        err = (r.stderr or b"").decode("cp866", errors="replace").strip()
        if r.returncode == 0:
            return True, out or "Процесс завершён"
        return False, err or out or f"Код {r.returncode}"
    except Exception as e:
        return False, str(e)


# ── Точка входа ──────────────────────────────────────────────────

def main():
    # CoInitialize для главного потока
    _coinit()

    log.info("=== ClassDeploy Agent запущен ===")
    log.info("Хост: %s  IP: %s", platform.node(), _local_ip())

    agent = Agent()
    log.info("Серверы: %s", ", ".join(agent.servers))

    # Глобальный перехват неожиданных исключений
    def _excepthook(exc_type, exc_val, exc_tb):
        log.critical(
            "НЕПЕРЕХВАЧЕННОЕ ИСКЛЮЧЕНИЕ:\n%s",
            "".join(traceback.format_exception(exc_type, exc_val, exc_tb)),
        )

    sys.excepthook = _excepthook

    try:
        asyncio.run(agent.run_forever())
    except KeyboardInterrupt:
        log.info("Остановлен пользователем")
    except Exception as e:
        log.critical("КРИТИЧЕСКАЯ ОШИБКА АГЕНТА: %s", e, exc_info=True)
        # Пауза и перезапуск (если запущен как служба — SCM перезапустит)
        time.sleep(5)


if __name__ == "__main__":
    main()
