"""
WebSocket-сервер: принимает агентов, рассылает команды.
Авторизация по паролю убрана — работает в доверенной сети класса.
"""
from __future__ import annotations
import os
import base64
import asyncio
import hashlib
import logging
import time
from pathlib import Path
from typing import Callable, Dict, Optional

import websockets

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from shared import protocol as P
from shared.protocol import FILE_CHUNK_SIZE
from shared.config import SERVER_PORT, OFFLINE_AFTER, WS_MAX_MESSAGE_SIZE
from server.database import Database

log = logging.getLogger("server.network")


class AgentConn:
    def __init__(self, ws, hostname: str, ip: str, os_info: str):
        self.ws        = ws
        self.hostname  = hostname
        self.ip        = ip
        self.os        = os_info
        self.last_seen = time.time()
        self.caps: set[str] = set()

    async def send(self, msg: P.Message):
        await self.ws.send(msg.to_json())


class Server:
    def __init__(self, db: Database):
        self.db    = db
        self.agents: Dict[str, AgentConn] = {}

        self.on_agent_change: Optional[Callable[[], None]]                   = None
        self.on_job_update:   Optional[Callable[[str, str, str], None]]      = None
        self.on_screen_frame: Optional[Callable[[str, str, int, int], None]] = None

        self.job_to_host: Dict[str, str]           = {}
        self._pending:    Dict[str, asyncio.Future] = {}

    async def start(self, host: str = "0.0.0.0", port: int = SERVER_PORT):
        self.ws_server = await websockets.serve(
            self._handle, host, port,
            max_size=WS_MAX_MESSAGE_SIZE,
            ping_interval=20, ping_timeout=60,
        )
        log.info("WebSocket слушает %s:%s", host, port)
        asyncio.create_task(self._monitor())

    async def _monitor(self):
        while True:
            await asyncio.sleep(5)
            now  = time.time()
            dead = [
                (h, a) for h, a in list(self.agents.items())
                if now - a.last_seen > OFFLINE_AFTER
            ]
            for h, a in dead:
                log.info("Агент %s оффлайн (таймаут)", h)
                try:
                    await a.ws.close()
                except Exception:
                    pass
            if dead and self.on_agent_change:
                self.on_agent_change()

    async def _handle(self, ws, *args):
        peer     = ws.remote_address
        hostname = None
        conn: AgentConn | None = None

        try:
            raw   = await asyncio.wait_for(ws.recv(), timeout=30)
            hello = P.Message.from_json(raw)

            if hello.type != P.HELLO:
                log.warning("Первое сообщение не HELLO от %s", peer)
                return

            hostname = hello.payload.get("hostname", str(peer))
            ip       = hello.payload.get("ip", peer[0] if peer else "?")
            os_info  = hello.payload.get("os", "?")

            # Закрываем старое соединение при переподключении
            old = self.agents.get(hostname)
            if old is not None and old.ws is not ws:
                log.info("Переподключение %s — закрываю старое", hostname)
                try:
                    await old.ws.close()
                except Exception:
                    pass

            conn = AgentConn(ws, hostname, ip, os_info)
            try:
                caps = hello.payload.get("caps", [])
                conn.caps = {str(x) for x in caps if x}
            except Exception:
                conn.caps = set()

            self.agents[hostname] = conn
            self.db.upsert_agent(hostname, ip, os_info)

            # Сразу подтверждаем без проверки пароля
            await ws.send(P.Message(type=P.WELCOME, payload={"ok": True}).to_json())
            log.info("Агент подключился: %s (%s)", hostname, ip)
            if self.on_agent_change:
                self.on_agent_change()

            async for raw in ws:
                try:
                    msg = P.Message.from_json(raw)
                    await self._handle_msg(conn, msg)
                except Exception as e:
                    log.exception("Ошибка сообщения от %s: %s", hostname, e)

        except asyncio.TimeoutError:
            log.warning("HELLO таймаут от %s", peer)
        except websockets.ConnectionClosed:
            log.info("Агент %s отключился", hostname or peer)
        except Exception as e:
            log.exception("_handle: %s", e)
        finally:
            if hostname and conn is not None:
                if self.agents.get(hostname) is conn:
                    self.agents.pop(hostname, None)
                    if self.on_agent_change:
                        self.on_agent_change()

    async def _handle_msg(self, conn: AgentConn, msg: P.Message):
        conn.last_seen = time.time()
        self.db.touch_agent(conn.hostname)

        if msg.type == P.HEARTBEAT:
            return

        if msg.type == P.STATUS:
            self.db.log_job(msg.job_id, msg.payload.get("status",""), msg.payload.get("message",""))
            if self.on_job_update:
                self.on_job_update(msg.job_id, msg.payload.get("status",""), msg.payload.get("message",""))

        elif msg.type == P.RESULT:
            ok      = bool(msg.payload.get("ok"))
            message = msg.payload.get("message", "")
            self.db.finish_job(msg.job_id, ok, message)
            if self.on_job_update:
                self.on_job_update(msg.job_id, "done" if ok else "error", message)
            fut = self._pending.pop(msg.job_id, None)
            if fut and not fut.done():
                fut.set_result(msg.payload)

        elif msg.type == P.SCREEN_FRAME:
            b64 = msg.payload.get("data", "")
            w   = int(msg.payload.get("w", 1920))
            h   = int(msg.payload.get("h", 1080))
            if b64 and self.on_screen_frame:
                self.on_screen_frame(conn.hostname, b64, w, h)

        elif msg.type in (P.PROCESS_LIST_RESULT, P.KILL_PROCESS_RESULT):
            fut = self._pending.pop(msg.job_id, None)
            if fut and not fut.done():
                fut.set_result(msg.payload)

    # ── Установка ───────────────────────────────────────────────

    async def install_file(self, hostnames: list[str], file_path: str,
                           custom_flags: str = "", force_autoclick: bool = False) -> list[str]:
        job_ids  = []
        size     = os.path.getsize(file_path)
        filename = os.path.basename(file_path)
        for host in hostnames:
            agent = self.agents.get(host)
            if not agent:
                continue
            job_id = P.new_job_id()
            job_ids.append(job_id)
            self.job_to_host[job_id] = host
            self.db.create_job(job_id, host, "install", filename)
            asyncio.create_task(self._send_file_and_install(
                agent, job_id, file_path, filename, size, custom_flags, force_autoclick,
            ))
        return job_ids

    async def _send_file_and_install(self, agent: AgentConn, job_id: str,
                                      file_path: str, filename: str, size: int,
                                      flags: str, force_ac: bool):
        try:
            sha = _sha256(file_path)
            await agent.send(P.Message(type=P.FILE_START, job_id=job_id,
                payload={"filename": filename, "size": size, "sha256": sha}))
            with open(file_path, "rb") as f:
                sent = 0
                while True:
                    chunk = f.read(FILE_CHUNK_SIZE)
                    if not chunk:
                        break
                    await agent.send(P.Message(type=P.FILE_CHUNK, job_id=job_id,
                        payload={"data": base64.b64encode(chunk).decode("ascii")}))
                    sent += len(chunk)
                    if self.on_job_update and size > 0:
                        self.on_job_update(job_id, "uploading", f"{int(sent*100/size)}%")
            await agent.send(P.Message(type=P.FILE_END, job_id=job_id))
            await agent.send(P.Message(type=P.INSTALL, job_id=job_id,
                payload={"flags": flags, "force_autoclick": force_ac}))
        except Exception as e:
            log.exception("install_file %s: %s", agent.hostname, e)
            self.db.finish_job(job_id, False, str(e))
            if self.on_job_update:
                self.on_job_update(job_id, P.ST_ERROR, str(e))

    async def uninstall(self, hostnames: list[str], program_name: str) -> list[str]:
        job_ids = []
        for host in hostnames:
            agent = self.agents.get(host)
            if not agent:
                continue
            job_id = P.new_job_id()
            job_ids.append(job_id)
            self.db.create_job(job_id, host, "uninstall", program_name)
            await agent.send(P.Message(type=P.UNINSTALL, job_id=job_id,
                payload={"program_name": program_name}))
        return job_ids

    async def push_file(self, hostnames: list[str], file_path: str,
                        target_dir: str, overwrite: bool = True) -> list[str]:
        job_ids  = []
        size     = os.path.getsize(file_path)
        filename = os.path.basename(file_path)
        for host in hostnames:
            agent = self.agents.get(host)
            if not agent:
                continue
            job_id = P.new_job_id()
            job_ids.append(job_id)
            self.job_to_host[job_id] = host
            self.db.create_job(job_id, host, "push_file", f"{filename} → {target_dir}")
            asyncio.create_task(self._send_file_and_push(
                agent, job_id, file_path, filename, size, target_dir, overwrite,
            ))
        return job_ids

    async def _send_file_and_push(self, agent: AgentConn, job_id: str,
                                   file_path: str, filename: str, size: int,
                                   target_dir: str, overwrite: bool):
        try:
            sha = _sha256(file_path)
            await agent.send(P.Message(type=P.FILE_START, job_id=job_id,
                payload={"filename": filename, "size": size, "sha256": sha}))
            with open(file_path, "rb") as f:
                sent = 0
                while True:
                    chunk = f.read(FILE_CHUNK_SIZE)
                    if not chunk:
                        break
                    await agent.send(P.Message(type=P.FILE_CHUNK, job_id=job_id,
                        payload={"data": base64.b64encode(chunk).decode("ascii")}))
                    sent += len(chunk)
                    if self.on_job_update and size > 0:
                        self.on_job_update(job_id, "uploading", f"{int(sent*100/size)}%")
            await agent.send(P.Message(type=P.FILE_END, job_id=job_id))
            await agent.send(P.Message(type=P.PUSH_FILE, job_id=job_id,
                payload={"target_dir": target_dir, "filename": filename, "overwrite": overwrite}))
        except Exception as e:
            log.exception("push_file %s: %s", agent.hostname, e)
            self.db.finish_job(job_id, False, str(e))
            if self.on_job_update:
                self.on_job_update(job_id, P.ST_ERROR, str(e))

    # ── Экраны ──────────────────────────────────────────────────

    async def screen_start(self, hostnames: list[str]):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.SCREEN_START))

    async def screen_stop(self, hostnames: list[str]):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.SCREEN_STOP))

    async def send_input(self, hostname: str, payload: dict):
        a = self.agents.get(hostname)
        if a:
            await a.send(P.Message(type=P.INPUT_EVENT, payload=payload))

    # ── Управление ──────────────────────────────────────────────

    async def lock_screens(self, hostnames: list[str], message: str = "Внимание учителю"):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.LOCK_SCREEN, payload={"message": message}))

    async def unlock_screens(self, hostnames: list[str]):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.UNLOCK_SCREEN))

    async def show_screamer(self, hostnames: list[str], image_b64: str = ""):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.SCREAMER, payload={"image_b64": image_b64}))

    async def power(self, hostnames: list[str], action: str):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.POWER, payload={"action": action}))

    async def send_message_box(self, hostnames: list[str], title: str, text: str):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.MESSAGE_BOX, payload={"title": title, "text": text}))

    async def installer_click(self, hostnames: list[str], button_text: str):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.INSTALLER_CLICK, job_id=P.new_job_id(),
                    payload={"button_text": button_text}))

    # ── Процессы ────────────────────────────────────────────────

    async def get_process_list(self, hostname: str, timeout: float = 20.0) -> dict:
        agent = self.agents.get(hostname)
        if not agent:
            return {"ok": False, "message": "Агент не онлайн"}
        job_id = P.new_job_id()
        fut    = asyncio.get_running_loop().create_future()
        self._pending[job_id] = fut
        try:
            await agent.send(P.Message(type=P.PROCESS_LIST, job_id=job_id))
            return await asyncio.wait_for(fut, timeout=timeout)
        except asyncio.TimeoutError:
            return {"ok": False, "message": "Таймаут"}
        finally:
            self._pending.pop(job_id, None)

    async def kill_process(self, hostname: str, pid: int, timeout: float = 15.0) -> dict:
        agent = self.agents.get(hostname)
        if not agent:
            return {"ok": False, "message": "Агент не онлайн"}
        job_id = P.new_job_id()
        fut    = asyncio.get_running_loop().create_future()
        self._pending[job_id] = fut
        try:
            await agent.send(P.Message(type=P.KILL_PROCESS, job_id=job_id, payload={"pid": pid}))
            return await asyncio.wait_for(fut, timeout=timeout)
        except asyncio.TimeoutError:
            return {"ok": False, "message": "Таймаут"}
        finally:
            self._pending.pop(job_id, None)

    # ── Звук / программы / ограничения ──────────────────────────

    async def sound_control(self, hostnames: list[str], action: str, volume: float = None):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                pl = {"action": action}
                if volume is not None:
                    pl["volume"] = volume
                await a.send(P.Message(type=P.SOUND_CONTROL, payload=pl))

    async def speak_text(self, hostnames: list[str], text: str, volume: float | None = None, rate: int | None = None):
        text = (text or "").strip()
        if not text:
            return
        for h in hostnames:
            a = self.agents.get(h)
            if not a:
                continue
            payload = {"text": text}
            if volume is not None:
                payload["volume"] = volume
            if rate is not None:
                payload["rate"] = int(rate)
            await a.send(P.Message(type=P.SPEAK_TEXT, payload=payload))

    async def play_audio(self, hostnames: list[str], file_path: str, volume: float | None = None) -> list[str]:
        job_ids = []
        size = os.path.getsize(file_path)
        filename = os.path.basename(file_path)
        for host in hostnames:
            agent = self.agents.get(host)
            if not agent:
                continue
            job_id = P.new_job_id()
            job_ids.append(job_id)
            self.job_to_host[job_id] = host
            self.db.create_job(job_id, host, "play_audio", filename)
            asyncio.create_task(self._send_file_and_play_audio(agent, job_id, file_path, filename, size, volume))
        return job_ids

    async def _send_file_and_play_audio(self, agent: AgentConn, job_id: str,
                                        file_path: str, filename: str, size: int,
                                        volume: float | None):
        try:
            sha = _sha256(file_path)
            await agent.send(P.Message(type=P.FILE_START, job_id=job_id,
                payload={"filename": filename, "size": size, "sha256": sha}))
            with open(file_path, "rb") as f:
                sent = 0
                while True:
                    chunk = f.read(FILE_CHUNK_SIZE)
                    if not chunk:
                        break
                    await agent.send(P.Message(type=P.FILE_CHUNK, job_id=job_id,
                        payload={"data": base64.b64encode(chunk).decode("ascii")}))
                    sent += len(chunk)
                    if self.on_job_update and size > 0:
                        self.on_job_update(job_id, "uploading", f"{int(sent*100/size)}%")
            await agent.send(P.Message(type=P.FILE_END, job_id=job_id))
            payload = {"volume": volume} if volume is not None else {}
            await agent.send(P.Message(type=P.PLAY_AUDIO, job_id=job_id, payload=payload))
        except Exception as e:
            log.exception("play_audio %s: %s", agent.hostname, e)
            self.db.finish_job(job_id, False, str(e))
            if self.on_job_update:
                self.on_job_update(job_id, P.ST_ERROR, str(e))

    async def run_program(self, hostnames: list[str], path: str, args: str = ""):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.RUN_PROGRAM, payload={"path": path, "args": args}))

    async def open_vscode(self, hostnames: list[str]):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.RUN_PROGRAM, payload={"kind": "vscode"}))

    async def block_domains(self, hostnames: list[str], domains: list[str]):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.BLOCK_DOMAIN, payload={"domains": domains}))

    async def unblock_domains(self, hostnames: list[str], domains: list[str]):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.UNBLOCK_DOMAIN, payload={"domains": domains}))

    async def block_apps(self, hostnames: list[str], apps: list[str]):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.BLOCK_APP, payload={"apps": apps}))

    async def unblock_apps(self, hostnames: list[str], apps: list[str]):
        for h in hostnames:
            a = self.agents.get(h)
            if a:
                await a.send(P.Message(type=P.UNBLOCK_APP, payload={"apps": apps}))


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()
