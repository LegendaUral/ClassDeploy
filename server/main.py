"""
Точка входа сервера ClassDeploy.
PyQt6 в главном потоке, asyncio WebSocket в фоновом.
"""
from __future__ import annotations
import sys
import os
import asyncio
import logging
import socket
import threading
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from PyQt6.QtWidgets import QApplication, QMessageBox

from shared.config import SERVER_PORT
from server.database import Database
from server.network import Server
from server.scheduler import Scheduler
from server.gui import MainWindow, QtBridge

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("main")


def _local_ips() -> list[str]:
    ips = {"127.0.0.1"}
    try:
        for ip in socket.gethostbyname_ex(socket.gethostname())[2]:
            if ip and "." in ip:
                ips.add(ip)
    except Exception:
        pass
    return sorted(ips)


def main():
    app = QApplication(sys.argv)

    db     = Database()
    server = Server(db)
    bridge = QtBridge()

    server.on_agent_change = bridge.agents_changed.emit
    server.on_job_update   = lambda j, s, m: bridge.job_update.emit(j, s, m)
    server.on_screen_frame = lambda h, b, w, ht: bridge.screen_frame.emit(h, b, w, ht)

    loop         = asyncio.new_event_loop()
    startup_done = threading.Event()
    startup_err: list[Exception] = []

    async def _startup():
        try:
            await server.start()
            loop.create_task(Scheduler(db, server).run_forever())
        except Exception as e:
            startup_err.append(e)
        finally:
            startup_done.set()

    def _run_loop():
        asyncio.set_event_loop(loop)
        loop.create_task(_startup())
        loop.run_forever()

    threading.Thread(target=_run_loop, daemon=True).start()

    if not startup_done.wait(timeout=15):
        QMessageBox.critical(None, "Ошибка", "Сервер не запустился за 15 сек.")
        sys.exit(1)

    if startup_err:
        QMessageBox.critical(None, "Ошибка запуска",
                             f"Не удалось запустить сервер:\n{startup_err[0]}\n\n"
                             f"Проверьте, что порт {SERVER_PORT} свободен.")
        sys.exit(1)

    win = MainWindow(server, db, bridge, loop)
    win.show()

    ips_str = "  /  ".join(_local_ips())
    log.info("Сервер запущен. Порт: %s", SERVER_PORT)
    log.info("Агенты подключаются к: ws://<IP>:%s", SERVER_PORT)
    log.info("IP-адреса этого ПК: %s", ips_str)

    rc = app.exec()
    loop.call_soon_threadsafe(loop.stop)
    sys.exit(rc)


if __name__ == "__main__":
    main()
