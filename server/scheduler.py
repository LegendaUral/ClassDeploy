"""
Планировщик: раз в 30 секунд запускает отложенные задания.
"""
from __future__ import annotations
import json
import asyncio
import logging
from server.database import Database

log = logging.getLogger("server.scheduler")


class Scheduler:
    def __init__(self, db: Database, server):
        self.db     = db
        self.server = server

    async def run_forever(self):
        while True:
            try:
                for row in self.db.due_schedules():
                    try:
                        hostnames = json.loads(row["hostnames"])
                        action    = row["action"]
                        log.info("Плановое задание #%s: %s на %s",
                                 row["id"], action, hostnames)
                        if action == "install":
                            ids = await self.server.install_file(
                                hostnames, row["file_path"], row["custom_flags"] or "",
                            )
                        elif action == "uninstall":
                            ids = await self.server.uninstall(
                                hostnames, row["file_path"],
                            )
                        else:
                            ids = []

                        if ids:
                            self.db.mark_schedule_done(row["id"])
                        else:
                            log.warning("Задание #%s не запущено — нет онлайн-агентов", row["id"])
                    except Exception as e:
                        log.exception("Ошибка планового задания #%s: %s", row["id"], e)
            except Exception as e:
                log.exception("Ошибка планировщика: %s", e)
            await asyncio.sleep(30)
