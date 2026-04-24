"""
Общие настройки для сервера и агентов.
Пароль убран — только IP сервера.
"""
from __future__ import annotations
import os
from pathlib import Path

SERVER_PORT = 8765

def _load_server_addr() -> str:
    env = os.environ.get("CLASS_DEPLOY_SERVER", "").strip()
    if env:
        return env
    try:
        p = Path(r"C:\ProgramData\ClassDeploy\server.txt")
        if p.exists():
            lines = [l.strip() for l in p.read_text(encoding="utf-8").splitlines()
                     if l.strip() and not l.startswith("#")]
            if lines:
                return lines[0]
    except Exception:
        pass
    return "127.0.0.1"

SERVER_ADDR = _load_server_addr()

HEARTBEAT_INTERVAL   = 10
OFFLINE_AFTER        = 30
INSTALL_TIMEOUT      = 30 * 60
SILENT_PROBE_TIMEOUT = 5

SILENT_FLAGS = [
    "/S", "/SILENT", "/VERYSILENT", "/quiet",
    "/silent", "/s", "-s", "-silent", "--silent", "/qn", "/passive",
]

AGENT_DATA_DIR = r"C:\ProgramData\ClassDeploy"
AGENT_TEMP_DIR = r"C:\ProgramData\ClassDeploy\temp"
AGENT_LOG_DIR  = r"C:\ProgramData\ClassDeploy\logs"
PORTABLE_DIR   = r"C:\PortableApps"

SCREEN_FPS          = 12
SCREEN_JPEG_QUALITY = 70
SCREEN_MAX_WIDTH    = 1280
WS_MAX_MESSAGE_SIZE = 16 * 1024 * 1024
