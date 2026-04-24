"""
Протокол сообщений. Все сообщения — JSON с полем "type".
Пароль убран: авторизация не нужна в доверенной сети класса.
"""
from __future__ import annotations
import json
import uuid
from dataclasses import dataclass, field, asdict
from typing import Any, Dict

# Агент → Сервер
HELLO               = "hello"
HEARTBEAT           = "heartbeat"
STATUS              = "status"
RESULT              = "result"
SCREEN_FRAME        = "screen_frame"

# Сервер → Агент
WELCOME             = "welcome"         # сервер принял агента
FILE_START          = "file_start"
FILE_CHUNK          = "file_chunk"
FILE_END            = "file_end"
INSTALL             = "install"
UNINSTALL           = "uninstall"
PING                = "ping"
SHUTDOWN            = "shutdown"

SCREEN_START        = "screen_start"
SCREEN_STOP         = "screen_stop"
INPUT_EVENT         = "input_event"
LOCK_SCREEN         = "lock_screen"
UNLOCK_SCREEN       = "unlock_screen"
POWER               = "power"
MESSAGE_BOX         = "message_box"
PROCESS_LIST        = "process_list"
PROCESS_LIST_RESULT = "process_list_result"
KILL_PROCESS        = "kill_process"
KILL_PROCESS_RESULT = "kill_process_result"
PUSH_FILE           = "push_file"
INSTALLER_CLICK     = "installer_click"
SOUND_CONTROL       = "sound_control"
PLAY_AUDIO          = "play_audio"
RUN_PROGRAM         = "run_program"
BLOCK_DOMAIN        = "block_domain"
UNBLOCK_DOMAIN      = "unblock_domain"
BLOCK_APP           = "block_app"
UNBLOCK_APP         = "unblock_app"
SCREAMER            = "screamer"
MOUSE_SENSITIVITY   = "mouse_sensitivity"
WASD_INVERSION      = "wasd_inversion"
SPEAK_TEXT          = "speak_text"

# Статусы заданий
ST_RECEIVING  = "receiving"
ST_INSTALLING = "installing"
ST_AUTOCLICK  = "autoclick"
ST_DONE       = "done"
ST_ERROR      = "error"
ST_TIMEOUT    = "timeout"


@dataclass
class Message:
    type: str
    job_id: str = ""
    payload: Dict[str, Any] = field(default_factory=dict)

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False)

    @staticmethod
    def from_json(raw: str) -> "Message":
        d = json.loads(raw)
        return Message(
            type=d.get("type", ""),
            job_id=d.get("job_id", ""),
            payload=d.get("payload", {}),
        )


def new_job_id() -> str:
    return uuid.uuid4().hex[:12]

FILE_CHUNK_SIZE = 512 * 1024
MAX_FILE_SIZE   = 5 * 1024 ** 3
