from __future__ import annotations
import os
import sqlite3
import threading
import time
from pathlib import Path


def _clear_readonly(path: Path):
    try:
        if path.exists():
            os.chmod(path, 0o666)
    except Exception:
        pass


def _test_dir(base: Path) -> bool:
    try:
        base.mkdir(parents=True, exist_ok=True)
        probe = base / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink(missing_ok=True)
        return True
    except Exception:
        return False


def _db_path() -> Path:
    env = os.environ.get("CLASS_DEPLOY_DB_PATH", "").strip()
    if env:
        p = Path(env).expanduser()
        p.parent.mkdir(parents=True, exist_ok=True)
        return p

    candidates = []
    data_env = os.environ.get("CLASS_DEPLOY_DATA_DIR", "").strip()
    if data_env:
        candidates.append(Path(data_env).expanduser())
    candidates.append(Path(os.environ.get("ProgramData", r"C:\ProgramData")) / "ClassDeploy" / "server")
    candidates.append(Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "ClassDeploy" / "server")
    candidates.append(Path(os.environ.get("TEMP", "/tmp")) / "ClassDeployServer")

    for base in candidates:
        if _test_dir(base):
            return base / "class_deploy.db"

    # Last resort.
    return Path("class_deploy.db").resolve()


class Database:
    def __init__(self):
        self._lock = threading.RLock()
        self.db_path = _db_path()
        self.conn = self._open_connection(self.db_path)
        self.conn.row_factory = sqlite3.Row
        try:
            self.conn.execute("PRAGMA journal_mode=WAL")
        except Exception:
            pass
        self._init_schema()

    def _open_connection(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        for extra in (path, Path(str(path) + "-wal"), Path(str(path) + "-shm")):
            _clear_readonly(extra)
        try:
            return sqlite3.connect(str(path), check_same_thread=False)
        except sqlite3.OperationalError as e:
            msg = str(e).lower()
            if "readonly" not in msg:
                raise
        except PermissionError:
            pass

        fb = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "ClassDeploy" / "server"
        fb.mkdir(parents=True, exist_ok=True)
        fb_db = fb / "class_deploy.db"
        return sqlite3.connect(str(fb_db), check_same_thread=False)

    def _execute_write(self, sql: str, params=()):
        try:
            self.conn.execute(sql, params)
            self.conn.commit()
            return
        except sqlite3.OperationalError as e:
            if "readonly" not in str(e).lower():
                raise

        # Reopen DB in a guaranteed writable fallback and retry once.
        with self._lock:
            try:
                self.conn.close()
            except Exception:
                pass
            fb = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "ClassDeploy" / "server"
            fb.mkdir(parents=True, exist_ok=True)
            self.db_path = fb / "class_deploy.db"
            self.conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
            self.conn.row_factory = sqlite3.Row
            try:
                self.conn.execute("PRAGMA journal_mode=WAL")
            except Exception:
                pass
            self._init_schema()
            self.conn.execute(sql, params)
            self.conn.commit()

    def _init_schema(self):
        with self._lock:
            self.conn.executescript("""
                CREATE TABLE IF NOT EXISTS agents (
                    hostname   TEXT PRIMARY KEY,
                    ip         TEXT,
                    os         TEXT,
                    last_seen  REAL,
                    first_seen REAL
                );
                CREATE TABLE IF NOT EXISTS jobs (
                    job_id      TEXT PRIMARY KEY,
                    hostname    TEXT,
                    action      TEXT,
                    filename    TEXT,
                    status      TEXT,
                    result_msg  TEXT,
                    started_at  REAL,
                    finished_at REAL
                );
                CREATE TABLE IF NOT EXISTS job_logs (
                    id      INTEGER PRIMARY KEY AUTOINCREMENT,
                    job_id  TEXT,
                    ts      REAL,
                    status  TEXT,
                    message TEXT
                );
                CREATE TABLE IF NOT EXISTS scheduled (
                    id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_at       REAL,
                    hostnames    TEXT,
                    file_path    TEXT,
                    action       TEXT,
                    custom_flags TEXT,
                    done         INTEGER DEFAULT 0
                );
            """)
            self.conn.commit()

    def upsert_agent(self, hostname: str, ip: str, os_info: str):
        now = time.time()
        with self._lock:
            self._execute_write("""
                INSERT INTO agents (hostname, ip, os, last_seen, first_seen)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(hostname) DO UPDATE SET
                    ip=excluded.ip, os=excluded.os, last_seen=excluded.last_seen
            """, (hostname, ip, os_info, now, now))

    def touch_agent(self, hostname: str):
        with self._lock:
            self._execute_write(
                "UPDATE agents SET last_seen=? WHERE hostname=?",
                (time.time(), hostname),
            )

    def list_agents(self):
        with self._lock:
            return list(self.conn.execute("SELECT * FROM agents ORDER BY hostname"))

    def create_job(self, job_id: str, hostname: str, action: str, filename: str):
        with self._lock:
            self._execute_write("""
                INSERT INTO jobs (job_id, hostname, action, filename, status, started_at)
                VALUES (?, ?, ?, ?, 'running', ?)
            """, (job_id, hostname, action, filename, time.time()))

    def finish_job(self, job_id: str, ok: bool, message: str):
        with self._lock:
            self._execute_write("""
                UPDATE jobs SET status=?, result_msg=?, finished_at=? WHERE job_id=?
            """, ("done" if ok else "error", message, time.time(), job_id))

    def log_job(self, job_id: str, status: str, message: str):
        with self._lock:
            self._execute_write("""
                INSERT INTO job_logs (job_id, ts, status, message) VALUES (?, ?, ?, ?)
            """, (job_id, time.time(), status, message))

    def get_job_logs(self, job_id: str):
        with self._lock:
            return list(self.conn.execute(
                "SELECT * FROM job_logs WHERE job_id=? ORDER BY ts", (job_id,)
            ))

    def recent_jobs(self, limit: int = 100):
        with self._lock:
            return list(self.conn.execute(
                "SELECT * FROM jobs ORDER BY started_at DESC LIMIT ?", (limit,)
            ))

    def add_schedule(self, run_at: float, hostnames_json: str,
                     file_path: str, action: str, flags: str):
        with self._lock:
            self._execute_write("""
                INSERT INTO scheduled (run_at, hostnames, file_path, action, custom_flags)
                VALUES (?, ?, ?, ?, ?)
            """, (run_at, hostnames_json, file_path, action, flags))

    def due_schedules(self):
        with self._lock:
            return list(self.conn.execute(
                "SELECT * FROM scheduled WHERE done=0 AND run_at<=? ORDER BY run_at",
                (time.time(),),
            ))

    def mark_schedule_done(self, sid: int):
        with self._lock:
            self._execute_write("UPDATE scheduled SET done=1 WHERE id=?", (sid,))

    def remove_schedule(self, sid: int):
        with self._lock:
            self._execute_write("DELETE FROM scheduled WHERE id=?", (sid,))
