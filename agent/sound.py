from __future__ import annotations

import ctypes
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
from pathlib import Path

log = logging.getLogger("agent.sound")

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

try:
    import winsound
except Exception:
    winsound = None


def _coinit():
    try:
        if not hasattr(sys, "coinit_flags"):
            sys.coinit_flags = 0
        ctypes.windll.ole32.CoInitializeEx(None, int(sys.coinit_flags))
    except Exception:
        pass


def _make_volume():
    _coinit()
    try:
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        speakers = AudioUtilities.GetSpeakers()
        if not speakers:
            return None
        if hasattr(speakers, "Activate"):
            iface = speakers.Activate(IAudioEndpointVolume._iid_, 0, None)
            return iface.QueryInterface(IAudioEndpointVolume)
        endpoint = getattr(speakers, "EndpointVolume", None)
        if endpoint is not None:
            return endpoint
        endpoint = getattr(speakers, "endpoint_volume", None)
        if endpoint is not None:
            return endpoint
        log.warning("Unsupported pycaw speaker object: %s", type(speakers).__name__)
        return None
    except ImportError:
        log.warning("pycaw is not installed, sound control unavailable")
        return None
    except Exception as e:
        log.warning("Sound initialization error: %s", e)
        return None


def _ps_quote(value: str) -> str:
    return "'" + str(value or "").replace("'", "''") + "'"


def _powershell_encoded(script: str) -> str:
    return script.encode("utf-16le").hex()


class SoundControl:
    def __init__(self):
        self._vol = _make_volume()
        self._play_lock = threading.RLock()
        self._play_proc: subprocess.Popen | None = None
        self._backend = ""
        self._cleanup_path: str | None = None
        self._tts_thread: threading.Thread | None = None
        self._tts_proc: subprocess.Popen | None = None
        self._tts_stop = threading.Event()
        if self._vol:
            log.info("Sound control ready")

    def _v(self):
        if self._vol is None:
            self._vol = _make_volume()
        return self._vol

    def set_volume(self, level: float):
        v = self._v()
        if not v:
            return
        try:
            v.SetMasterVolumeLevelScalar(max(0.0, min(1.0, level)), None)
        except Exception:
            self._vol = None

    def get_volume(self) -> float:
        v = self._v()
        if not v:
            return 0.5
        try:
            return float(v.GetMasterVolumeLevelScalar())
        except Exception:
            self._vol = None
            return 0.5

    def mute(self):
        v = self._v()
        if not v:
            return
        try:
            v.SetMute(1, None)
        except Exception:
            self._vol = None

    def unmute(self):
        v = self._v()
        if not v:
            return
        try:
            v.SetMute(0, None)
        except Exception:
            self._vol = None

    def is_muted(self) -> bool:
        v = self._v()
        if not v:
            return False
        try:
            return bool(v.GetMute())
        except Exception:
            self._vol = None
            return False

    def _remember_proc(self, proc: subprocess.Popen | None, backend: str, cleanup_path: str | None):
        with self._play_lock:
            self._play_proc = proc
            self._backend = backend
            self._cleanup_path = cleanup_path

    def _clear_proc_state(self):
        with self._play_lock:
            self._play_proc = None
            self._backend = ""
            self._cleanup_path = None

    def _cleanup_file(self, path: str | None):
        if not path:
            return
        for _ in range(12):
            try:
                os.remove(path)
                return
            except FileNotFoundError:
                return
            except PermissionError:
                time.sleep(0.25)
            except Exception:
                time.sleep(0.15)

    def _finish_playback(self):
        with self._play_lock:
            cleanup_path = self._cleanup_path
        self._clear_proc_state()
        self._cleanup_file(cleanup_path)

    def _kill_known_players(self):
        # List of common media players and audio applications
        media_players = (
            # Windows built-in players
            "wmplayer.exe",
            "Music.UI.exe",
            "Microsoft.Media.Player.exe",
            "GrooveMusic.exe",
            "ApplicationFrameHost.exe",
            # Third-party media players
            "vlc.exe",
            "mpv.exe",
            "mpc-hc.exe",
            "mpc-hc64.exe",
            "foobar2000.exe",
            "winamp.exe",
            "mediaplayerclassic.exe",
            "potplayer.exe",
            "gom.exe",
            # Streaming and communication apps
            "spotify.exe",
            "spotifyplayer.exe",
            "audacious.exe",
            "itunes.exe",
            "mediamonkey.exe",
            # Browser-based audio
            "firefox.exe",
            "chrome.exe",
            "msedge.exe",
        )
        for image in media_players:
            try:
                subprocess.run(
                    ["taskkill", "/F", "/IM", image],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=2,
                    creationflags=CREATE_NO_WINDOW,
                )
            except Exception:
                pass

    def stop_playback(self):
        with self._play_lock:
            proc = self._play_proc
            backend = self._backend
            cleanup_path = self._cleanup_path
            self._play_proc = None
            self._backend = ""
            self._cleanup_path = None

        try:
            if winsound is not None:
                winsound.PlaySound(None, 0)
        except Exception:
            pass

        if proc is not None:
            try:
                if proc.poll() is None:
                    proc.terminate()
                    try:
                        proc.wait(timeout=3)
                    except Exception:
                        proc.kill()
            except Exception as e:
                log.debug("stop_playback tracked proc: %s", e)

        # Always kill all known media players to ensure any playing audio is stopped
        self._kill_known_players()
        self._stop_tts()
        self._cleanup_file(cleanup_path)

    def play_file(self, path: str, volume: float | None = None, cleanup_path: str | None = None) -> bool:
        path = os.fspath(path or "").strip()
        if not path:
            return False
        if not os.path.exists(path):
            log.warning("Audio file not found: %s", path)
            return False

        self.stop_playback()
        if volume is not None:
            try:
                self.set_volume(float(volume))
            except Exception:
                pass
        try:
            self.unmute()
        except Exception:
            pass

        ext = Path(path).suffix.lower()
        if ext == ".wav" and self._play_wav_async(path, cleanup_path):
            log.info("audio backend: winsound (%s)", path)
            return True
        if self._play_dotnet_media(path, cleanup_path):
            log.info("audio backend: dotnet_media (%s)", path)
            return True
        if self._play_wmplayer_exe(path, cleanup_path):
            log.info("audio backend: wmplayer_exe (%s)", path)
            return True
        if self._play_shell_open(path, cleanup_path):
            log.info("audio backend: shell_open (%s)", path)
            return True
        if self._play_powershell_start(path, cleanup_path):
            log.info("audio backend: powershell_start (%s)", path)
            return True

        log.warning("all audio backends failed: %s", path)
        self._cleanup_file(cleanup_path)
        return False

    def _play_wav_async(self, path: str, cleanup_path: str | None) -> bool:
        if winsound is None:
            return False

        def worker():
            try:
                winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_SYNC)
            except Exception as e:
                log.warning("winsound playback failed: %s", e)
            finally:
                self._finish_playback()

        th = threading.Thread(target=worker, daemon=True, name="SoundWavPlayer")
        self._remember_proc(None, "winsound", cleanup_path)
        th.start()
        return True

    def _play_dotnet_media(self, path: str, cleanup_path: str | None) -> bool:
        if os.name != "nt":
            return False
        status_file = os.path.join(tempfile.gettempdir(), f"classdeploy_audio_{uuid.uuid4().hex}.status")
        file_uri = Path(path).resolve().as_uri()
        script = f"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
$statusFile = {_ps_quote(status_file)}
$uriText = {_ps_quote(file_uri)}
$player = New-Object System.Windows.Media.MediaPlayer
$script:started = $false
$script:stopped = $false
$dispatcher = [System.Windows.Threading.Dispatcher]::CurrentDispatcher

$player.add_MediaOpened({{
    if (-not $script:started) {{
        $script:started = $true
        Set-Content -LiteralPath $statusFile -Value 'STARTED' -Encoding ASCII -Force
    }}
}})

$player.add_MediaEnded({{
    $script:stopped = $true
    $dispatcher.BeginInvokeShutdown([System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
}})

$player.add_MediaFailed({{
    param($sender, $eventArgs)
    $msg = 'FAILED'
    try {{
        if ($eventArgs -and $eventArgs.ErrorException) {{
            $msg = 'FAILED:' + $eventArgs.ErrorException.Message
        }}
    }} catch {{}}
    Set-Content -LiteralPath $statusFile -Value $msg -Encoding UTF8 -Force
    $script:stopped = $true
    $dispatcher.BeginInvokeShutdown([System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
}})

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(200)
$deadline = (Get-Date).AddSeconds(10)
$timer.add_Tick({{
    if (-not $script:started -and (Get-Date) -gt $deadline) {{
        if (-not (Test-Path -LiteralPath $statusFile)) {{
            Set-Content -LiteralPath $statusFile -Value 'FAILED:dotnet start timeout' -Encoding UTF8 -Force
        }}
        $script:stopped = $true
        $dispatcher.BeginInvokeShutdown([System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
    }}
}})
$timer.Start()

$player.Open([System.Uri]$uriText)
$player.Volume = 1.0
$player.Play()
[System.Windows.Threading.Dispatcher]::Run()

try {{ $timer.Stop() }} catch {{}}
try {{ $player.Stop() }} catch {{}}
try {{ $player.Close() }} catch {{}}
"""
        encoded = script.encode("utf-16le")
        import base64
        encoded_b64 = base64.b64encode(encoded).decode("ascii")
        try:
            proc = subprocess.Popen(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Sta",
                    "-EncodedCommand",
                    encoded_b64,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW,
            )
        except Exception as e:
            log.warning("dotnet media playback failed: %s", e)
            self._cleanup_file(status_file)
            return False

        deadline = time.time() + 12.0
        failure = None
        started = False
        while time.time() < deadline:
            if os.path.exists(status_file):
                try:
                    marker = Path(status_file).read_text(encoding="utf-8", errors="replace").strip()
                except Exception:
                    marker = ""
                if marker == "STARTED":
                    started = True
                    break
                if marker.startswith("FAILED"):
                    failure = marker
                    break
            code = proc.poll()
            if code is not None:
                failure = f"powershell exited with code {code}"
                break
            time.sleep(0.2)

        if not started:
            try:
                if proc.poll() is None:
                    proc.terminate()
            except Exception:
                pass
            log.warning("dotnet media playback failed: %s", failure or "start timeout")
            self._cleanup_file(status_file)
            return False

        def waiter():
            try:
                proc.wait(timeout=60 * 60)
            except Exception:
                pass
            finally:
                self._cleanup_file(status_file)
                self._finish_playback()

        th = threading.Thread(target=waiter, daemon=True, name="SoundDotNetWaiter")
        self._remember_proc(proc, "dotnet_media", cleanup_path)
        th.start()
        return True

    def _find_wmplayer(self) -> str | None:
        candidates = [
            shutil.which("wmplayer.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), "Windows Media Player", "wmplayer.exe"),
            os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "Windows Media Player", "wmplayer.exe"),
        ]
        for item in candidates:
            if item and os.path.exists(item):
                return item
        return None

    def _play_wmplayer_exe(self, path: str, cleanup_path: str | None) -> bool:
        wmplayer = self._find_wmplayer()
        if not wmplayer:
            return False
        try:
            proc = subprocess.Popen(
                [wmplayer, "/play", "/close", path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW,
            )
            time.sleep(0.8)
            if proc.poll() not in (None, 0):
                log.warning("wmplayer.exe playback failed: exit code %s", proc.poll())
                return False

            def waiter():
                try:
                    proc.wait(timeout=60 * 60)
                except Exception:
                    pass
                finally:
                    self._finish_playback()

            th = threading.Thread(target=waiter, daemon=True, name="SoundWmpExeWaiter")
            self._remember_proc(proc, "wmplayer_exe", cleanup_path)
            th.start()
            return True
        except Exception as e:
            log.warning("wmplayer.exe playback failed: %s", e)
            return False

    def _play_shell_open(self, path: str, cleanup_path: str | None) -> bool:
        if os.name != "nt" or not hasattr(os, "startfile"):
            return False
        try:
            os.startfile(path)
            self._remember_proc(None, "shell_open", cleanup_path)
            return True
        except Exception as e:
            log.warning("shell open playback failed: %s", e)
            return False

    def _play_powershell_start(self, path: str, cleanup_path: str | None) -> bool:
        safe = path.replace("'", "''")
        cmd = [
            "powershell",
            "-NoProfile",
            "-WindowStyle",
            "Hidden",
            "-Command",
            f"Start-Process -FilePath '{safe}'",
        ]
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW,
            )
            rc = proc.wait(timeout=5)
            if rc != 0:
                log.warning("powershell start playback failed: exit code %s", rc)
                return False
            self._remember_proc(None, "powershell_start", cleanup_path)
            return True
        except Exception as e:
            log.warning("powershell start playback failed: %s", e)
            return False

    def _stop_tts(self):
        proc = None
        with self._play_lock:
            self._tts_stop.set()
            proc = self._tts_proc
            self._tts_proc = None
        if proc is not None:
            try:
                if proc.poll() is None:
                    proc.terminate()
                    try:
                        proc.wait(timeout=2)
                    except Exception:
                        proc.kill()
            except Exception:
                pass

    def speak_text(self, text: str, volume: float = 1.0, rate: int | None = None) -> bool:
        text = (text or "").strip()
        if not text:
            return False
        volume = max(0.0, min(1.0, float(volume or 1.0)))
        rate = -3 if rate is None else max(-10, min(10, int(rate)))

        self.stop_playback()
        try:
            self.unmute()
        except Exception:
            pass
        try:
            self.set_volume(volume)
        except Exception:
            pass

        stop_evt = threading.Event()
        with self._play_lock:
            self._tts_stop = stop_evt

        def worker():
            try:
                import pyttsx3
                engine = pyttsx3.init()
                try:
                    engine.setProperty("rate", 110)
                except Exception:
                    pass
                try:
                    engine.setProperty("volume", volume)
                except Exception:
                    pass
                engine.say(text)
                engine.runAndWait()
                log.info("speak_text completed: %s", text[:50])
                return
            except Exception as e:
                log.debug("speak_text pyttsx3 failed: %s", e)
            self._speak_via_powershell(text, volume, rate)

        th = threading.Thread(target=worker, daemon=True, name="SpeakThread")
        with self._play_lock:
            self._tts_thread = th
        th.start()
        return True

    def _speak_via_powershell(self, text: str, volume: float = 1.0, rate: int = -3) -> bool:
        text = (text or "").strip()
        if not text:
            return False
        safe_text = text.replace("'", "''")
        volume_int = max(0, min(100, int(volume * 100)))
        ps_code = (
            "$speak = New-Object -ComObject SAPI.SPVoice; "
            f"$speak.Volume = {volume_int}; "
            f"$speak.Rate = {rate}; "
            f"[void]$speak.Speak('{safe_text}')"
        )
        try:
            proc = subprocess.Popen(
                ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_code],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW,
            )
            with self._play_lock:
                self._tts_proc = proc
            log.info("speak_text via PowerShell: %s", text[:50])
            return True
        except Exception as e:
            log.warning("speak_via_powershell: %s", e)
            return False
