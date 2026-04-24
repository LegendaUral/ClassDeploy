"""
Захват экрана учебного ПК.
Использует mss для быстрого захвата и Pillow для JPEG.
Работает в фоновом потоке, шлёт кадры через колбэк.
"""
from __future__ import annotations
import io
import time
import base64
import logging
import threading
from typing import Callable, Optional

log = logging.getLogger("agent.screen")


class ScreenStreamer:
    def __init__(
        self,
        send_frame: Callable[[str, int, int], None],
        fps: int = 10,
        quality: int = 75,
        max_width: int = 1280,
    ):
        self.send_frame = send_frame
        self.fps = max(1, fps)
        self.quality = max(1, min(100, quality))
        self.max_width = max(320, max_width)
        self._thread: Optional[threading.Thread] = None
        self._stop = threading.Event()

    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True, name="ScreenStreamer")
        self._thread.start()
        log.info("Трансляция экрана запущена (fps=%d, quality=%d)", self.fps, self.quality)

    def stop(self):
        self._stop.set()
        log.info("Трансляция экрана остановлена")

    def _loop(self):
        # Импорты внутри потока — не роняют агент если библиотек нет
        try:
            import mss
            from PIL import Image
        except ImportError as e:
            log.error("mss/Pillow не установлены, трансляция невозможна: %s", e)
            return

        period = 1.0 / self.fps
        consecutive_errors = 0
        black_streak = 0

        with mss.mss() as sct:
            monitor = sct.monitors[1]  # основной монитор
            while not self._stop.is_set():
                t0 = time.monotonic()
                try:
                    raw = sct.grab(monitor)
                    img = Image.frombytes("RGB", raw.size, raw.rgb)

                    # Диагностика чёрных кадров (session 0 без доступа к рабочему столу)
                    if img.getbbox() is None:
                        black_streak += 1
                        if black_streak == 20:
                            log.warning("Кадры пустые (session 0?), пробую ImageGrab fallback")
                    else:
                        black_streak = 0

                    if black_streak > 20:
                        try:
                            from PIL import ImageGrab
                            fb = ImageGrab.grab()
                            if fb.getbbox() is not None:
                                img = fb
                                black_streak = 0
                        except Exception:
                            pass

                    w, h = img.size
                    if w > self.max_width:
                        scale = self.max_width / w
                        img = img.resize(
                            (int(w * scale), int(h * scale)),
                            Image.Resampling.BILINEAR,
                        )

                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=self.quality, optimize=False)
                    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                    self.send_frame(b64, img.size[0], img.size[1])
                    consecutive_errors = 0

                except Exception as e:
                    consecutive_errors += 1
                    if consecutive_errors <= 3:
                        log.debug("Ошибка кадра (#%d): %s", consecutive_errors, e)
                    if consecutive_errors > 30:
                        log.error("Слишком много ошибок кадра, останавливаю трансляцию")
                        break

                elapsed = time.monotonic() - t0
                sleep_for = period - elapsed
                if sleep_for > 0:
                    self._stop.wait(sleep_for)
