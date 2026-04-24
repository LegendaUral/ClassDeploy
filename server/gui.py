"""
GUI сервера ClassDeploy (PyQt6).
Вкладки:
  📦 Установка   — список ПК, drag-and-drop, быстрые кнопки
  🖥 Экраны      — сетка трансляций
  ⚙ Управление  — блокировка, сообщения, питание, звук, клавиши
  🚫 Ограничения — домены и приложения
"""
from __future__ import annotations
import os
import sys
import json
import base64
import asyncio
import logging
from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import Qt, QTimer, QObject, QDateTime, pyqtSignal
from PyQt6.QtGui import QColor, QPixmap, QFont, QIcon
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QLabel, QPushButton,
    QFileDialog, QMessageBox, QCheckBox, QHeaderView,
    QLineEdit, QSplitter, QTextEdit, QInputDialog,
    QDialog, QDialogButtonBox, QDateTimeEdit,
    QTabWidget, QListWidget, QGroupBox,
)

sys.path.insert(0, str(Path(__file__).parent.parent))

from shared import protocol as P
from server.network import Server
from server.database import Database
from server.screen_widgets import ScreenGrid, FullScreenView

log = logging.getLogger("server.gui")


# ── Сигнальный мост asyncio → Qt ────────────────────────────────

class QtBridge(QObject):
    agents_changed = pyqtSignal()
    job_update     = pyqtSignal(str, str, str)
    screen_frame   = pyqtSignal(str, str, int, int)


# ── Зона drag-and-drop ───────────────────────────────────────────

class DropZone(QLabel):
    file_dropped = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumHeight(90)
        self._reset()

    def _reset(self):
        self.setText("⬇  Перетащите .msi / .exe / .zip  ⬇\nили нажмите для выбора файла")
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #555;  border-radius:8px;
                font-size:12pt; color:#888; background:#f9f9f9; padding:8px;
            }
        """)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self.setStyleSheet("""
                QLabel {
                    border:3px solid #2a9; border-radius:8px;
                    font-size:12pt; color:#2a9; background:#eafaf2; padding:8px;
                }
            """)

    def dragLeaveEvent(self, e): self._reset()

    def dropEvent(self, e):
        urls = e.mimeData().urls()
        if urls:
            self.file_dropped.emit(urls[0].toLocalFile())
        self._reset()

    def mousePressEvent(self, e):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл", "",
            "Установщики (*.msi *.exe *.zip);;Все файлы (*.*)",
        )
        if path:
            self.file_dropped.emit(path)


# ── Диспетчер задач ─────────────────────────────────────────────

class ProcessManagerDialog(QDialog):
    def __init__(self, parent, server: Server, loop, hostname: str):
        super().__init__(parent)
        self.server   = server
        self.loop     = loop
        self.hostname = hostname
        self.setWindowTitle(f"Процессы — {hostname}")
        self.resize(1000, 580)

        lay = QVBoxLayout(self)

        # Панель кнопок + поиск
        top = QHBoxLayout()
        self.btn_ref  = QPushButton("🔄 Обновить")
        self.btn_ref.clicked.connect(self.refresh)
        self.btn_kill = QPushButton("⛔ Завершить")
        self.btn_kill.clicked.connect(self.kill_selected)
        self.btn_kill.setStyleSheet("QPushButton { color:#c0392b; font-weight:bold; }")
        self.search = QLineEdit()
        self.search.setPlaceholderText("Поиск по имени или пути…")
        self.search.textChanged.connect(self._filter)
        top.addWidget(self.btn_ref)
        top.addWidget(self.btn_kill)
        top.addStretch()
        top.addWidget(QLabel("🔍"))
        top.addWidget(self.search)
        lay.addLayout(top)

        # Таблица процессов
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["PID", "Имя процесса", "Сессия", "Память", "Путь к файлу"])
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().hide()
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(self._on_row_double_click)
        lay.addWidget(self.table)

        self.status = QLabel("Загрузка…")
        self.status.setStyleSheet("color:#666; font-size:10px;")
        lay.addWidget(self.status)

        self._all_rows: list[list] = []
        self.refresh()

    def refresh(self):
        self.setEnabled(False)
        self.status.setText("Получаю список процессов…")
        QApplication.processEvents()
        try:
            fut    = asyncio.run_coroutine_threadsafe(
                self.server.get_process_list(self.hostname), self.loop
            )
            result = fut.result(timeout=25)
        except Exception as e:
            self.status.setText(f"Ошибка: {e}")
            self.setEnabled(True)
            return

        if not result.get("ok"):
            self.status.setText(result.get("message", "Ошибка"))
            self.setEnabled(True)
            return

        self._all_rows = []
        for p in result.get("processes", []):
            self._all_rows.append([
                p.get("pid", 0),
                str(p.get("name", "")),
                str(p.get("session", "")),
                str(p.get("memory", "")),
                str(p.get("path", "")),
            ])

        self._populate(self._all_rows)
        self.status.setText(f"Найдено процессов: {len(self._all_rows)}")
        self.setEnabled(True)

    def _populate(self, rows: list):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        for row in rows:
            r = self.table.rowCount()
            self.table.insertRow(r)

            pid_item = QTableWidgetItem()
            pid_item.setData(Qt.ItemDataRole.DisplayRole, int(row[0]))
            self.table.setItem(r, 0, pid_item)

            name_item = QTableWidgetItem(row[1])
            path_str  = row[4]
            if path_str:
                name_item.setToolTip(path_str)
            self.table.setItem(r, 1, name_item)

            self.table.setItem(r, 2, QTableWidgetItem(row[2]))

            # Память — выровнять вправо
            mem_item = QTableWidgetItem(row[3])
            mem_item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(r, 3, mem_item)

            # Путь — серый шрифт
            path_item = QTableWidgetItem(path_str)
            path_item.setForeground(QColor("#888"))
            self.table.setItem(r, 4, path_item)

        self.table.setSortingEnabled(True)

    def _filter(self, text: str):
        t = text.lower()
        filtered = [
            r for r in self._all_rows
            if not t or t in r[1].lower() or t in r[4].lower()
        ]
        self._populate(filtered)
        self.status.setText(f"Показано: {len(filtered)} / {len(self._all_rows)}")

    def _on_row_double_click(self, idx):
        """Двойной клик — показать полный путь."""
        path_item = self.table.item(idx.row(), 4)
        if path_item and path_item.text():
            QMessageBox.information(self, "Путь к файлу", path_item.text())

    def kill_selected(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            QMessageBox.warning(self, "Выбор", "Выберите процесс в таблице.")
            return
        r   = rows[0].row()
        pid = int(self.table.item(r, 0).data(Qt.ItemDataRole.DisplayRole))
        nm  = self.table.item(r, 1).text() if self.table.item(r, 1) else ""

        if QMessageBox.question(
            self, "Завершить процесс",
            f"Завершить:\n{nm}  (PID {pid})?",
        ) != QMessageBox.StandardButton.Yes:
            return

        self.setEnabled(False)
        QApplication.processEvents()
        try:
            fut    = asyncio.run_coroutine_threadsafe(
                self.server.kill_process(self.hostname, pid), self.loop
            )
            result = fut.result(timeout=20)
            if result.get("ok"):
                QMessageBox.information(self, "Готово", result.get("message", "Процесс завершён"))
            else:
                QMessageBox.warning(self, "Ошибка", result.get("message", "Не удалось завершить"))
            self.refresh()
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", str(e))
        finally:
            self.setEnabled(True)


# ── Главное окно ─────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self, server: Server, db: Database, bridge: QtBridge, loop):
        super().__init__()
        self.server = server
        self.db     = db
        self.bridge = bridge
        self.loop   = loop

        self.blocked_domains: set[str]  = set()
        self.blocked_apps:    set[str]  = set()
        self._screamer_b64:   str       = ""
        self._current_file:   str       = ""
        self.fullscreen_view: FullScreenView | None = None

        self.setWindowTitle("ClassDeploy - управление учебным классом")
        self.resize(1320, 860)

        self._build_ui()
        self._apply_theme()

        self.bridge.agents_changed.connect(self.refresh_agents)
        self.bridge.job_update.connect(self.on_job_update)
        self.bridge.screen_frame.connect(self.on_screen_frame)

        # Обновление списка агентов каждые 2 сек
        self._agent_timer = QTimer(self)
        self._agent_timer.timeout.connect(self.refresh_agents)
        self._agent_timer.start(2000)

        self.refresh_agents()

    # ── Построение UI ─────────────────────────────────────────────

    def _build_ui(self):
        cw   = QWidget()
        root = QVBoxLayout(cw)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(10)

        hdr_wrap = QVBoxLayout()
        hdr = QHBoxLayout()
        title_box = QVBoxLayout()
        lbl = QLabel("ClassDeploy")
        lbl.setObjectName("appTitle")
        sub = QLabel("Панель учителя для установки, трансляции экранов и удаленного управления")
        sub.setObjectName("appSubtitle")
        title_box.addWidget(lbl)
        title_box.addWidget(sub)
        hdr.addLayout(title_box)
        hdr.addStretch()

        self.stat_label = QLabel("Онлайн: 0 / 0")
        self.stat_label.setObjectName("topStat")
        hdr.addWidget(self.stat_label)
        hdr_wrap.addLayout(hdr)

        cards = QHBoxLayout()
        self.card_online = QLabel("Онлайн\n0")
        self.card_online.setObjectName("summaryCard")
        self.card_selected = QLabel("Выбрано\n0")
        self.card_selected.setObjectName("summaryCard")
        self.card_file = QLabel("Текущий файл\nНе выбран")
        self.card_file.setObjectName("summaryCardWide")
        cards.addWidget(self.card_online)
        cards.addWidget(self.card_selected)
        cards.addWidget(self.card_file, stretch=1)
        hdr_wrap.addLayout(cards)
        root.addLayout(hdr_wrap)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._tab_install(),  "📦 Установка")
        self.tabs.addTab(self._tab_screens(),  "🖥 Экраны")
        self.tabs.addTab(self._tab_control(),  "⚙ Управление")
        self.tabs.addTab(self._tab_restrict(), "🚫 Ограничения")
        root.addWidget(self.tabs)

        self.setCentralWidget(cw)

    def _apply_theme(self):
        self.setStyleSheet("""
            QWidget { background:#f4f7fb; color:#1f2937; }
            #appTitle { font-size:24px; font-weight:700; color:#0f172a; }
            #appSubtitle { color:#64748b; font-size:12px; }
            #topStat { background:#e2e8f0; border-radius:10px; padding:8px 12px; font-weight:600; }
            #summaryCard, #summaryCardWide {
                background:white; border:1px solid #dbe3ee; border-radius:14px;
                padding:12px 14px; font-size:14px; font-weight:600;
            }
            QTabWidget::pane { border:1px solid #dbe3ee; border-radius:14px; background:white; }
            QTabBar::tab {
                background:#e9eef6; border:1px solid #dbe3ee; border-bottom:none;
                padding:10px 16px; margin-right:4px; border-top-left-radius:10px; border-top-right-radius:10px;
                color:#334155; font-weight:600;
            }
            QTabBar::tab:selected { background:white; color:#0f172a; }
            QPushButton {
                background:#2563eb; color:white; border:none; border-radius:10px;
                padding:8px 12px; font-weight:600;
            }
            QPushButton:hover { background:#1d4ed8; }
            QPushButton:disabled { background:#cbd5e1; color:#64748b; }
            QLineEdit, QTextEdit, QListWidget, QDateTimeEdit {
                border:1px solid #cbd5e1; border-radius:10px; padding:6px 8px; background:white;
            }
            QTableWidget {
                border:1px solid #dbe3ee; border-radius:12px; gridline-color:#edf2f7;
                background:white; alternate-background-color:#f8fafc;
            }
            QHeaderView::section {
                background:#eef2ff; color:#334155; padding:8px; border:none; border-bottom:1px solid #dbe3ee;
                font-weight:700;
            }
            QGroupBox {
                font-weight:700; border:1px solid #dbe3ee; border-radius:12px; margin-top:12px; padding-top:12px;
                background:#fbfdff;
            }
            QGroupBox::title { subcontrol-origin: margin; left:12px; padding:0 4px; color:#0f172a; }
            QScrollArea { border:none; }
        """)

    def _update_summary_cards(self):
        total = self.agents_table.rowCount() if hasattr(self, 'agents_table') else 0
        online = len(self.server.agents)
        selected = len(self._checked_hosts()) if hasattr(self, 'agents_table') else 0
        current_file = Path(self._current_file).name if self._current_file else 'Не выбран'
        self.card_online.setText(f"Онлайн\n{online} из {total}")
        self.card_selected.setText(f"Выбрано\n{selected}")
        self.card_file.setText(f"Текущий файл\n{current_file}")

    def _filter_agents_table(self):
        if not hasattr(self, 'agents_table'):
            return
        query = self.agent_filter.text().strip().lower() if hasattr(self, 'agent_filter') else ''
        for row in range(self.agents_table.rowCount()):
            values = []
            for col in range(1, self.agents_table.columnCount()):
                item = self.agents_table.item(row, col)
                values.append(item.text().lower() if item else '')
            hidden = bool(query) and not any(query in value for value in values)
            self.agents_table.setRowHidden(row, hidden)
        self._update_summary_cards()

    # ── Вкладка: Установка ────────────────────────────────────────

    def _tab_install(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        sp  = QSplitter(Qt.Orientation.Horizontal)

        # Левая: список ПК
        left  = QWidget()
        ll    = QVBoxLayout(left)
        ll.addWidget(QLabel("<b>Учебные компьютеры</b>"))
        hint = QLabel("Выберите нужные ПК, затем запустите установку или управление справа.")
        hint.setStyleSheet("color:#64748b; font-size:11px;")
        ll.addWidget(hint)

        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Поиск:"))
        self.agent_filter = QLineEdit()
        self.agent_filter.setPlaceholderText("Хост, IP, ОС или статус")
        self.agent_filter.textChanged.connect(self._filter_agents_table)
        filter_row.addWidget(self.agent_filter)
        ll.addLayout(filter_row)

        sel_row = QHBoxLayout()
        for lbl, fn in [("Все", lambda: self._set_checks(True)),
                        ("Снять", lambda: self._set_checks(False)),
                        ("Онлайн", self._select_online)]:
            b = QPushButton(lbl)
            b.clicked.connect(fn)
            b.setFixedHeight(26)
            sel_row.addWidget(b)
        ll.addLayout(sel_row)

        self.agents_table = QTableWidget(0, 5)
        self.agents_table.setHorizontalHeaderLabels(["✓", "Хост", "IP", "ОС", "Статус"])
        h = self.agents_table.horizontalHeader()
        h.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        h.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.agents_table.verticalHeader().hide()
        self.agents_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.agents_table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        ll.addWidget(self.agents_table)
        sp.addWidget(left)

        # Правая: действия
        right = QWidget()
        rl    = QVBoxLayout(right)

        self.drop = DropZone()
        self.drop.file_dropped.connect(self.on_file_dropped)
        rl.addWidget(self.drop)

        # Флаги + форс автокликер
        opt = QHBoxLayout()
        opt.addWidget(QLabel("Флаги:"))
        self.flags_edit = QLineEdit()
        self.flags_edit.setPlaceholderText("/S  /VERYSILENT  /quiet …")
        opt.addWidget(self.flags_edit)
        self.chk_force_ac = QCheckBox("Форс автокликер")
        opt.addWidget(self.chk_force_ac)
        rl.addLayout(opt)

        # Быстрые клики по кнопкам установщика
        grp_q = QGroupBox("Быстрые кнопки установщика")
        ql    = QHBoxLayout(grp_q)
        for lbl in ["Далее", "Next", "Принять", "I Agree",
                    "Установить", "Install", "Готово", "Finish"]:
            b = QPushButton(lbl)
            b.setFixedHeight(26)
            b.clicked.connect(lambda _, t=lbl: self._quick_click(t))
            ql.addWidget(b)
        rl.addWidget(grp_q)

        # Кнопки действий
        act = QHBoxLayout()
        self.btn_install = QPushButton("📦 Установить")
        self.btn_install.setStyleSheet("font-weight:bold; padding:6px 14px;")
        self.btn_install.clicked.connect(self.on_install)
        act.addWidget(self.btn_install)

        self.btn_uninstall = QPushButton("🗑 Удалить…")
        self.btn_uninstall.clicked.connect(self.on_uninstall)
        act.addWidget(self.btn_uninstall)

        self.btn_push = QPushButton("📁 Отправить файл…")
        self.btn_push.clicked.connect(self._on_push_file)
        act.addWidget(self.btn_push)

        self.btn_schedule = QPushButton("📅 Запланировать…")
        self.btn_schedule.clicked.connect(self.on_schedule)
        act.addWidget(self.btn_schedule)
        rl.addLayout(act)

        # Прогресс
        rl.addWidget(QLabel("<b>Ход установки:</b>"))
        self.jobs_table = QTableWidget(0, 4)
        self.jobs_table.setHorizontalHeaderLabels(["Хост", "Файл", "Статус", "Сообщение"])
        self.jobs_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.jobs_table.verticalHeader().hide()
        self.jobs_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.jobs_table.setMaximumHeight(180)
        rl.addWidget(self.jobs_table)

        # Журнал
        rl.addWidget(QLabel("<b>Журнал:</b>"))
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFont(QFont("Consolas", 9))
        self.log_view.setMaximumHeight(140)
        rl.addWidget(self.log_view)

        sp.addWidget(right)
        sp.setSizes([300, 720])
        lay.addWidget(sp)
        return w

    # ── Вкладка: Экраны ───────────────────────────────────────────

    def _tab_screens(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(4, 4, 4, 4)

        btns = QHBoxLayout()
        for lbl, fn in [
            ("▶ Запустить (выбранные)", self._scr_start),
            ("⏹ Стоп (выбранные)",      self._scr_stop),
            ("▶▶ Все онлайн",            self._scr_start_all),
            ("⏹⏹ Стоп всех",             self._scr_stop_all),
        ]:
            b = QPushButton(lbl)
            b.clicked.connect(fn)
            btns.addWidget(b)
        btns.addStretch()
        lay.addLayout(btns)

        self.screen_grid = ScreenGrid()
        self.screen_grid.tile_double_clicked.connect(self._open_fullscreen)
        lay.addWidget(self.screen_grid, stretch=1)

        hint = QLabel(
            "Двойной клик по плитке — полноэкранный просмотр."
            "  В режиме управления доступна только мышь, клавиатура в трансляции отключена."
        )
        hint.setStyleSheet("color:#999; font-size:10px; padding:2px;")
        lay.addWidget(hint)
        return w

    # ── Вкладка: Управление ───────────────────────────────────────

    def _tab_control(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(6)

        def grp(title, items):
            g = QGroupBox(title)
            l = QHBoxLayout(g)
            for lbl, fn in items:
                b = QPushButton(lbl)
                b.clicked.connect(fn)
                l.addWidget(b)
            l.addStretch()
            return g

        lay.addWidget(grp("🔒 Экран ученика", [
            ("🔒 Заблокировать",  self._on_lock),
            ("🔓 Разблокировать", self._on_unlock),
            ("💬 Сообщение",      self._on_msg),
            ("😱 Скример",        self._on_screamer),
            ("🖼 Картинка скримера…", self._load_screamer_img),
        ]))

        lay.addWidget(grp("⏻ Питание", [
            ("⏻ Выключить",       lambda: self._on_power("shutdown")),
            ("🔄 Перезагрузить",   lambda: self._on_power("reboot")),
            ("👤 Завершить сеанс", lambda: self._on_power("logoff")),
            ("🔐 Lock Windows",    lambda: self._on_power("lock")),
        ]))

        lay.addWidget(grp("🔊 Звук", [
            ("🔇 Заглушить",    self._on_mute),
            ("🔊 Включить",     self._on_unmute),
            ("🎚 Уровень…",     self._on_volume),
            ("🎵 Проиграть файл…", self._on_play_audio),
            ("🗣 Озвучить текст…", self._on_speak_text),
            ("⏹ Стоп аудио",    self._on_stop_audio),
        ]))

        lay.addWidget(grp("▶ Программы", [
            ("🧑‍💻 VS Code (всем)", self._on_vscode_all),
            ("▶ Запустить…",       self._on_run_prog),
            ("🔍 Диспетчер задач", self._on_task_mgr),
        ]))

        # Кнопка клика по установщику
        lay.addWidget(grp("🖱 Установщик", [
            ("🖱 Нажать кнопку…", self._on_installer_click_dlg),
        ]))

        # Клавиатурные комбинации
        grp_kbd = QGroupBox("⌨ Клавиатурные комбинации на ПК ученика")
        kl = QHBoxLayout(grp_kbd)
        for lbl, vks in [
            ("Win+D",       [0x5B, 0x44]),
            ("Win+E",       [0x5B, 0x45]),
            ("Win+L",       [0x5B, 0x4C]),
            ("Alt+F4",      [0x12, 0x73]),
            ("Ctrl+Alt+Del",[0x11, 0x12, 0x2E]),
            ("PrintScr",    [0x2C]),
            ("Ctrl+W",      [0x11, 0x57]),
            ("Ctrl+Shift+Esc",[0x11, 0x10, 0x1B]),
        ]:
            b = QPushButton(lbl)
            b.setFixedHeight(26)
            b.clicked.connect(lambda _, v=vks: self._send_combo(v))
            kl.addWidget(b)
        kl.addStretch()
        lay.addWidget(grp_kbd)

        lay.addStretch()
        return w

    # ── Вкладка: Ограничения ──────────────────────────────────────

    def _tab_restrict(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(8)

        def make_grp(title, btn_items, list_attr):
            g  = QGroupBox(title)
            gl = QVBoxLayout(g)
            bl = QHBoxLayout()
            for lbl, fn in btn_items:
                b = QPushButton(lbl)
                b.clicked.connect(fn)
                bl.addWidget(b)
            bl.addStretch()
            gl.addLayout(bl)
            lw = QListWidget()
            lw.setMaximumHeight(100)
            lw.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
            setattr(self, list_attr, lw)
            gl.addWidget(lw)
            return g

        lay.addWidget(make_grp("🌐 Блокировка доменов (через hosts)", [
            ("🚫 Заблокировать…",  self._on_block_dom),
            ("✅ Разблокировать…", self._on_unblock_dom),
            ("❌ Снять выбранные", self._on_unblock_sel_dom),
        ], "blocked_domains_view"))

        lay.addWidget(make_grp("📵 Блокировка приложений (taskkill)", [
            ("🚫 Заблокировать…",  self._on_block_app),
            ("✅ Разблокировать…", self._on_unblock_app),
            ("❌ Снять выбранные", self._on_unblock_sel_app),
        ], "blocked_apps_view"))

        lay.addStretch()
        return w

    # ── Обновление агентов ────────────────────────────────────────

    def refresh_agents(self):
        online    = set(self.server.agents.keys())
        db_agents = self.db.list_agents()
        checked   = set(self._checked_hosts())

        self.agents_table.setRowCount(len(db_agents))
        n_online = 0
        for i, a in enumerate(db_agents):
            host      = a["hostname"]
            is_online = host in online
            if is_online:
                n_online += 1

            conn   = self.server.agents.get(host)
            ip_str = (conn.ip if conn else a["ip"]) or "?"
            os_str = (conn.os if conn else a["os"]) or "?"

            cb = QCheckBox()
            cb.setChecked(host in checked)
            cb.stateChanged.connect(self._update_summary_cards)
            self.agents_table.setCellWidget(i, 0, cb)

            for c, val in enumerate([host, ip_str, os_str], start=1):
                item = QTableWidgetItem(val)
                if not is_online:
                    item.setForeground(QColor("#aaa"))
                self.agents_table.setItem(i, c, item)

            st = QTableWidgetItem("● Онлайн" if is_online else "○ Оффлайн")
            st.setForeground(QColor("#1a7f37") if is_online else QColor("#aaa"))
            self.agents_table.setItem(i, 4, st)

            # Обновляем плитку в сетке экранов
            self.screen_grid.set_online(host, is_online)

        self.stat_label.setText(f"Онлайн: {n_online} / {len(db_agents)}")
        self._filter_agents_table()
        self._update_summary_cards()

    # ── Обновление задания ────────────────────────────────────────

    def on_job_update(self, job_id: str, status: str, message: str):
        host = self.server.job_to_host.get(job_id, "?")

        # Ищем существующую строку
        for r in range(self.jobs_table.rowCount()):
            it = self.jobs_table.item(r, 0)
            if it and it.data(Qt.ItemDataRole.UserRole) == job_id:
                self.jobs_table.item(r, 2).setText(status)
                self.jobs_table.item(r, 3).setText(message[:100])
                self._color_job_row(r, status)
                return

        # Новая строка
        r = self.jobs_table.rowCount()
        self.jobs_table.insertRow(r)
        host_item = QTableWidgetItem(host)
        host_item.setData(Qt.ItemDataRole.UserRole, job_id)
        self.jobs_table.setItem(r, 0, host_item)
        fn = Path(self._current_file).name if self._current_file else ""
        self.jobs_table.setItem(r, 1, QTableWidgetItem(fn))
        self.jobs_table.setItem(r, 2, QTableWidgetItem(status))
        self.jobs_table.setItem(r, 3, QTableWidgetItem(message[:100]))
        self._color_job_row(r, status)
        self.jobs_table.scrollToBottom()

        icon = "✅" if status == "done" else ("❌" if status == "error" else "⏳")
        self._log(f"{icon} [{host}] {status}: {message}")

    def _color_job_row(self, r: int, status: str):
        color = None
        if status == "done":    color = QColor("#e6f9ee")
        elif status == "error": color = QColor("#fde8e8")
        elif status in ("autoclick", "installing"): color = QColor("#fff9e6")
        if color:
            for c in range(self.jobs_table.columnCount()):
                it = self.jobs_table.item(r, c)
                if it:
                    it.setBackground(color)

    def on_screen_frame(self, hostname: str, b64: str, w: int, h: int):
        self.screen_grid.update_frame(hostname, b64, w, h)
        if self.fullscreen_view and self.fullscreen_view.hostname == hostname:
            self.fullscreen_view.set_frame(b64, w, h)

    # ── Хелперы выбора ────────────────────────────────────────────

    def _checked_hosts(self) -> list[str]:
        r = []
        for i in range(self.agents_table.rowCount()):
            cb   = self.agents_table.cellWidget(i, 0)
            item = self.agents_table.item(i, 1)
            if isinstance(cb, QCheckBox) and item and cb.isChecked():
                r.append(item.text())
        return r

    def _online_checked(self) -> list[str]:
        return [h for h in self._checked_hosts() if h in self.server.agents]

    def _all_online(self) -> list[str]:
        return list(self.server.agents.keys())

    def _set_checks(self, v: bool):
        for i in range(self.agents_table.rowCount()):
            cb = self.agents_table.cellWidget(i, 0)
            if isinstance(cb, QCheckBox):
                cb.setChecked(v)
        self._update_summary_cards()

    def _select_online(self):
        online = set(self.server.agents.keys())
        for i in range(self.agents_table.rowCount()):
            cb   = self.agents_table.cellWidget(i, 0)
            item = self.agents_table.item(i, 1)
            if isinstance(cb, QCheckBox) and item:
                cb.setChecked(item.text() in online)
        self._update_summary_cards()

    def _run(self, coro):
        asyncio.run_coroutine_threadsafe(coro, self.loop)

    def _log(self, text: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_view.append(f"<span style='color:#999'>[{ts}]</span> {text}")

    def _warn(self, msg: str = "Отметьте ПК галочками."):
        QMessageBox.warning(self, "Нет выбора", msg)

    def _need_hosts(self) -> list[str]:
        h = self._online_checked()
        if not h:
            self._warn()
        return h

    # ── Установка ─────────────────────────────────────────────────

    def on_file_dropped(self, path: str):
        self._current_file = path
        fn = Path(path).name
        self.drop.setText(f"📦 {fn}\n(перетащите снова или нажмите «Установить»)")
        self.drop.setStyleSheet("""
            QLabel { border:2px solid #2a9; border-radius:8px;
                     font-size:11pt; color:#2a9; background:#eafaf2; padding:8px; }
        """)

        self._update_summary_cards()
        hosts = self._online_checked()
        if not hosts:
            return

        all_ch  = self._checked_hosts()
        offline = len(set(all_ch) - set(hosts))
        flags   = self.flags_edit.text().strip()
        force   = self.chk_force_ac.isChecked()

        msg = f"Установить:\n{fn}\n\nНа {len(hosts)} онлайн-ПК?"
        if offline:
            msg += f"\n({offline} оффлайн — пропущены)"
        if flags:
            msg += f"\nФлаги: {flags}"
        if force:
            msg += "\nРежим: форс автокликер"

        if QMessageBox.question(self, "Подтвердите установку", msg) != QMessageBox.StandardButton.Yes:
            return

        self._run(self.server.install_file(hosts, path, flags, force))
        self._log(f"📦 Установка {fn} → {len(hosts)} ПК")

    def on_install(self):
        if not self._current_file:
            QMessageBox.warning(self, "Нет файла", "Сначала выберите файл (перетащите или нажмите в зону).")
            return
        hosts = self._online_checked()
        if not hosts:
            return self._warn()
        flags = self.flags_edit.text().strip()
        force = self.chk_force_ac.isChecked()
        self._run(self.server.install_file(hosts, self._current_file, flags, force))
        self._log(f"📦 Установка {Path(self._current_file).name} → {len(hosts)} ПК")

    def on_uninstall(self):
        hosts = self._online_checked()
        if not hosts:
            return self._warn()
        name, ok = QInputDialog.getText(
            self, "Удалить программу",
            "Имя программы (как в «Программы и компоненты»):",
        )
        if not ok or not name.strip():
            return
        self._run(self.server.uninstall(hosts, name.strip()))
        self._log(f"🗑 Удаление «{name}» → {len(hosts)} ПК")

    def on_schedule(self):
        if not self._current_file:
            QMessageBox.warning(self, "Нет файла", "Выберите файл.")
            return
        hosts = self._online_checked()
        if not hosts:
            return self._warn()

        dlg = QDialog(self)
        dlg.setWindowTitle("Запланировать установку")
        ll  = QVBoxLayout(dlg)
        ll.addWidget(QLabel(f"Файл: {Path(self._current_file).name}"))
        ll.addWidget(QLabel(f"ПК: {len(hosts)}"))
        dt = QDateTimeEdit(QDateTime.currentDateTime())
        dt.setCalendarPopup(True)
        ll.addWidget(QLabel("Время запуска:"))
        ll.addWidget(dt)
        bb = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        ll.addWidget(bb)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        self.db.add_schedule(
            run_at        = dt.dateTime().toSecsSinceEpoch(),
            hostnames_json= json.dumps(hosts),
            file_path     = self._current_file,
            action        = "install",
            flags         = self.flags_edit.text(),
        )
        self._log(f"📅 Запланировано на {dt.dateTime().toString()}")

    def _on_push_file(self):
        hosts = self._online_checked()
        if not hosts:
            return self._warn()
        path, _ = QFileDialog.getOpenFileName(self, "Файл для отправки")
        if not path:
            return
        target, ok = QInputDialog.getText(
            self, "Папка назначения",
            "Папка на ПК ученика (можно %переменные%):",
            text="%USERPROFILE%\\Desktop",
        )
        if not ok or not target:
            return
        self._run(self.server.push_file(hosts, path, target, True))
        self._log(f"📁 Отправка {Path(path).name} → {target} на {len(hosts)} ПК")

    def _quick_click(self, text: str):
        hosts = self._online_checked()
        if not hosts:
            return self._warn()
        self._run(self.server.installer_click(hosts, text))
        self._log(f"🖱 Клик «{text}» → {len(hosts)} ПК")

    # ── Экраны ────────────────────────────────────────────────────

    def _scr_start(self):
        hosts = self._online_checked()
        if not hosts:
            return self._warn()
        self._run(self.server.screen_start(hosts))
        self.screen_grid.set_hosts(sorted(set(self.screen_grid.tiles) | set(hosts)))
        self._log(f"▶ Трансляция → {len(hosts)} ПК")

    def _scr_stop(self):
        hosts = self._online_checked()
        if not hosts:
            return
        self._run(self.server.screen_stop(hosts))
        for h in hosts:
            self.screen_grid.clear_host(h)
        self._log(f"⏹ Стоп → {len(hosts)} ПК")

    def _scr_start_all(self):
        hosts = self._all_online()
        if not hosts:
            return
        self._run(self.server.screen_start(hosts))
        self.screen_grid.set_hosts(hosts)
        self._log(f"▶▶ Трансляция → все ({len(hosts)})")

    def _scr_stop_all(self):
        self._run(self.server.screen_stop(self._all_online()))
        for h in list(self.screen_grid.tiles):
            self.screen_grid.clear_host(h)
        self._log("⏹⏹ Трансляция остановлена")

    def _open_fullscreen(self, hostname: str):
        if self.fullscreen_view:
            self.fullscreen_view.close()
        self.fullscreen_view = FullScreenView(
            hostname,
            send_input=lambda h, pl: self._run(self.server.send_input(h, pl)),
            parent=self,
        )
        self.fullscreen_view.finished.connect(
            lambda: setattr(self, "fullscreen_view", None)
        )
        self.fullscreen_view.show()

    # ── Управление ────────────────────────────────────────────────

    def _on_lock(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        text, ok = QInputDialog.getText(
            self, "Текст блокировки", "Что увидит ученик:",
            text="Внимание! Смотрите на доску",
        )
        if not ok:
            return
        self._run(self.server.lock_screens(hosts, text or "Внимание учителю"))
        self._log(f"🔒 Блокировка → {len(hosts)} ПК")

    def _on_unlock(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        self._run(self.server.unlock_screens(hosts))
        self._log(f"🔓 Разблокировано → {len(hosts)} ПК")

    def _on_screamer(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        info = "(своя картинка)" if self._screamer_b64 else "(красный экран по умолчанию)"
        if QMessageBox.question(
            self, "Скример",
            f"Показать скример на {len(hosts)} ПК?\n{info}",
        ) != QMessageBox.StandardButton.Yes:
            return
        self._run(self.server.show_screamer(hosts, self._screamer_b64))
        self._log(f"😱 Скример → {len(hosts)} ПК")

    def _load_screamer_img(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Картинка для скримера", "",
            "Изображения (*.jpg *.jpeg *.png *.bmp)",
        )
        if not path:
            return
        try:
            with open(path, "rb") as f:
                self._screamer_b64 = base64.b64encode(f.read()).decode("ascii")
            QMessageBox.information(self, "Готово", f"Загружено: {Path(path).name}")
            self._log(f"🖼 Картинка скримера: {Path(path).name}")
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", str(e))

    def _on_msg(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        text, ok = QInputDialog.getMultiLineText(
            self, "Сообщение ученикам", "Текст:"
        )
        if not ok or not text:
            return
        self._run(self.server.send_message_box(hosts, "Сообщение от учителя", text))
        self._log(f"💬 Сообщение → {len(hosts)} ПК")

    def _on_power(self, action: str):
        hosts = self._need_hosts()
        if not hosts:
            return
        names = {"shutdown": "выключить", "reboot": "перезагрузить",
                 "logoff": "завершить сеанс на", "lock": "заблокировать Windows на"}
        if QMessageBox.question(
            self, "Подтверждение",
            f"{names.get(action, action)} {len(hosts)} ПК?",
        ) != QMessageBox.StandardButton.Yes:
            return
        self._run(self.server.power(hosts, action))
        self._log(f"⏻ {action} → {len(hosts)} ПК")

    def _on_mute(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        self._run(self.server.sound_control(hosts, "mute"))
        self._log(f"🔇 Звук выкл → {len(hosts)} ПК")

    def _on_unmute(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        self._run(self.server.sound_control(hosts, "unmute"))
        self._log(f"🔊 Звук вкл → {len(hosts)} ПК")

    def _on_volume(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        level, ok = QInputDialog.getDouble(
            self, "Громкость", "Уровень (0.0 — 1.0):", 0.5, 0.0, 1.0, 2
        )
        if not ok:
            return
        self._run(self.server.sound_control(hosts, "set_volume", level))
        self._log(f"🎚 Громкость {level:.2f} → {len(hosts)} ПК")

    def _on_play_audio(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите аудиофайл", "",
            "Аудио (*.wav *.mp3 *.wma *.m4a *.aac);;Все файлы (*.*)",
        )
        if not path:
            return
        level, ok = QInputDialog.getDouble(
            self, "Громкость перед запуском", "Уровень (0.0 — 1.0):", 0.8, 0.0, 1.0, 2
        )
        if not ok:
            return
        self._run(self.server.play_audio(hosts, path, level))
        self._log(f"🎵 {os.path.basename(path)} → {len(hosts)} ПК")

    def _on_speak_text(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        text, ok = QInputDialog.getMultiLineText(
            self, "Озвучка текста", "Текст для произнесения на выбранных ПК:"
        )
        if not ok or not text.strip():
            return
        level, ok = QInputDialog.getDouble(
            self, "Громкость озвучки", "Уровень (0.0 - 1.0):", 1.0, 0.0, 1.0, 2
        )
        if not ok:
            return
        self._run(self.server.speak_text(hosts, text.strip(), level))
        self._log(f"🗣 Озвучка текста → {len(hosts)} ПК")

    def _on_stop_audio(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        self._run(self.server.sound_control(hosts, "stop_audio"))
        self._log(f"⏹ Стоп аудио → {len(hosts)} ПК")

    def _on_run_prog(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        path, ok = QInputDialog.getText(
            self, "Запустить программу",
            "Путь или команда:", text="notepad.exe",
        )
        if not ok or not path:
            return
        args, ok = QInputDialog.getText(self, "Аргументы", "Аргументы (необязательно):")
        if not ok:
            return
        self._run(self.server.run_program(hosts, path, args or ""))
        self._log(f"▶ {path} → {len(hosts)} ПК")

    def _on_vscode_all(self):
        hosts = self._all_online()
        if not hosts:
            QMessageBox.information(self, "Нет агентов", "Нет онлайн-агентов.")
            return
        if QMessageBox.question(
            self, "VS Code", f"Открыть VS Code на {len(hosts)} ПК?",
        ) != QMessageBox.StandardButton.Yes:
            return
        self._run(self.server.open_vscode(hosts))
        self._log(f"🧑‍💻 VS Code → {len(hosts)} ПК")

    def _on_task_mgr(self):
        hosts = self._online_checked()
        if not hosts:
            return self._warn()
        if len(hosts) != 1:
            QMessageBox.warning(self, "Выбор", "Выберите ровно один ПК.")
            return
        ProcessManagerDialog(self, self.server, self.loop, hosts[0]).exec()

    def _on_installer_click_dlg(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        text, ok = QInputDialog.getText(
            self, "Кнопка установщика",
            "Текст кнопки (Далее, Next, I Agree, Install, Finish…):",
            text="Далее",
        )
        if not ok or not text.strip():
            return
        self._run(self.server.installer_click(hosts, text.strip()))
        self._log(f"🖱 Клик «{text.strip()}» → {len(hosts)} ПК")

    def _send_combo(self, vks: list[int]):
        hosts = self._need_hosts()
        if not hosts:
            return
        for h in hosts:
            self._run(self.server.send_input(h, {"kind": "combo", "vks": vks}))

    # ── Ограничения ───────────────────────────────────────────────

    def _on_block_dom(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        text, ok = QInputDialog.getText(
            self, "Блокировать домены",
            "Домены через запятую:", text="vk.com, youtube.com",
        )
        if not ok or not text:
            return
        domains = [d.strip() for d in text.split(",") if d.strip()]
        self._run(self.server.block_domains(hosts, domains))
        self.blocked_domains.update(domains)
        self._refresh_lists()
        self._log(f"🚫 Домены: {', '.join(domains)}")

    def _on_unblock_dom(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        text, ok = QInputDialog.getText(self, "Разблокировать", "Домены:")
        if not ok or not text:
            return
        domains = [d.strip() for d in text.split(",") if d.strip()]
        self._run(self.server.unblock_domains(hosts, domains))
        self.blocked_domains.difference_update(domains)
        self._refresh_lists()

    def _on_unblock_sel_dom(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        domains = [i.text() for i in self.blocked_domains_view.selectedItems()]
        if not domains:
            return
        self._run(self.server.unblock_domains(hosts, domains))
        self.blocked_domains.difference_update(domains)
        self._refresh_lists()

    def _on_block_app(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        text, ok = QInputDialog.getText(
            self, "Блокировать приложения",
            "Имена процессов через запятую (с .exe):",
            text="chrome.exe, firefox.exe",
        )
        if not ok or not text:
            return
        apps = [a.strip() for a in text.split(",") if a.strip()]
        self._run(self.server.block_apps(hosts, apps))
        self.blocked_apps.update(apps)
        self._refresh_lists()
        self._log(f"📵 Приложения: {', '.join(apps)}")

    def _on_unblock_app(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        text, ok = QInputDialog.getText(self, "Разблокировать", "Приложения:")
        if not ok or not text:
            return
        apps = [a.strip() for a in text.split(",") if a.strip()]
        self._run(self.server.unblock_apps(hosts, apps))
        self.blocked_apps.difference_update(apps)
        self._refresh_lists()

    def _on_unblock_sel_app(self):
        hosts = self._need_hosts()
        if not hosts:
            return
        apps = [i.text() for i in self.blocked_apps_view.selectedItems()]
        if not apps:
            return
        self._run(self.server.unblock_apps(hosts, apps))
        self.blocked_apps.difference_update(apps)
        self._refresh_lists()

    def _refresh_lists(self):
        self.blocked_domains_view.clear()
        self.blocked_domains_view.addItems(sorted(self.blocked_domains))
        self.blocked_apps_view.clear()
        self.blocked_apps_view.addItems(sorted(self.blocked_apps))
