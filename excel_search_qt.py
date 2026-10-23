#!/usr/bin/env python3
"""
Excel Search 2 (Qt) — SSC UZ
Тот же функционал, что excel_search.py (tkinter), на Qt (PySide6): таблица на модели — все строки
без лимита, ширины колонок (подгон, ручная, двойной клик) и правка ячеек с записью в тот же .xlsx.
Настройки — excel_search_settings.json рядом с программой (тот же формат, что у 1.x).
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import threading
import time
import zipfile

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, QTimer, Signal, QObject
from PySide6.QtGui import QColor, QFont, QFontMetrics, QGuiApplication, QIcon, QPixmap
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QFileDialog, QHBoxLayout, QHeaderView,
                               QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton, QRadioButton,
                               QScrollArea, QSplitter, QTableView, QVBoxLayout, QWidget,
                               QAbstractItemView, QButtonGroup, QColorDialog)

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(APP_DIR, "excel_search_settings.json")
DEFAULT_EXCEL = os.path.join(APP_DIR, "维修日报数据.xlsx")

VERSION      = "2.0.1"
APP_NAME     = "Excel Search 2"
UPDATE_REPO  = "MaxSmych/excel-quick-search"
UPDATE_ASSET = "Excel-Search-Qt.zip"   # своё имя: 1.x (tkinter) ищет Excel-Search.zip и на Qt не обновится
API_RELEASES = f"https://api.github.com/repos/{UPDATE_REPO}/releases?per_page=20"
EDITABLE_EXT = (".xlsx", ".xlsm")
MARK_RGB = (198, 239, 206)            # заливка отмеченной ✓ строки в файле (светло-зелёная)
# цвета отметок (как стили Excel «Хороший/Нейтральный/Плохой» + ещё два); «…» — любой свой
MARK_COLORS = (("зелёный", (198, 239, 206)), ("жёлтый", (255, 235, 156)), ("красный", (255, 199, 206)),
               ("голубой", (189, 215, 238)), ("оранжевый", (248, 203, 173)))
CHECK = "✓"
SHEET_COL = "_Лист_"
ROW_COL = "_строка_"                   # служебная: номер строки в своём листе (для записи в Excel)

HELP_TEXT = """Быстрый поиск по любому Excel (.xlsx / .xls / .xlsb).

1. «Обзор» — выбрать файл (или из списка недавних в поле «Файл»).
2. «Лист» / «Вся книга» — искать в одном листе или во всех сразу.
3. Слева — колонки: галочка показывает колонку, поле справа от неё — фильтр по этой колонке.
4. «Поиск» — ищет по всем видимым колонкам сразу, совпадения подсвечиваются.
5. Колонка ✓ — отметить строку «обработано»; отметки запоминаются для каждого файла.
6. Клик по заголовку — сортировка; клик по ячейке — значение копируется в буфер.
7. Ширина колонки: тянуть границу заголовка или двойной клик по ней (подгон) — запоминается.
8. Правка: двойной клик по ячейке или F2, Enter — значение записывается в ту же ячейку файла.
   Есть Excel — запись идёт через него (файл открыт у вас — правка видна прямо в окне Excel).
   ✓ заливает строку в самом файле — отметки видны всем, кому отправите файл. Цвет — квадратики
   «✓ цвет» рядом с поиском («…» — любой свой). ✓ строки другого цвета перекрашивает её в выбранный,
   ✓ того же цвета снимает отметку и заливку.
   Перед первой правкой файла копия кладётся в папку программы _backup.
9. «Авто↺» — сама перечитывает файл, когда он изменился. «Экспорт» — найденное в .xlsx/.csv.

Если что-то не так:
• без Excel на машине: .xls/.xlsb и файлы с диаграммами/сводными — только просмотр;
• обновление программа предлагает сама при запуске."""

C = {"base": "#1e1e2e", "mantle": "#181825", "surface0": "#313244", "surface1": "#45475a",
     "text": "#cdd6f4", "blue": "#89b4fa", "green": "#a6e3a1", "overlay": "#6c7086",
     "yellow": "#f9e2af", "red": "#f38ba8"}

STYLE = f"""
QWidget {{ background: {C['base']}; color: {C['text']}; font: 10pt 'Segoe UI'; }}
QLineEdit, QComboBox {{ background: {C['surface0']}; border: 1px solid {C['surface1']}; padding: 3px; }}
QPushButton {{ background: {C['surface0']}; border: none; padding: 5px 10px; }}
QPushButton:hover {{ background: {C['surface1']}; }}
QTableView {{ background: {C['mantle']}; gridline-color: {C['surface0']}; font: 9pt 'Segoe UI';
              selection-background-color: {C['surface1']}; selection-color: {C['text']}; }}
QHeaderView::section {{ background: {C['surface0']}; color: {C['blue']}; font: bold 9pt 'Segoe UI';
                        border: 1px solid {C['surface1']}; padding: 3px; }}
QScrollBar {{ background: {C['base']}; }}
QCheckBox, QRadioButton {{ font: 9pt 'Segoe UI'; }}
"""


def _resource(name):
    return os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))), name)


def _parse_ver(tag):
    out = []
    for part in str(tag).lstrip("vV").split(".")[:3]:
        num = "".join(ch for ch in part if ch.isdigit())
        out.append(int(num) if num else 0)
    return tuple(out + [0] * (3 - len(out)))


def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def as_text(df):
    """Текст ячеек: пустое (NaN, NaT) — "". В pandas 3 fillna("") не трогает NaT, а astype(str) делает
    из него NaN — поэтому where(notna) (то же, что в 1.1.0; ключи отметок совпадают с 1.x)."""
    return df.fillna("").astype(str).where(df.notna(), "")


def row_keys(df):
    """Ключ строки = md5 всех значений (без служебных колонок) — как в 1.x, отметки ✓ переносятся."""
    cols = [c for c in df.columns if c not in (SHEET_COL, ROW_COL)]
    sub = as_text(df[cols])
    return sub.apply(lambda r: hashlib.md5("\x1f".join(r.values).encode("utf-8")).hexdigest(), axis=1)


# ── модель таблицы ──────────────────────────────────────────────────────────────
class Model(QAbstractTableModel):
    """Колонка 0 — ✓ (галочка), дальше видимые колонки. Строки — индексы df после фильтра/сортировки."""

    def __init__(self, app):
        super().__init__()
        self.app = app
        self.cols, self.idx, self.text = [], [], None
        self.query, self.case = "", False

    def reset(self, cols, idx, text, query, case):
        self.beginResetModel()
        self.cols, self.idx, self.text = cols, idx, text
        self.query, self.case = query, case
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.idx)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() or not self.cols else len(self.cols) + 1

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation != Qt.Horizontal or role != Qt.DisplayRole:
            return None
        name = CHECK if section == 0 else self.cols[section - 1]
        arrow = (" ▲" if self.app.sort_asc else " ▼") if name == self.app.sort_col else ""
        return name + arrow

    def cell(self, row, col):
        return self.text.at[self.idx[row], self.cols[col - 1]]

    def data(self, index, role=Qt.DisplayRole):
        r, c = index.row(), index.column()
        if c == 0:
            if role == Qt.CheckStateRole:
                return Qt.Checked if self.app.keys.at[self.idx[r]] in self.app.checked else Qt.Unchecked
            return None
        if role in (Qt.DisplayRole, Qt.EditRole, Qt.ToolTipRole):
            v = self.cell(r, c)
            return v if role != Qt.ToolTipRole or len(v) > 40 else None
        if self.query and role in (Qt.BackgroundRole, Qt.ForegroundRole):
            v = self.cell(r, c)
            hit = (self.query in v) if self.case else (self.query.lower() in v.lower())
            if hit:
                return QColor(C["yellow"]) if role == Qt.BackgroundRole else QColor(C["base"])
        key = self.app.keys.at[self.idx[r]]
        if role == Qt.BackgroundRole and key in self.app.checked:
            return QColor(*self.app.row_color(key), 70)   # тот же цвет, что заливка в файле
        return None

    def flags(self, index):
        f = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if index.column() == 0:
            return f | Qt.ItemIsUserCheckable
        if self.app.can_edit and self.cols[index.column() - 1] not in (SHEET_COL,):
            f |= Qt.ItemIsEditable
        return f

    def setData(self, index, value, role=Qt.EditRole):
        r, c = index.row(), index.column()
        if c == 0 and role == Qt.CheckStateRole:
            key = self.app.keys.at[self.idx[r]]
            on = Qt.CheckState(value) == Qt.Checked
            if not on and key in self.app.checked and self.app.row_color(key) != self.app.mark_rgb:
                on = True                     # отмечена другим цветом — перекрасить в активный
            if self.app.can_edit and not self.app.mark_in_file(self.idx[r], on):
                return False
            self.app.set_checked(key, on)
            self.dataChanged.emit(index, self.index(r, self.columnCount() - 1))
            return True
        if c > 0 and role == Qt.EditRole:
            ok = self.app.write_cell(self.idx[r], self.cols[c - 1], value)
            if ok:
                self.dataChanged.emit(index, self.index(r, self.columnCount() - 1))
            return ok
        return False


def excel_installed():
    """Есть ли Excel (COM Excel.Application) и xlwings — тогда пишем через Excel: он сохраняет в файле всё."""
    try:
        import winreg
        import xlwings  # noqa: F401
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CurVer"))
        return True
    except Exception:
        return False


class Excel:
    """Запись через Excel (xlwings). Книга уже открыта у пользователя — правим прямо в его окне
    (видно сразу, файл не заблокирован). Иначе — свой скрытый Excel; он живёт, пока идут правки,
    и закрывается через минуту простоя (close_own), чтобы не держать файл."""

    def __init__(self):
        self.app = None
        self.own = {}

    def book(self, path):
        import xlwings as xw
        full = os.path.normcase(os.path.abspath(path))
        if full in self.own:
            return self.own[full]
        for app in xw.apps:
            if self.app is not None and app.pid == self.app.pid:
                continue
            for b in app.books:
                if os.path.normcase(b.fullname) == full:
                    return b
        if self.app is None:
            self.app = xw.App(visible=False, add_book=False)
            self.app.display_alerts = False
        b = self.app.books.open(path, update_links=False)
        self.own[full] = b
        return b

    def close_own(self):
        for b in self.own.values():
            try:
                b.close()
            except Exception:
                pass
        self.own = {}
        if self.app is not None:
            try:
                self.app.quit()
            except Exception:
                pass
            self.app = None


class Bus(QObject):
    """Сигналы из фоновых потоков в окно (Qt-виджеты трогать только из главного потока)."""
    update_found = Signal(str, str, int)
    message = Signal(str, str)


# ── заставка SM (кадры assets/splash, 60 к/с, затем растворение) ──────────────────
class Splash(QWidget):
    FPS, HOLD, FADE = 60, 0.1, 0.25

    def __init__(self):
        super().__init__(None, Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        folder = _resource(os.path.join("assets", "splash"))
        self.files = sorted(os.path.join(folder, f) for f in os.listdir(folder)
                            if f.endswith(".png")) if os.path.isdir(folder) else []
        self.label = QLabel(self)
        self.label.setStyleSheet("background: transparent;")
        if self.files:
            pm = QPixmap(self.files[0])
            self.label.setPixmap(pm)
            self.resize(pm.size())
            self.label.resize(pm.size())
            g = QGuiApplication.primaryScreen().availableGeometry()
            self.move(g.center() - self.rect().center())

    def play(self, on_done):
        if not self.files:
            on_done()
            return
        self.show()
        t0, st = time.perf_counter(), {"shown": 0, "revealed": False, "fade_at": None}
        end_anim = len(self.files) / self.FPS

        def tick():
            now = time.perf_counter()
            t = now - t0
            if t < end_anim:
                i = min(int(t * self.FPS), len(self.files) - 1)     # кадр по часам, а не по счётчику
                if i != st["shown"]:
                    self.label.setPixmap(QPixmap(self.files[i]))
                    st["shown"] = i
            elif not st["revealed"]:
                self.label.setPixmap(QPixmap(self.files[-1]))
                st["revealed"] = True
                on_done()
                self.raise_()
            elif st["fade_at"] is None:
                st["fade_at"] = now + self.HOLD
            else:
                alpha = 1 - (now - st["fade_at"]) / self.FADE
                if alpha <= 0:
                    timer.stop()
                    self.close()
                    return
                self.setWindowOpacity(min(1.0, alpha))
        timer = QTimer(self)
        timer.timeout.connect(tick)
        timer.start(5)


# ── окно ────────────────────────────────────────────────────────────────────────
class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Excel Search 2 — SSC UZ  ·  {VERSION}")
        self.resize(1280, 720)
        self.setMinimumSize(800, 500)
        if os.path.exists(_resource("icon.ico")):
            self.setWindowIcon(QIcon(_resource("icon.ico")))
        self.settings = load_settings()
        self.df = None
        self.text = None
        self.keys = None
        self.checked = set()
        self.mark_colors = {}                 # ключ строки → (r, g, b) её заливки
        self.mark_rgb = tuple(self.settings.get("mark_color") or MARK_RGB)
        self.sort_col, self.sort_asc = None, True
        self.current_file = None
        self.last_mtime = None
        self.can_edit = False
        self.backed_up = set()
        self.use_excel = excel_installed()
        self.excel = Excel()
        self.excel_idle = QTimer(self)
        self.excel_idle.setSingleShot(True)
        self.excel_idle.timeout.connect(self.excel.close_own)
        self._applying_widths = False
        self.col_checks, self.col_filters = {}, {}
        self.bus = Bus()
        self.bus.update_found.connect(self._offer_update)
        self.bus.message.connect(lambda t, m: QMessageBox.warning(self, t, m))
        self._build_ui()
        self.auto_timer = QTimer(self)
        self.auto_timer.timeout.connect(self._auto_tick)
        start = self.settings.get("last_file") or DEFAULT_EXCEL
        if start and os.path.exists(start):
            self.file_cb.setEditText(start)
            self.load_file(start)

    # ── интерфейс ──
    def _build_ui(self):
        root = QWidget()
        lay = QVBoxLayout(root)
        lay.setContentsMargins(6, 6, 6, 6)
        self.setCentralWidget(root)

        top = QHBoxLayout()
        top.addWidget(QLabel("Файл:"))
        self.file_cb = QComboBox()
        self.file_cb.setEditable(True)
        # поле узкое (видно имя файла — конец пути), полный путь — в подсказке и в выпадающем списке
        self.file_cb.setFixedWidth(300)
        self.file_cb.view().setMinimumWidth(700)
        self.file_cb.editTextChanged.connect(lambda t: (self.file_cb.setToolTip(t),
                                                        self.file_cb.lineEdit().end(False)))
        self.file_cb.addItems(self.settings.get("recent_files", []))
        self.file_cb.setCurrentIndex(-1)
        self.file_cb.activated.connect(lambda _i: self._on_load())
        top.addWidget(self.file_cb)
        for text, fn in (("Обзор", self._browse), ("Загрузить", self._on_load), ("↺ Обновить", self._on_refresh)):
            b = QPushButton(text)
            b.clicked.connect(fn)
            top.addWidget(b)
        self.rb_sheet, self.rb_all = QRadioButton("Лист"), QRadioButton("Вся книга")
        self.rb_sheet.setChecked(True)
        self.rb_sheet.toggled.connect(lambda _on: self.current_file and self.load_sheet())
        top.addWidget(self.rb_sheet)
        top.addWidget(self.rb_all)
        top.addWidget(QLabel("Лист:"))
        self.sheet_cb = QComboBox()
        self.sheet_cb.setMinimumWidth(150)
        self.sheet_cb.activated.connect(lambda _i: self.current_file and self.load_sheet())
        top.addWidget(self.sheet_cb)
        self.auto_chk = QCheckBox("Авто↺")
        self.auto_chk.toggled.connect(self._toggle_auto)
        top.addWidget(self.auto_chk)
        top.addStretch(1)
        self.status = QLabel("Файл не загружен")
        self.status.setStyleSheet(f"color: {C['green']};")
        top.addWidget(self.status)
        lay.addLayout(top)

        bar = QHBoxLayout()
        lbl = QLabel("Поиск:")
        lbl.setStyleSheet("font-size: 11pt;")
        bar.addWidget(lbl)
        self.query = QLineEdit()
        self.query.setMinimumWidth(380)
        self.query.setStyleSheet("font-size: 12pt;")
        self.query.textChanged.connect(self.search)
        bar.addWidget(self.query)
        self.case_chk = QCheckBox("Регистр")
        self.case_chk.toggled.connect(self.search)
        bar.addWidget(self.case_chk)
        clr = QPushButton("✕ Очистить")
        clr.clicked.connect(lambda: self.query.setText(""))
        bar.addWidget(clr)
        bar.addSpacing(16)
        bar.addWidget(QLabel("✓ цвет:"))
        self.color_group = QButtonGroup(self)
        self.color_group.setExclusive(True)
        for name, rgb in MARK_COLORS:
            self.color_group.addButton(self._swatch(name, rgb))
            bar.addWidget(self.color_group.buttons()[-1])
        self.custom_btn = self._swatch("свой цвет", self.mark_rgb)
        self.custom_btn.setVisible(self.mark_rgb not in [c for _, c in MARK_COLORS])
        self.color_group.addButton(self.custom_btn)
        bar.addWidget(self.custom_btn)
        more = QPushButton("…")
        more.setFixedWidth(28)
        more.setToolTip("Выбрать любой цвет")
        more.clicked.connect(self._pick_color)
        bar.addWidget(more)
        self._sync_swatches()
        bar.addStretch(1)
        exp = QPushButton("Экспорт")
        exp.clicked.connect(self.export)
        bar.addWidget(exp)
        hlp = QPushButton("?")
        hlp.setFixedWidth(32)
        hlp.clicked.connect(lambda: QMessageBox.information(self, f"Excel Search {VERSION}", HELP_TEXT))
        bar.addWidget(hlp)
        lay.addLayout(bar)

        self.split = QSplitter(Qt.Horizontal)
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(4, 4, 4, 4)
        hdr = QHBoxLayout()
        t = QLabel("Колонки")
        t.setStyleSheet("font-weight: bold;")
        hdr.addWidget(t)
        hdr.addStretch(1)
        for text, val in (("Ноль", False), ("Все", True)):
            b = QPushButton(text)
            b.clicked.connect(lambda _c=False, v=val: self._toggle_all(v))
            hdr.addWidget(b)
        ll.addLayout(hdr)
        self.col_area = QScrollArea()
        self.col_area.setWidgetResizable(True)
        self.col_box = QWidget()
        self.col_lay = QVBoxLayout(self.col_box)
        self.col_lay.setContentsMargins(0, 0, 0, 0)
        self.col_lay.addStretch(1)
        self.col_area.setWidget(self.col_box)
        ll.addWidget(self.col_area)
        self.split.addWidget(left)

        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        self.model = Model(self)
        self.view = QTableView()
        self.view.setModel(self.model)
        self.view.verticalHeader().setVisible(False)
        self.view.verticalHeader().setDefaultSectionSize(22)
        self.view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        h = self.view.horizontalHeader()
        h.setSectionResizeMode(QHeaderView.Interactive)      # двойной клик по границе = подгон (Qt сам)
        h.setResizeContentsPrecision(500)                     # подгон смотрит до 500 строк — быстро
        h.setSectionsClickable(True)
        h.sectionClicked.connect(self._on_header_click)
        h.sectionResized.connect(self._on_col_resized)
        self.view.clicked.connect(self._on_cell_click)
        rl.addWidget(self.view)
        fb = QHBoxLayout()
        fx = QLabel("fx")
        fx.setStyleSheet(f"color: {C['overlay']}; font-style: italic;")
        fb.addWidget(fx)
        self.fbar = QLineEdit()
        self.fbar.setReadOnly(True)
        fb.addWidget(self.fbar)
        rl.addLayout(fb)
        self.split.addWidget(right)
        self.split.setSizes([self.settings.get("col_panel_width", 230), 1000])
        self.split.splitterMoved.connect(self._on_split)
        lay.addWidget(self.split, 1)

    # ── файлы ──
    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel", os.path.dirname(self.current_file or ""),
                                              "Excel (*.xlsx *.xlsm *.xls *.xlsb);;Все (*.*)")
        if path:
            self.file_cb.setEditText(path)
            self.load_file(path)

    def _on_load(self):
        path = self.file_cb.currentText().strip()
        if path:
            self.load_file(path)

    def _on_refresh(self):
        if self.current_file:
            self.load_file(self.current_file)
        else:
            self.status.setText("Нет загруженного файла")

    def _add_recent(self, path):
        lst = [p for p in self.settings.get("recent_files", []) if p != path]
        lst.insert(0, path)
        self.settings["recent_files"] = lst[:15]
        save_settings(self.settings)
        self.file_cb.blockSignals(True)
        self.file_cb.clear()
        self.file_cb.addItems(self.settings["recent_files"])
        self.file_cb.setEditText(path)
        self.file_cb.blockSignals(False)

    def _xl(self, path):
        return pd.ExcelFile(path, engine="pyxlsb" if path.lower().endswith(".xlsb") else None)

    def load_file(self, path):
        self.status.setText("Загрузка…")
        QApplication.processEvents()
        xl = None
        try:
            xl = self._xl(path)
            cur = self.sheet_cb.currentText()
            self.sheet_cb.blockSignals(True)
            self.sheet_cb.clear()
            self.sheet_cb.addItems(xl.sheet_names)
            self.sheet_cb.setCurrentText(cur if cur in xl.sheet_names else xl.sheet_names[0])
            self.sheet_cb.blockSignals(False)
            if path != self.current_file:
                self.backed_up.discard(path)
            self.current_file = path
            self.last_mtime = os.path.getmtime(path)
            self.settings["last_file"] = path
            self._add_recent(path)
            self.load_sheet(xl)
        except Exception as exc:
            self.status.setText(f"Ошибка: {exc}")
        finally:
            if xl is not None:
                xl.close()       # иначе файл занят нами и Excel не может его сохранить (22.10.2026)

    def load_sheet(self, xl=None):
        own = xl is None
        try:
            xl = xl or self._xl(self.current_file)
            if self.rb_all.isChecked():
                frames = []
                for name in xl.sheet_names:
                    d = xl.parse(name)
                    d.insert(0, SHEET_COL, name)
                    d[ROW_COL] = range(2, len(d) + 2)
                    frames.append(d)
                self.df = pd.concat(frames, ignore_index=True)
                self.sheet_cb.setEnabled(False)
            else:
                self.df = xl.parse(self.sheet_cb.currentText())
                self.df[ROW_COL] = range(2, len(self.df) + 2)
                self.sheet_cb.setEnabled(True)
            self.text = as_text(self.df)
            self.keys = row_keys(self.df)
            self.checked = set(self.settings.get("checked", {}).get(self.current_file, []))
            self.mark_colors = {k: tuple(v) for k, v in
                                self.settings.get("mark_colors", {}).get(self.current_file, {}).items()}
            self.can_edit, why = self._editable()
            self._rebuild_columns([c for c in self.df.columns if c != ROW_COL])
            self.search()
            self.status.setText(f"{len(self.df)} строк · {len(self.df.columns) - 1} колонок  |  "
                                f"{os.path.basename(self.current_file)}" + ("" if self.can_edit else f"  ·  {why}"))
        except Exception as exc:
            self.status.setText(f"Ошибка: {exc}")
        finally:
            if own and xl is not None:
                xl.close()

    def _editable(self):
        """Через Excel — любой файл (Excel сохраняет всё). Без Excel (openpyxl) — только .xlsx/.xlsm без
        диаграмм/сводных/картинок: openpyxl их при сохранении теряет."""
        f = self.current_file or ""
        if self.use_excel:
            return True, ""
        if not f.lower().endswith(EDITABLE_EXT):
            return False, "только просмотр (.xls/.xlsb)"
        try:
            names = zipfile.ZipFile(f).namelist()
        except Exception:
            return False, "только просмотр"
        if any(k in n for n in names for k in ("xl/charts/", "xl/pivot", "xl/drawings/", "xl/media/")):
            return False, "только просмотр (в файле диаграммы/сводные/рисунки)"
        return True, ""

    # ── колонки слева ──
    def _rebuild_columns(self, columns):
        while self.col_lay.count() > 1:
            w = self.col_lay.takeAt(0).widget()
            if w:
                w.deleteLater()
        self.col_checks, self.col_filters = {}, {}
        vis = self.settings.get("col_visibility", {})
        flt = self.settings.get("col_filters", {})
        for col in columns:
            row = QWidget()
            hl = QHBoxLayout(row)
            hl.setContentsMargins(0, 0, 0, 0)
            cb = QCheckBox(str(col))
            cb.setChecked(vis.get(col, True))
            cb.toggled.connect(self._on_col_toggle)
            fe = QLineEdit(flt.get(col, ""))
            fe.setStyleSheet("font-size: 8pt; padding: 1px;")
            # гибко, но не шире 15 знаков SN (как в 1.1.2)
            fe.setMinimumWidth(40)
            fe.setMaximumWidth(QFontMetrics(QFont("Segoe UI", 8)).horizontalAdvance("W" * 15) + 8)
            fe.textChanged.connect(self._on_filter_change)
            hl.addWidget(cb)
            hl.addWidget(fe, 1)
            self.col_lay.insertWidget(self.col_lay.count() - 1, row)
            self.col_checks[col], self.col_filters[col] = cb, fe

    def _on_col_toggle(self):
        self.settings["col_visibility"] = {c: cb.isChecked() for c, cb in self.col_checks.items()}
        save_settings(self.settings)
        self.search()

    def _toggle_all(self, val):
        for cb in self.col_checks.values():
            cb.blockSignals(True)
            cb.setChecked(val)
            cb.blockSignals(False)
        self._on_col_toggle()

    def _on_filter_change(self):
        self.settings["col_filters"] = {c: fe.text() for c, fe in self.col_filters.items() if fe.text().strip()}
        save_settings(self.settings)
        self.search()

    def _on_split(self, *_):
        self.settings["col_panel_width"] = self.split.sizes()[0]
        save_settings(self.settings)

    # ── поиск ──
    def _mask(self, cols):
        q, case, df = self.query.text().strip(), self.case_chk.isChecked(), self.text
        mask = pd.Series(True, index=df.index)
        if q:
            mask &= df[cols].apply(lambda s: s.str.contains(q, case=case, regex=False)).any(axis=1)
        for col, fe in self.col_filters.items():
            fq = fe.text().strip()
            if fq and col in df.columns:
                mask &= df[col].str.contains(fq, case=case, regex=False)
        return mask

    def search(self, *_):
        if self.df is None:
            return
        cols = [c for c, cb in self.col_checks.items() if cb.isChecked()]
        if not cols:
            self.model.reset([], [], self.text, "", False)
            return
        mask = self._mask(cols)
        idx = self.df.index[mask]
        if self.sort_col == CHECK:
            chk = self.keys.loc[idx].isin(self.checked)
            idx = chk.sort_values(ascending=self.sort_asc, kind="stable").index
        elif self.sort_col in cols:
            s = self.df.loc[idx, self.sort_col]
            s = s.astype(str).str.lower() if s.dtype == object or str(s.dtype) == "str" else s
            idx = s.sort_values(ascending=self.sort_asc, kind="stable", na_position="last").index
        self.model.reset(cols, list(idx), self.text, self.query.text().strip(), self.case_chk.isChecked())
        self._apply_widths()
        self.status.setText(f"{int(mask.sum())} совпадений  |  {os.path.basename(self.current_file or '')}")

    # ── ширина колонок: подгон под содержимое + ручная/двойной клик, запоминается по файлу ──
    def _saved_widths(self):
        return self.settings.get("col_widths", {}).get(self.current_file or "", {})

    def _apply_widths(self):
        self._applying_widths = True
        try:
            self.view.resizeColumnsToContents()
            h = self.view.horizontalHeader()
            h.resizeSection(0, 34)
            saved = self._saved_widths()
            for i, name in enumerate(self.model.cols, start=1):
                w = saved.get(name)
                if w:
                    h.resizeSection(i, int(w))
                elif h.sectionSize(i) > 500:      # очень длинный текст не растягивает колонку на весь экран
                    h.resizeSection(i, 500)
        finally:
            self._applying_widths = False

    def _on_col_resized(self, logical, _old, new):
        if self._applying_widths or logical == 0 or logical > len(self.model.cols):
            return
        store = self.settings.setdefault("col_widths", {}).setdefault(self.current_file or "", {})
        store[self.model.cols[logical - 1]] = int(new)
        save_settings(self.settings)

    # ── клики ──
    def _on_header_click(self, section):
        col = CHECK if section == 0 else self.model.cols[section - 1]
        if self.sort_col == col:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col, self.sort_asc = col, True
        self.search()

    def _on_cell_click(self, index):
        if index.column() == 0:
            return
        text = self.model.cell(index.row(), index.column())
        self.fbar.setText(text)
        if text:
            QGuiApplication.clipboard().setText(text)
            self.status.setText(f"Скопировано: {text[:80]}")

    # ── отметки ✓ ──
    def set_checked(self, key, on):
        (self.checked.add if on else self.checked.discard)(key)
        if on:
            self.mark_colors[key] = self.mark_rgb
        else:
            self.mark_colors.pop(key, None)
        f = self.current_file or ""
        self.settings.setdefault("checked", {})[f] = sorted(self.checked)
        self.settings.setdefault("mark_colors", {})[f] = {k: list(v) for k, v in self.mark_colors.items()}
        save_settings(self.settings)

    def row_color(self, key):
        return self.mark_colors.get(key, MARK_RGB)

    # ── цвет отметки ──
    def _swatch(self, name, rgb):
        b = QPushButton()
        b.setCheckable(True)
        b.setFixedSize(24, 24)
        b.setToolTip(f"Отмечать ✓ цветом: {name}")
        b.setProperty("rgb", rgb)
        self._paint_swatch(b, rgb)
        b.clicked.connect(lambda _=False, btn=b: self._set_mark_color(tuple(btn.property("rgb"))))
        return b

    @staticmethod
    def _paint_swatch(b, rgb):
        hexc = "#%02x%02x%02x" % tuple(rgb)
        b.setStyleSheet(f"QPushButton {{ background: {hexc}; border: 2px solid {C['surface1']}; padding: 0; }}"
                        f"QPushButton:checked {{ border: 3px solid {C['text']}; }}")

    def _set_mark_color(self, rgb):
        self.mark_rgb = tuple(rgb)
        self.settings["mark_color"] = list(self.mark_rgb)
        save_settings(self.settings)
        self._sync_swatches()

    def _sync_swatches(self):
        for b in self.color_group.buttons():
            b.setChecked(tuple(b.property("rgb")) == self.mark_rgb)

    def _pick_color(self):
        c = QColorDialog.getColor(QColor(*self.mark_rgb), self, "Цвет отметки ✓")
        if not c.isValid():
            return
        rgb = (c.red(), c.green(), c.blue())
        self.custom_btn.setProperty("rgb", rgb)
        self._paint_swatch(self.custom_btn, rgb)
        self.custom_btn.setVisible(True)
        self._set_mark_color(rgb)

    # ── правка ячейки → запись в тот же файл ──
    def write_cell(self, i, col, value):
        """Записать значение в ячейку файла. Строка Excel = _строка_ (заголовок в строке 1), колонка —
        позиция колонки в листе. Перед записью сверяется старое значение: файл могли поменять."""
        import openpyxl
        new = "" if value is None else str(value)
        old_txt = self.text.at[i, col]
        if new == old_txt:
            return False
        sheet = self.df.at[i, SHEET_COL] if SHEET_COL in self.df.columns else self.sheet_cb.currentText()
        cols = [c for c in self.df.columns if c not in (SHEET_COL, ROW_COL)]
        if SHEET_COL in self.df.columns:      # «вся книга»: колонки листа = те, что у его строк не пустые
            sub = self.df[self.df[SHEET_COL] == sheet]
            cols = [c for c in cols if c in sub.columns]
        xrow, xcol = int(self.df.at[i, ROW_COL]), cols.index(col) + 1
        path = self.current_file
        try:
            self._before_write(path)
            if self.use_excel:
                self._excel_write(path, sheet, xrow, xcol, col, old_txt, new)
                raise _Done
            wb = openpyxl.load_workbook(path, keep_vba=path.lower().endswith(".xlsm"))
            ws = wb[sheet]
            head = str(ws.cell(1, xcol).value or "").strip()
            if head != str(col).split(".")[0].strip() and head != str(col).strip():
                raise RuntimeError(f"в файле колонка {xcol} называется «{head}», а не «{col}»")
            cell = ws.cell(xrow, xcol)
            cur = cell.value
            cur_txt = "" if cur is None else str(cur)
            if cur_txt != old_txt and not (isinstance(cur, (int, float)) and old_txt and _same_num(cur, old_txt)):
                if not (hasattr(cur, "isoformat") and old_txt.startswith(str(cur)[:10])):
                    raise RuntimeError(f"в файле в этой ячейке «{cur_txt}», а в таблице «{old_txt}» — "
                                       "нажмите «Обновить»")
            cell.value = _typed(new, cur)
            wb.save(path)
        except _Done:
            pass
        except PermissionError:
            QMessageBox.warning(self, "Правка", "Файл открыт в Excel — закройте его и повторите правку.")
            return False
        except Exception as exc:
            QMessageBox.warning(self, "Правка", f"Не записал: {exc}")
            return False
        was = self.keys.at[i] in self.checked
        self.df.at[i, col] = _typed(new, None) if new else None
        self.text.at[i, col] = new
        nk = row_keys(self.df.loc[[i]]).iloc[0]
        if was:
            old_rgb = self.row_color(self.keys.at[i])
            self.checked.discard(self.keys.at[i])
            self.mark_colors.pop(self.keys.at[i], None)
            self.keys.at[i] = nk
            keep, self.mark_rgb = self.mark_rgb, old_rgb
            self.set_checked(nk, True)
            self.mark_rgb = keep
        else:
            self.keys.at[i] = nk
        self.last_mtime = os.path.getmtime(path)
        self.status.setText(f"Записано в файл: {col} = {new[:60]}  (строка {xrow})")
        return True

    def _before_write(self, path):
        if os.path.getmtime(path) != self.last_mtime:
            raise RuntimeError("файл изменился после загрузки — нажмите «Обновить» и повторите")
        if path not in self.backed_up:
            bdir = os.path.join(APP_DIR, "_backup")
            os.makedirs(bdir, exist_ok=True)
            base, ext = os.path.splitext(os.path.basename(path))
            shutil.copy2(path, os.path.join(bdir, f"{base}_{time.strftime('%Y%m%d_%H%M%S')}{ext}"))
            self.backed_up.add(path)

    def _excel_sheet(self, path, sheet):
        self.status.setText("Excel: открываю файл…")
        QApplication.processEvents()
        b = self.excel.book(path)
        return b, b.sheets[sheet]

    def _excel_saved(self, b):
        b.save()
        self.last_mtime = os.path.getmtime(self.current_file)
        self.excel_idle.start(60_000)

    def _excel_write(self, path, sheet, xrow, xcol, col, old_txt, new):
        b, ws = self._excel_sheet(path, sheet)
        head = str(ws.range(1, xcol).value or "").strip()
        if head != str(col).split(".")[0].strip() and head != str(col).strip():
            raise RuntimeError(f"в файле колонка {xcol} называется «{head}», а не «{col}»")
        rng = ws.range(xrow, xcol)
        cur = rng.value
        cur_txt = "" if cur is None else str(cur)
        if cur_txt != old_txt and not (isinstance(cur, (int, float)) and old_txt and _same_num(cur, old_txt)):
            if not (hasattr(cur, "isoformat") and old_txt.startswith(str(cur)[:10])):
                raise RuntimeError(f"в файле в этой ячейке «{cur_txt}», а в таблице «{old_txt}» — "
                                   "нажмите «Обновить»")
        rng.value = _typed(new, cur)
        self._excel_saved(b)

    def mark_in_file(self, i, on):
        """✓ → залить строку в файле (MARK_RGB), снять ✓ → убрать заливку. Через Excel — любой файл,
        без Excel — openpyxl (только .xlsx/.xlsm)."""
        sheet = self.df.at[i, SHEET_COL] if SHEET_COL in self.df.columns else self.sheet_cb.currentText()
        xrow = int(self.df.at[i, ROW_COL])
        path = self.current_file
        try:
            self._before_write(path)
            if self.use_excel:
                b, ws = self._excel_sheet(path, sheet)
                last = ws.used_range.last_cell.column
                ws.range((xrow, 1), (xrow, last)).color = self.mark_rgb if on else None
                self._excel_saved(b)
            else:
                import openpyxl
                from openpyxl.styles import PatternFill
                wb = openpyxl.load_workbook(path, keep_vba=path.lower().endswith(".xlsm"))
                ws = wb[sheet]
                fill = PatternFill("solid", fgColor="%02X%02X%02X" % self.mark_rgb) if on else PatternFill()
                for c in range(1, ws.max_column + 1):
                    ws.cell(xrow, c).fill = fill
                wb.save(path)
                self.last_mtime = os.path.getmtime(path)
        except PermissionError:
            QMessageBox.warning(self, "Отметка", "Файл открыт в Excel — закройте его и повторите.")
            return False
        except Exception as exc:
            QMessageBox.warning(self, "Отметка", f"Не записал в файл: {exc}")
            return False
        self.status.setText(f"{'Отмечено' if on else 'Снята отметка'} в файле: строка {xrow}")
        return True

    # ── авто-перезагрузка ──
    def _toggle_auto(self, on):
        if on and self.current_file:
            self.auto_timer.start(5000)
        else:
            self.auto_timer.stop()
            self.auto_chk.setChecked(False) if on else None

    def _auto_tick(self):
        if self.current_file and os.path.exists(self.current_file) \
                and os.path.getmtime(self.current_file) != self.last_mtime:
            self.load_file(self.current_file)

    # ── экспорт ──
    def export(self):
        if self.df is None:
            return
        cols = [c for c, cb in self.col_checks.items() if cb.isChecked()]
        if not cols:
            return
        result = self.df.loc[self.df.index[self._mask(cols)], cols].copy()
        result.insert(0, CHECK, self.keys.loc[result.index].isin(self.checked).map({True: CHECK, False: ""}).values)
        path, _ = QFileDialog.getSaveFileName(self, "Экспорт", "", "Excel (*.xlsx);;CSV (*.csv)")
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                result.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                result.to_excel(path if path.lower().endswith(".xlsx") else path + ".xlsx", index=False)
            QMessageBox.information(self, "Экспорт", f"Сохранено {len(result)} строк\n{path}")
        except Exception as exc:
            QMessageBox.warning(self, "Экспорт", f"Ошибка: {exc}")

    # ── ярлык и обновление (стандарт GUI, brain 04 §10в) ──
    def make_shortcut(self):
        if not getattr(sys, "frozen", False) or self.settings.get("shortcut2_done"):
            return
        exe = sys.executable
        ps = ("$l=Join-Path ([Environment]::GetFolderPath('Desktop')) '{0}.lnk';"
              "if(-not (Test-Path -LiteralPath $l)){{"
              "$s=(New-Object -ComObject WScript.Shell).CreateShortcut($l);"
              "$s.TargetPath='{1}';$s.WorkingDirectory='{2}';$s.IconLocation='{1},0';$s.Save()}}"
              ).format(*(p.replace("'", "''") for p in (APP_NAME, exe, APP_DIR)))
        try:
            subprocess.Popen(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
                             creationflags=0x08000000)
            self.settings["shortcut2_done"] = True
            save_settings(self.settings)
        except OSError:
            pass

    def _fail_path(self):
        return os.path.join(APP_DIR, ".excel_search.update-failed.txt")

    def check_update(self):
        why = self._fail_path()
        if os.path.exists(why):
            try:
                msg = open(why, encoding="utf-8").read().strip()
                os.remove(why)
                QMessageBox.warning(self, "Обновление", f"Прошлое обновление не встало: {msg}")
            except OSError:
                pass

        def worker():
            import urllib.request
            try:
                req = urllib.request.Request(API_RELEASES, headers={
                    "Accept": "application/vnd.github+json", "User-Agent": "excel-search-qt"})
                rels = json.load(urllib.request.urlopen(req, timeout=6))
            except Exception:
                return
            # своя линия сборок: самый новый релиз (в т.ч. пре-релиз) с ассетом Excel-Search-Qt.zip
            best = None
            for r in rels:
                a = next((a for a in r.get("assets", []) if a.get("name") == UPDATE_ASSET), None)
                if a and not r.get("draft") and (best is None or _parse_ver(r["tag_name"]) > _parse_ver(best[0])):
                    best = (r["tag_name"], a["browser_download_url"], int(a.get("size") or 0))
            if best and _parse_ver(best[0]) > _parse_ver(VERSION) and getattr(sys, "frozen", False):
                self.bus.update_found.emit(*best)
        threading.Thread(target=worker, daemon=True).start()

    def _offer_update(self, tag, url, size):
        if QMessageBox.question(self, "Обновление", f"Доступна версия {tag} (у вас {VERSION}).\n\n"
                                "Скачать и обновить сейчас? Программа закроется и откроется заново."
                                ) == QMessageBox.Yes:
            threading.Thread(target=self._update_worker, args=(url, size), daemon=True).start()

    def _update_worker(self, url, size):
        """Сборка onedir: архив во временную папку → после выхода скрытый PowerShell распаковывает его
        поверх папки программы и запускает exe. Настроек в архиве нет — не трогаются."""
        import tempfile, urllib.request
        zp = os.path.join(tempfile.gettempdir(), UPDATE_ASSET)
        try:
            with urllib.request.urlopen(url, timeout=180) as r, open(zp, "wb") as f:
                shutil.copyfileobj(r, f)
            if size and os.path.getsize(zp) != size:
                raise OSError(f"скачано {os.path.getsize(zp)} байт из {size}")
        except Exception as exc:
            self.bus.message.emit("Обновление", f"Не скачалось: {exc}")
            return

        def q(p):
            return "'" + p.replace("'", "''") + "'"
        tar = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "tar.exe")
        script = (
            "$ErrorActionPreference='SilentlyContinue'\r\n"
            f"Wait-Process -Id {os.getpid()} -Timeout 60\r\n"
            "$ok=$false\r\n"
            "for($i=0;$i -lt 120;$i++){\r\n"
            f"  & {q(tar)} -x -f {q(zp)} -C {q(APP_DIR)}\r\n"
            "  if($LASTEXITCODE -eq 0){$ok=$true;break}\r\n"
            "  Start-Sleep -Milliseconds 500}\r\n"
            f"if(-not $ok){{Set-Content -LiteralPath {q(self._fail_path())} "
            "-Value 'файлы программы были заняты' -Encoding UTF8}\r\n"
            f"Remove-Item -LiteralPath {q(zp)}\r\n"
            f"Start-Process -FilePath {q(sys.executable)}\r\n"
            "Remove-Item -LiteralPath $MyInvocation.MyCommand.Path\r\n")
        ps1 = os.path.join(tempfile.gettempdir(), "excel_search_qt_update.ps1")
        with open(ps1, "w", encoding="utf-8-sig") as f:
            f.write(script)
        env = {k: v for k, v in os.environ.items() if not k.startswith(("_MEI", "_PYI"))}
        cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", ps1]
        try:
            subprocess.Popen(cmd, creationflags=0x08000000 | 0x01000000, env=env)
        except OSError:
            subprocess.Popen(cmd, creationflags=0x08000000, env=env)
        QTimer.singleShot(0, QApplication.instance().quit)

    def closeEvent(self, ev):
        self.excel.close_own()
        save_settings(self.settings)
        super().closeEvent(ev)


class _Done(Exception):
    """Запись прошла через Excel — ветку openpyxl пропустить."""


def _same_num(cur, txt):
    try:
        return float(txt) == float(cur)
    except ValueError:
        return False


def _typed(text, like):
    """Введённый текст → число, если похоже на число (как значение в Excel); иначе строка; пусто — None."""
    if text == "":
        return None
    t = text.strip().replace(",", ".") if isinstance(like, (int, float)) or like is None else None
    if t is not None:
        try:
            f = float(t)
            return int(f) if f.is_integer() and "." not in t else f
        except ValueError:
            pass
    return text


def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLE)
    win = App()
    if getattr(sys, "frozen", False) or os.environ.get("SPLASH") == "1":
        sp = Splash()
        sp.play(win.show)
        win._splash = sp
    else:
        win.show()
    win.make_shortcut()
    QTimer.singleShot(1500, win.check_update)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
