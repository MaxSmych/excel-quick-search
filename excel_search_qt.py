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
from PySide6.QtGui import QColor, QFont, QFontMetrics, QGuiApplication, QIcon, QPen, QPixmap
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QFileDialog, QHBoxLayout, QHeaderView,
                               QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton, QRadioButton,
                               QScrollArea, QSplitter, QTableView, QVBoxLayout, QWidget,
                               QAbstractItemView, QStyle, QStyledItemDelegate, QStyleOptionViewItem)

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(APP_DIR, "excel_search_settings.json")
DEFAULT_EXCEL = os.path.join(APP_DIR, "维修日报数据.xlsx")

VERSION      = "2.0.3"
APP_NAME     = "Excel Search 2"
UPDATE_REPO  = "MaxSmych/excel-quick-search"
UPDATE_ASSET = "Excel-Search-Qt.zip"   # своё имя: 1.x (tkinter) ищет Excel-Search.zip и на Qt не обновится
API_RELEASES = f"https://api.github.com/repos/{UPDATE_REPO}/releases?per_page=20"
EDITABLE_EXT = (".xlsx", ".xlsm")
# цвета заливки ячеек — насыщенные, как в рабочих таблицах пользователя (23.10.2026)
MARK_COLORS = (("зелёный", (0, 176, 80)), ("жёлтый", (255, 255, 0)), ("красный", (231, 76, 60)),
               ("оранжевый", (247, 150, 70)))
CHECK = "✓"
SHEET_COL = "_Лист_"
ROW_COL = "_строка_"                   # служебная: номер строки в своём листе (для записи в Excel)

HELP_TEXT = """Быстрый поиск по любому Excel (.xlsx / .xls / .xlsb).

1. «Обзор» — выбрать файл (или из списка недавних в поле «Файл»).
2. «Лист» / «Вся книга» — искать в одном листе или во всех сразу.
3. Слева — колонки: галочка показывает колонку, поле справа от неё — фильтр по этой колонке.
   Галочка «Колонки» — показать все / скрыть все.
4. «Поиск» — ищет по всем видимым колонкам сразу, совпадения подсвечиваются.
5. Колонка ✓ — отметка для себя; запоминается для каждого файла, в файл не пишется.
6. Клик по заголовку — сортировка; клик по ячейке — значение копируется в буфер.
7. Ширина колонки: тянуть границу заголовка или двойной клик по ней (подгон) — запоминается.
8. Правка: двойной клик по ячейке или F2, Enter — значение записывается в ту же ячейку файла.
   Есть Excel — запись идёт через него (файл открыт у вас — правка видна прямо в окне Excel).
   Цвет: выделите ячейки (мышью, Shift/Ctrl) и нажмите квадратик справа от поиска — ячейки
   зальются и в программе, и в самом файле (видно всем, кому отправите файл). Ещё раз тот же
   цвет — заливка снимается. Заливка, сделанная в самом Excel, тоже видна (.xlsx/.xlsm).
   Выделенные ячейки обведены оранжевой рамкой, строка и колонка текущей ячейки подсвечены.
   ✓ — просто отметка для себя, в файл не пишется.
   Перед первой правкой файла копия кладётся в папку программы _backup.
9. «Авто↺» — сама перечитывает файл, когда он изменился. Кнопка языка — RU → EN → 中文.

Если что-то не так:
• без Excel на машине: .xls/.xlsb и файлы с диаграммами/сводными — только просмотр;
• обновление программа предлагает сама при запуске."""

# ── языки (2.0.2, 23.10.2026): RU → EN → 中文 по кругу; нет перевода — показываем оригинал ──
LANGS = ("ru", "en", "zh")
LANG_NEXT = {"ru": "en", "en": "zh", "zh": "ru"}
LANG_BTN = {"ru": "EN", "en": "中文", "zh": "RU"}       # на кнопке — язык, на который переключишь
LANG = "ru"

TR = {  # русский → (английский, китайский)
    "Файл:": ("File:", "文件:"),
    "Обзор": ("Browse", "浏览"),
    "Загрузить": ("Load", "加载"),
    "↺ Обновить": ("↺ Reload", "↺ 刷新"),
    "Лист": ("Sheet", "工作表"),
    "Вся книга": ("Whole workbook", "整个工作簿"),
    "Лист:": ("Sheet:", "工作表:"),
    "Авто↺": ("Auto↺", "自动↺"),
    "Файл не загружен": ("No file loaded", "未加载文件"),
    "Поиск:": ("Search:", "搜索:"),
    "Регистр": ("Match case", "区分大小写"),
    "✕ Очистить": ("✕ Clear", "✕ 清除"),
    "Колонки": ("Columns", "列"),
    "Показать все колонки / скрыть все": ("Show all columns / hide all", "显示全部列 / 全部隐藏"),
    "Все": ("All", "全部"),
    "Нет загруженного файла": ("No file loaded", "未加载文件"),
    "Загрузка…": ("Loading…", "加载中…"),
    "Ошибка: {exc}": ("Error: {exc}", "错误: {exc}"),
    "{n} строк · {c} колонок  |  {f}": ("{n} rows · {c} columns  |  {f}", "{n} 行 · {c} 列  |  {f}"),
    "только просмотр (.xls/.xlsb)": ("view only (.xls/.xlsb)", "仅查看 (.xls/.xlsb)"),
    "только просмотр": ("view only", "仅查看"),
    "только просмотр (в файле диаграммы/сводные/рисунки)":
        ("view only (charts/pivots/pictures in file)", "仅查看（文件含图表/数据透视表/图片）"),
    "{n} совпадений  |  {f}": ("{n} matches  |  {f}", "{n} 条匹配  |  {f}"),
    "строк {n} / {total}": ("rows {n} / {total}", "行 {n} / {total}"),
    "Скопировано: {t}": ("Copied: {t}", "已复制: {t}"),
    "зелёный": ("green", "绿色"),
    "жёлтый": ("yellow", "黄色"),
    "красный": ("red", "红色"),
    "оранжевый": ("orange", "橙色"),
    "Залить выделенные ячейки: {name} (ещё раз — снять заливку)":
        ("Fill selected cells: {name} (again — remove fill)", "填充所选单元格: {name}（再点一次 — 取消填充）"),
    "Выделите ячейки, потом нажмите цвет": ("Select cells, then press a colour", "请先选择单元格，再点击颜色"),
    "Цвет": ("Colour", "颜色"),
    "Файл открыт в Excel — закройте его и повторите.":
        ("The file is open in Excel — close it and try again.", "文件已在 Excel 中打开 — 请关闭后重试。"),
    "Не записал в файл: {exc}": ("Not written to file: {exc}", "未写入文件: {exc}"),
    "Залито: {n} яч.": ("Filled: {n} cells", "已填充: {n} 个单元格"),
    "Заливка снята: {n} яч.": ("Fill removed: {n} cells", "已取消填充: {n} 个单元格"),
    " (только в программе — файл не правится)": (" (in the program only — file not changed)", "（仅在程序中 — 文件未修改）"),
    "в файле колонка {x} называется «{h}», а не «{c}»":
        ("column {x} in the file is «{h}», not «{c}»", "文件中第 {x} 列为「{h}」，而不是「{c}」"),
    "в файле в этой ячейке «{a}», а в таблице «{b}» — нажмите «Обновить»":
        ("the file has «{a}» in this cell, the table has «{b}» — press «Reload»",
         "文件中该单元格为「{a}」，表格中为「{b}」— 请点击「刷新」"),
    "файл изменился после загрузки — нажмите «Обновить» и повторите":
        ("the file changed after loading — press «Reload» and try again", "加载后文件已更改 — 请点击「刷新」后重试"),
    "Правка": ("Edit", "编辑"),
    "Файл открыт в Excel — закройте его и повторите правку.":
        ("The file is open in Excel — close it and repeat the edit.", "文件已在 Excel 中打开 — 请关闭后重新编辑。"),
    "Не записал: {exc}": ("Not written: {exc}", "未写入: {exc}"),
    "Записано в файл: {col} = {v}  (строка {r})": ("Written to file: {col} = {v}  (row {r})", "已写入文件: {col} = {v}（第 {r} 行）"),
    "Excel: открываю файл…": ("Excel: opening the file…", "Excel: 正在打开文件…"),
    "Обновление": ("Update", "更新"),
    "Прошлое обновление не встало: {msg}": ("The previous update failed: {msg}", "上次更新未成功: {msg}"),
    "Доступна версия {tag} (у вас {v}).\n\nСкачать и обновить сейчас? Программа закроется и откроется заново.":
        ("Version {tag} is available (you have {v}).\n\nDownload and update now? The program will close and reopen.",
         "有新版本 {tag}（当前 {v}）。\n\n现在下载并更新吗？程序将关闭并重新打开。"),
    "Не скачалось: {exc}": ("Download failed: {exc}", "下载失败: {exc}"),
    "Язык: RU → EN → 中文": ("Language: RU → EN → 中文", "语言: RU → EN → 中文"),
}

# Заголовки колонок из выгрузок (собраны с файлов «Запчасти» на 110, 23.10.2026): китайский →
# (русский, английский). Показ — только на экране, в файле и в настройках колонка зовётся как есть.
COL_TR = {
    "序列号": ("Серийный номер", "SN"), "品名": ("Наименование", "Item name"),
    "库位": ("Ячейка", "Location"), "区域类型编码": ("Код зоны", "Zone code"),
    "区域类型名称": ("Зона", "Zone"), "物料编码": ("Код материала", "Part number"),
    "物料代码": ("Код материала", "Part number"), "物料名称": ("Название материала", "Part name"),
    "备注": ("Примечание", "Remark"), "数量": ("Количество", "Quantity"),
    "申请原因": ("Причина заявки", "Request reason"), "申请事由": ("Основание заявки", "Request grounds"),
    "申请物料SN（领）": ("SN выданного материала", "Issued part SN"),
    "申请物料名称（领）": ("Выданный материал", "Issued part name"),
    "申请物料编码（领）": ("Код выданного материала", "Issued part number"),
    "更换物料SN（退）": ("SN возвращённого материала", "Returned part SN"),
    "更换物料名称（退）": ("Возвращённый материал", "Returned part name"),
    "更换物料编码（退）": ("Код возвращённого материала", "Returned part number"),
    "生成的业务单号": ("Номер созданного документа", "Generated document no."),
    "状态归类保内": ("Статус (гарантия)", "Status (in warranty)"), "系统内字段": ("Системное поле", "System field"),
    "申请时间": ("Время заявки", "Request time"), "申请人": ("Заявитель", "Applicant"),
    "替代料参考": ("Замена (справочно)", "Substitute reference"), "申请数量": ("Количество в заявке", "Requested qty"),
    "PO号": ("Номер PO", "PO no."), "SO号": ("Номер SO", "SO no."), "状态": ("Статус", "Status"),
    "结算管理": ("Расчёты", "Settlement"), "承诺交期时间": ("Обещанный срок поставки", "Promised delivery"),
    "费用承担人工号": ("Табельный плательщика", "Cost bearer ID"), "搬运单号": ("Номер перемещения", "Transfer no."),
    "发料时间": ("Время выдачи", "Issue time"), "子库位": ("Подъячейка", "Sub-location"),
    "通知仓库打单日期": ("Дата заявки на склад", "Warehouse notice date"), "计划状态": ("Статус плана", "Plan status"),
    "物料类型": ("Тип материала", "Part type"), "发货单号": ("Номер отгрузки", "Shipment no."),
    "客服组": ("Группа сервиса", "Service group"), "入库时间": ("Время прихода", "Receipt time"),
    "创建时间": ("Создано", "Created"), "更新时间": ("Обновлено", "Updated"),
    "组织名称": ("Организация", "Organization"), "所属仓库": ("Склад", "Warehouse"),
    "货主名称": ("Владелец груза", "Cargo owner"), "批次编码": ("Код партии", "Batch code"),
    "库区编码": ("Код зоны склада", "Area code"), "库区名称": ("Зона склада", "Area name"),
    "库位编码": ("Код ячейки", "Location code"), "容器": ("Контейнер", "Container"),
    "单位名称": ("Единица", "Unit"), "供应商名称": ("Поставщик", "Supplier"),
    "供应商批号": ("Партия поставщика", "Supplier batch"), "批次状态": ("Статус партии", "Batch status"),
    "生产日期": ("Дата производства", "Production date"), "有效日期": ("Годен до", "Expiry date"),
    "下次质检日期": ("Следующая проверка", "Next QC date"), "库存数量": ("Остаток", "Stock qty"),
    "保留数量": ("Зарезервировано", "Reserved qty"), "待定数量": ("Ожидает", "Pending qty"),
    "可用数量": ("Доступно", "Available qty"), "新品数量": ("Новых", "New qty"),
    "物料：长": ("Длина", "Length"), "物料：宽": ("Ширина", "Width"), "物料：高": ("Высота", "Height"),
    "物料：重量": ("Вес", "Weight"), "物料：标准装箱量": ("Штук в коробке", "Standard pack qty"),
    "物料类别": ("Категория материала", "Part category"), "是否贵重": ("Ценный", "Valuable"),
    "价格区间": ("Ценовой диапазон", "Price range"), "销售状态": ("Статус продаж", "Sales status"),
}


def tr(ru, **kw):
    """Текст интерфейса на текущем языке. Нет перевода — оригинал (по-русски), никогда не падаем."""
    s = ru
    if LANG != "ru":
        pair = TR.get(ru)
        s = pair[LANGS.index(LANG) - 1] if pair else ru
    try:
        return s.format(**kw) if kw else s
    except (KeyError, IndexError, ValueError):
        return ru.format(**kw) if kw else ru


def col_title(name):
    """Заголовок колонки на экране: перевод, если знаем; в 中文 и без перевода — как в файле."""
    pair = COL_TR.get(str(name).strip())
    if not pair or LANG == "zh":
        return str(name)
    return pair[0] if LANG == "ru" else pair[1]

HELP_EN = """Quick search in any Excel file (.xlsx / .xls / .xlsb).

1. «Browse» — pick a file (or one of the recent ones in the «File» box).
2. «Sheet» / «Whole workbook» — search one sheet or all of them.
3. Left — columns: the tick shows a column, the box next to it filters by that column.
   The «Columns» tick shows all / hides all.
4. «Search» — searches all visible columns at once, matches are highlighted.
5. Column ✓ — a note for yourself; remembered per file, not written to the file.
6. Click a header — sort; click a cell — its value is copied.
7. Column width: drag the header border; double-click it — fit to the widest value.
8. Edit: double-click a cell or F2, Enter — the value is written to the same cell of the file
   (through Excel if installed; if the file is open in Excel you see the edit right there).
   Colour: select cells and press a colour square next to the search — the cells are filled
   in the program and in the file. The same colour again — fill removed.
   Before the first edit a copy of the file goes to the program's _backup folder.
9. «Auto↺» — reloads the file when it changes. The language button — RU → EN → 中文.

If something is wrong:
• without Excel: .xls/.xlsb and files with charts/pivots — view only;
• the program offers updates itself at start."""

HELP_ZH = """快速搜索任意 Excel 文件（.xlsx / .xls / .xlsb）。

1. 「浏览」— 选择文件（或在「文件」框中选最近的文件）。
2. 「工作表」/「整个工作簿」— 在一个或全部工作表中搜索。
3. 左侧为列：勾选显示该列，旁边的框按该列筛选。「列」勾选 — 全部显示 / 全部隐藏。
4. 「搜索」— 同时搜索所有可见列，匹配项高亮。
5. ✓ 列 — 自己的标记，按文件记住，不写入文件。
6. 点击表头 — 排序；点击单元格 — 复制其值。
7. 列宽：拖动表头边界；双击边界 — 按最宽内容自适应。
8. 编辑：双击单元格或 F2，回车 — 写入文件的同一单元格（有 Excel 时通过 Excel 写入；
   文件已在 Excel 中打开时可直接看到修改）。
   颜色：选择单元格后点击搜索框右侧的色块 — 程序和文件中都会填充；再点同一颜色 — 取消填充。
   第一次编辑前，文件副本保存在程序的 _backup 文件夹。
9. 「自动↺」— 文件变化时自动重新加载。语言按钮 — RU → EN → 中文。

如有问题：
• 没有 Excel 时：.xls/.xlsb 及含图表/数据透视表的文件仅可查看；
• 程序启动时会自动提示更新。"""

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
QHeaderView::section:checked {{ background: {C['surface1']}; color: #2f80ff; }}
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
        if orientation != Qt.Horizontal or role not in (Qt.DisplayRole, Qt.ToolTipRole):
            return None
        if role == Qt.ToolTipRole and section > 0:
            return str(self.cols[section - 1])           # оригинал заголовка из файла
        name = CHECK if section == 0 else self.cols[section - 1]
        arrow = (" ▲" if self.app.sort_asc else " ▼") if name == self.app.sort_col else ""
        return (name if section == 0 else col_title(name)) + arrow

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
        if role in (Qt.BackgroundRole, Qt.ForegroundRole):
            rgb = self.app.cell_colors.get((self.app.keys.at[self.idx[r]], self.cols[c - 1]))
            if rgb:                           # залита здесь же — тот же цвет, что в файле
                return QColor(*rgb) if role == Qt.BackgroundRole else QColor("#1e1e2e")
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
            # ✓ — просто отметка для себя (в настройках программы), в файл не пишется
            self.app.set_checked(self.app.keys.at[self.idx[r]], Qt.CheckState(value) == Qt.Checked)
            self.dataChanged.emit(index, index)
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


SEL_COLOR = QColor("#2f80ff")            # рамка выделенных ячеек: синий — не совпадает ни с одним цветом заливки
CROSS = QColor(137, 180, 250, 38)        # «прицел»: строка и колонка текущей ячейки


class CellDelegate(QStyledItemDelegate):
    """Выделение не закрашивает ячейку (её заливку видно), а обводит рамкой; строка и колонка
    текущей ячейки слегка подсвечены по всей таблице — сразу видно, где ты в массиве."""

    def __init__(self, view):
        super().__init__(view)
        self.view = view

    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        selected = bool(opt.state & QStyle.State_Selected)
        opt.state &= ~(QStyle.State_Selected | QStyle.State_HasFocus)
        cur = self.view.currentIndex()
        style = self.view.style()
        bg = index.data(Qt.BackgroundRole)
        if isinstance(bg, QColor) and index.column() > 0:
            # залитая ячейка: фон и тёмный текст сами — стиль окна красит текст светлым
            painter.fillRect(opt.rect, bg)
            painter.save()
            painter.setPen(QColor(index.data(Qt.ForegroundRole) or "#1e1e2e"))
            # та же область текста, что у обычной ячейки — иначе подгон ширины не сходится
            r = style.subElementRect(QStyle.SE_ItemViewItemText, opt, self.view).adjusted(3, 0, 0, 0)
            painter.drawText(r, int(Qt.AlignVCenter | Qt.AlignLeft),
                             opt.fontMetrics.elidedText(opt.text, Qt.ElideRight, r.width()))
            painter.restore()
        else:
            style.drawControl(QStyle.CE_ItemViewItem, opt, painter, self.view)
        if cur.isValid() and (index.row() == cur.row() or index.column() == cur.column()):
            painter.fillRect(opt.rect, CROSS)
        if selected:
            painter.save()
            painter.setPen(QPen(SEL_COLOR, 3))
            painter.drawRect(opt.rect.adjusted(2, 2, -2, -2))
            painter.restore()


def theme_colors(wb):
    """Цвета темы книги в порядке индексов Excel (0 lt1, 1 dk1, 2 lt2, 3 dk2, 4-9 акценты)."""
    import re
    xml = wb.loaded_theme.decode("utf-8", "replace") if wb.loaded_theme else ""
    m = re.search(r"<a:clrScheme.*?</a:clrScheme>", xml, re.S)
    if not m:
        return []
    names = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6")
    got = {}
    for n in names:
        mm = re.search(rf"<a:{n}>.*?(?:srgbClr val|lastClr)=\"([0-9A-Fa-f]{{6}})\"", m.group(0), re.S)
        got[n] = mm.group(1) if mm else "FFFFFF"
    order = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6")
    return [got[n] for n in order]


def fill_rgb(color, theme):
    """openpyxl Color → (r, g, b) или None. rgb, индексная палитра, тема (+ оттенок tint)."""
    import colorsys
    from openpyxl.styles.colors import COLOR_INDEX
    try:
        if color.type == "rgb" and isinstance(color.rgb, str):
            h = color.rgb[-6:]
        elif color.type == "indexed" and color.indexed is not None and color.indexed < len(COLOR_INDEX):
            h = COLOR_INDEX[color.indexed][-6:]
        elif color.type == "theme" and color.theme is not None and color.theme < len(theme):
            h = theme[color.theme]
        else:
            return None
        r, g, b = (int(h[k:k + 2], 16) for k in (0, 2, 4))
        t = float(color.tint or 0)
        if t:
            hh, ll, ss = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
            ll = ll * (1 + t) if t < 0 else ll * (1 - t) + t
            r, g, b = (round(x * 255) for x in colorsys.hls_to_rgb(hh, ll, ss))
        return (r, g, b)
    except Exception:
        return None


def read_fills(path, sheets):
    """Заливка ячеек из файла (.xlsx/.xlsm): {(лист, строка, колонка): (r, g, b)}; строка 1 — заголовок."""
    import openpyxl
    out = {}
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        theme = theme_colors(wb)
        for name in sheets:
            if name not in wb.sheetnames:
                continue
            for r, row in enumerate(wb[name].iter_rows(min_row=2), start=2):
                for c, cell in enumerate(row, start=1):
                    f = getattr(cell, "fill", None)
                    if f is not None and f.fill_type == "solid":
                        rgb = fill_rgb(f.fgColor, theme)
                        if rgb and rgb != (255, 255, 255):
                            out[(name, r, c)] = rgb
    finally:
        wb.close()
    return out


class Bus(QObject):
    """Сигналы из фоновых потоков в окно (Qt-виджеты трогать только из главного потока)."""
    update_found = Signal(str, str, int)
    message = Signal(str, str)
    fills_ready = Signal(str, object)     # путь, {(лист, строка Excel, колонка Excel): (r, g, b)}


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
        self.cell_colors = {}                 # (ключ строки, колонка) → (r, g, b) заливки ячейки
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
        self.bus.fills_ready.connect(self._on_fills)
        global LANG
        LANG = self.settings.get("lang") if self.settings.get("lang") in LANGS else "ru"
        self._i18n = []                       # (виджет, метод, русский текст) — перерисовка при смене языка
        self._build_ui()
        self.auto_timer = QTimer(self)
        self.auto_timer.timeout.connect(self._auto_tick)
        # файл в командной строке («Открыть склад SN» из GASP Scrap, 23.10.2026) главнее последнего
        arg = next((a for a in sys.argv[1:] if os.path.isfile(a)), None)
        start = arg or self.settings.get("last_file") or DEFAULT_EXCEL
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
        top.addWidget(self._t(QLabel(), "Файл:"))
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
            b = self._t(QPushButton(), text)
            b.clicked.connect(fn)
            top.addWidget(b)
        self.rb_sheet, self.rb_all = self._t(QRadioButton(), "Лист"), self._t(QRadioButton(), "Вся книга")
        self.rb_sheet.setChecked(True)
        self.rb_sheet.toggled.connect(lambda _on: self.current_file and self.load_sheet())
        top.addWidget(self.rb_sheet)
        top.addWidget(self.rb_all)
        top.addWidget(self._t(QLabel(), "Лист:"))
        self.sheet_cb = QComboBox()
        self.sheet_cb.setMinimumWidth(150)
        self.sheet_cb.activated.connect(lambda _i: self.current_file and self.load_sheet())
        top.addWidget(self.sheet_cb)
        self.auto_chk = self._t(QCheckBox(), "Авто↺")
        self.auto_chk.toggled.connect(self._toggle_auto)
        top.addWidget(self.auto_chk)
        top.addStretch(1)
        self.status = QLabel(tr("Файл не загружен"))
        self.status.setStyleSheet(f"color: {C['green']};")
        top.addWidget(self.status)
        lay.addLayout(top)

        bar = QHBoxLayout()
        lbl = self._t(QLabel(), "Поиск:")
        lbl.setStyleSheet("font-size: 11pt;")
        bar.addWidget(lbl)
        self.query = QLineEdit()
        self.query.setFixedWidth(250)
        self.query.setStyleSheet("font-size: 12pt;")
        self.query.textChanged.connect(self.search)
        bar.addWidget(self.query)
        self.case_chk = self._t(QCheckBox(), "Регистр")
        self.case_chk.toggled.connect(self.search)
        bar.addWidget(self.case_chk)
        clr = self._t(QPushButton(), "✕ Очистить")
        clr.clicked.connect(lambda: self.query.setText(""))
        bar.addWidget(clr)
        bar.addStretch(1)
        for name, rgb in MARK_COLORS:
            bar.addWidget(self._swatch(name, rgb))
        bar.addSpacing(12)
        # на месте «Экспорт» (правки и цвета теперь в самом файле) — язык: RU → EN → 中文
        self.lang_btn = QPushButton(LANG_BTN[LANG])
        self.lang_btn.setFixedWidth(52)
        self._t(self.lang_btn, "Язык: RU → EN → 中文", "setToolTip")
        self.lang_btn.clicked.connect(self._switch_lang)
        bar.addWidget(self.lang_btn)
        hlp = QPushButton("?")
        hlp.setFixedWidth(32)
        hlp.clicked.connect(lambda: QMessageBox.information(
            self, f"Excel Search {VERSION}", {"ru": HELP_TEXT, "en": HELP_EN, "zh": HELP_ZH}[LANG]))
        bar.addWidget(hlp)
        lay.addLayout(bar)

        self.split = QSplitter(Qt.Horizontal)
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(4, 4, 4, 4)
        hdr = QHBoxLayout()
        self.all_chk = self._t(QCheckBox(), "Колонки")      # нажал — все колонки, ещё раз — ни одной
        self.all_chk.setStyleSheet("font-weight: bold;")
        self._t(self.all_chk, "Показать все колонки / скрыть все", "setToolTip")
        self.all_chk.clicked.connect(lambda: self._toggle_all(not self._all_visible()))
        hdr.addWidget(self.all_chk)
        hdr.addStretch(1)
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
        self.view.setItemDelegate(CellDelegate(self.view))
        self.view.horizontalHeader().setHighlightSections(True)
        self.view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        h = self.view.horizontalHeader()
        h.setSectionResizeMode(QHeaderView.Interactive)
        h.setSectionsClickable(True)
        h.sectionClicked.connect(self._on_header_click)
        h.sectionResized.connect(self._on_col_resized)
        # двойной клик по границе — свой подгон (по всем строкам, снимает ручную ширину);
        # штатный подгон Qt (первые 500 строк) отключаем
        h.sectionHandleDoubleClicked.disconnect()
        h.sectionHandleDoubleClicked.connect(self._on_col_fit)
        self.view.clicked.connect(self._on_cell_click)
        self.view.selectionModel().currentChanged.connect(lambda *_: self.view.viewport().update())
        rl.addWidget(self.view)
        fb = QHBoxLayout()
        fx = QLabel("fx")
        fx.setStyleSheet(f"color: {C['overlay']}; font-style: italic;")
        fb.addWidget(fx)
        self.fbar = QLineEdit()
        self.fbar.setReadOnly(True)
        fb.addWidget(self.fbar)
        # справа — сколько строк показано из скольких с данными (23.10.2026, просьба пользователя)
        self.rows_lbl = QLabel("")
        self.rows_lbl.setStyleSheet(f"color: {C['blue']}; padding: 0 6px;")
        fb.addWidget(self.rows_lbl)
        rl.addLayout(fb)
        self.split.addWidget(right)
        self.split.setSizes([self.settings.get("col_panel_width", 230), 1000])
        self.split.splitterMoved.connect(self._on_split)
        lay.addWidget(self.split, 1)

    # ── файлы ──
    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel", os.path.dirname(self.current_file or ""),
                                              "Excel (*.xlsx *.xlsm *.xls *.xlsb);;" + tr("Все") + " (*.*)")
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
            self.status.setText(tr("Нет загруженного файла"))

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
        self.status.setText(tr("Загрузка…"))
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
            self.status.setText(tr("Ошибка: {exc}", exc=exc))
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
            self.cell_colors = {tuple(k.split("\x1f", 1)): tuple(v) for k, v in
                                self.settings.get("cell_colors", {}).get(self.current_file, {}).items()}
            self._load_fills()
            self.can_edit, why = self._editable()
            self._rebuild_columns([c for c in self.df.columns if c != ROW_COL])
            self.search()
            self.status.setText(tr("{n} строк · {c} колонок  |  {f}", n=len(self.df), c=len(self.df.columns) - 1,
                                   f=os.path.basename(self.current_file)) + ("" if self.can_edit else f"  ·  {why}"))
        except Exception as exc:
            self.status.setText(tr("Ошибка: {exc}", exc=exc))
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
            return False, tr("только просмотр (.xls/.xlsb)")
        try:
            names = zipfile.ZipFile(f).namelist()
        except Exception:
            return False, tr("только просмотр")
        if any(k in n for n in names for k in ("xl/charts/", "xl/pivot", "xl/drawings/", "xl/media/")):
            return False, tr("только просмотр (в файле диаграммы/сводные/рисунки)")
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
            cb = QCheckBox(col_title(col))
            cb.setToolTip(str(col))
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
        self._sync_all_chk()

    def _all_visible(self):
        return bool(self.col_checks) and all(cb.isChecked() for cb in self.col_checks.values())

    def _sync_all_chk(self):
        on = [cb.isChecked() for cb in self.col_checks.values()]
        st = Qt.Checked if on and all(on) else Qt.Unchecked if not any(on) else Qt.PartiallyChecked
        self.all_chk.setCheckState(st)

    def _on_col_toggle(self):
        self.settings["col_visibility"] = {c: cb.isChecked() for c, cb in self.col_checks.items()}
        save_settings(self.settings)
        self._sync_all_chk()
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
        mask = self._mask(cols) if cols else None
        if not cols:
            self.model.reset([], [], self.text, "", False)
            self.rows_lbl.setText(tr("строк {n} / {total}", n=0, total=len(self.df)))
            return
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
        self.status.setText(tr("{n} совпадений  |  {f}", n=int(mask.sum()), f=os.path.basename(self.current_file or "")))
        # «с данными» — строка, где хоть одна ячейка не пустая (пустые хвосты листа не в счёт)
        has = (self.text.drop(columns=[c for c in (SHEET_COL, ROW_COL) if c in self.text.columns])
               != "").any(axis=1)
        self.rows_lbl.setText(tr("строк {n} / {total}", n=int(has.loc[list(idx)].sum()), total=int(has.sum())))

    # ── язык ──
    def _t(self, w, ru, method="setText"):
        """Надпись виджета на текущем языке + запомнить её для смены языка."""
        getattr(w, method)(tr(ru))
        self._i18n.append((w, method, ru))
        return w

    def _switch_lang(self):
        global LANG
        LANG = LANG_NEXT[LANG]
        self.settings["lang"] = LANG
        save_settings(self.settings)
        for w, method, ru in self._i18n:
            try:
                getattr(w, method)(ru() if callable(ru) else tr(ru))
            except RuntimeError:              # виджет уже удалён (колонки пересобраны)
                pass
        self.lang_btn.setText(LANG_BTN[LANG])
        for col, cb in self.col_checks.items():
            cb.setText(col_title(col))
        self.model.headerDataChanged.emit(Qt.Horizontal, 0, max(self.model.columnCount() - 1, 0))
        if self.df is not None:
            self.search()
        else:
            self.status.setText(tr("Файл не загружен"))

    # ── ширина колонок: подгон под содержимое + ручная/двойной клик, запоминается по файлу ──
    def _saved_widths(self):
        return self.settings.get("col_widths", {}).get(self.current_file or "", {})

    def _fit_width(self, i):
        """Ширина колонки i по самому широкому значению среди ВСЕХ показанных строк (Qt смотрит
        только первые 500) и по заголовку. Меряем 30 самых длинных строк — быстро и на 100 тыс."""
        name = self.model.cols[i - 1]
        fm = QFontMetrics(QFont("Segoe UI", 9))
        s = self.text.loc[self.model.idx, name] if len(self.model.idx) else self.text[name].iloc[:0]
        top = s.loc[s.str.len().nlargest(30).index] if len(s) else []
        cell = max((fm.horizontalAdvance(v) for v in top), default=0) + 16
        head = QFontMetrics(QFont("Segoe UI", 9, QFont.Bold)).horizontalAdvance(col_title(name) + " ▲") + 20
        return max(cell, head, 40)

    def _apply_widths(self):
        """Открытие: колонка подогнана по содержимому (не шире 500 — длинный текст не растягивает её
        на весь экран), а потянутая руками — своей запомненной шириной."""
        self._applying_widths = True
        try:
            h = self.view.horizontalHeader()
            h.resizeSection(0, 34)
            saved = self._saved_widths()
            for i, name in enumerate(self.model.cols, start=1):
                w = saved.get(name)
                h.resizeSection(i, int(w) if w else min(self._fit_width(i), 500))
        finally:
            self._applying_widths = False

    def _on_col_fit(self, logical):
        """Двойной клик по границе: подогнать по самому широкому и снять ручную ширину."""
        if logical == 0 or logical > len(self.model.cols):
            return
        self._applying_widths = True
        try:
            self.view.horizontalHeader().resizeSection(logical, self._fit_width(logical))
        finally:
            self._applying_widths = False
        store = self.settings.setdefault("col_widths", {}).setdefault(self.current_file or "", {})
        if store.pop(self.model.cols[logical - 1], None) is not None:
            save_settings(self.settings)

    def _on_col_resized(self, logical, _old, new):
        """Ширину потянули руками — запомнить для этого файла."""
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
            self.status.setText(tr("Скопировано: {t}", t=text[:80]))

    # ── отметки ✓ ──
    def set_checked(self, key, on):
        (self.checked.add if on else self.checked.discard)(key)
        self.settings.setdefault("checked", {})[self.current_file or ""] = sorted(self.checked)
        save_settings(self.settings)

    # ── заливка ячеек: выделил → цвет (как в Excel) ──
    def _swatch(self, name, rgb):
        b = QPushButton()
        b.setFixedSize(34, 24)
        b.setToolTip(tr("Залить выделенные ячейки: {name} (ещё раз — снять заливку)", name=tr(name)))
        self._i18n.append((b, "setToolTip", lambda n=name: tr(
            "Залить выделенные ячейки: {name} (ещё раз — снять заливку)", name=tr(n))))
        hexc = "#%02x%02x%02x" % rgb
        b.setStyleSheet(f"QPushButton {{ background: {hexc}; border: 1px solid {C['surface1']}; "
                        f"border-radius: 3px; padding: 0; }}")
        b.clicked.connect(lambda _=False, c=rgb: self.color_cells(c))
        return b

    def _load_fills(self):
        """Заливка из самого файла — в фоне, таблица уже показана. .xls/.xlsb — не читается."""
        path = self.current_file
        if not path or not path.lower().endswith(EDITABLE_EXT):
            return
        if SHEET_COL in self.df.columns:
            sheets = list(dict.fromkeys(self.df[SHEET_COL]))
        else:
            sheets = [self.sheet_cb.currentText()]

        def worker():
            try:
                fills = read_fills(path, sheets)
            except Exception:
                return
            self.bus.fills_ready.emit(path, fills)
        threading.Thread(target=worker, daemon=True).start()

    def _on_fills(self, path, fills):
        if path != self.current_file or self.df is None:
            return
        cols = [c for c in self.df.columns if c not in (SHEET_COL, ROW_COL)]
        sheet_of = (self.df[SHEET_COL] if SHEET_COL in self.df.columns
                    else pd.Series(self.sheet_cb.currentText(), index=self.df.index))
        where = {(sheet_of.at[i], int(self.df.at[i, ROW_COL])): i for i in self.df.index}
        colors = {}
        for (sheet, r, c), rgb in fills.items():
            i = where.get((sheet, r))
            if i is not None and c <= len(cols):
                colors[(self.keys.at[i], cols[c - 1])] = rgb
        self.cell_colors = colors          # файл — источник правды о заливке
        self.model.layoutChanged.emit()

    def _save_cell_colors(self):
        self.settings.setdefault("cell_colors", {})[self.current_file or ""] = {
            f"{k}\x1f{c}": list(v) for (k, c), v in self.cell_colors.items()}
        save_settings(self.settings)

    def _addr(self, i, col):
        """Лист, строка и колонка ячейки в файле (заголовок — строка 1)."""
        sheet = self.df.at[i, SHEET_COL] if SHEET_COL in self.df.columns else self.sheet_cb.currentText()
        cols = [c for c in self.df.columns if c not in (SHEET_COL, ROW_COL)]
        return sheet, int(self.df.at[i, ROW_COL]), cols.index(col) + 1

    def color_cells(self, rgb):
        """Залить выделенные ячейки цветом rgb в таблице и в файле. Все уже этого цвета — снять заливку."""
        cells = [(self.model.idx[ix.row()], self.model.cols[ix.column() - 1])
                 for ix in self.view.selectionModel().selectedIndexes() if ix.column() > 0]
        if not cells:
            self.status.setText(tr("Выделите ячейки, потом нажмите цвет"))
            return
        keys = [(self.keys.at[i], col) for i, col in cells]
        new = None if all(self.cell_colors.get(k) == rgb for k in keys) else rgb
        if self.can_edit:
            path = self.current_file
            try:
                self._before_write(path)
                by_sheet = {}
                for i, col in cells:
                    sheet, xr, xc = self._addr(i, col)
                    by_sheet.setdefault(sheet, []).append((xr, xc))
                if self.use_excel:
                    for sheet, addrs in by_sheet.items():
                        b, ws = self._excel_sheet(path, sheet)
                        for xr, xc in addrs:
                            ws.range(xr, xc).color = new
                    self._excel_saved(b)
                else:
                    import openpyxl
                    from openpyxl.styles import PatternFill
                    wb = openpyxl.load_workbook(path, keep_vba=path.lower().endswith(".xlsm"))
                    fill = PatternFill("solid", fgColor="%02X%02X%02X" % new) if new else PatternFill()
                    for sheet, addrs in by_sheet.items():
                        for xr, xc in addrs:
                            wb[sheet].cell(xr, xc).fill = fill
                    wb.save(path)
                    self.last_mtime = os.path.getmtime(path)
            except PermissionError:
                QMessageBox.warning(self, tr("Цвет"), tr("Файл открыт в Excel — закройте его и повторите."))
                return
            except Exception as exc:
                QMessageBox.warning(self, tr("Цвет"), tr("Не записал в файл: {exc}", exc=exc))
                return
        for k in keys:
            if new:
                self.cell_colors[k] = new
            else:
                self.cell_colors.pop(k, None)
        self._save_cell_colors()
        self.model.layoutChanged.emit()
        self.status.setText(tr("Залито: {n} яч." if new else "Заливка снята: {n} яч.", n=len(cells))
                            + ("" if self.can_edit else tr(" (только в программе — файл не правится)")))

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
                raise RuntimeError(tr("в файле колонка {x} называется «{h}», а не «{c}»", x=xcol, h=head, c=col))
            cell = ws.cell(xrow, xcol)
            cur = cell.value
            cur_txt = "" if cur is None else str(cur)
            if cur_txt != old_txt and not (isinstance(cur, (int, float)) and old_txt and _same_num(cur, old_txt)):
                if not (hasattr(cur, "isoformat") and old_txt.startswith(str(cur)[:10])):
                    raise RuntimeError(tr("в файле в этой ячейке «{a}», а в таблице «{b}» — нажмите «Обновить»",
                                          a=cur_txt, b=old_txt))
            cell.value = _typed(new, cur)
            wb.save(path)
        except _Done:
            pass
        except PermissionError:
            QMessageBox.warning(self, tr("Правка"), tr("Файл открыт в Excel — закройте его и повторите правку."))
            return False
        except Exception as exc:
            QMessageBox.warning(self, tr("Правка"), tr("Не записал: {exc}", exc=exc))
            return False
        was = self.keys.at[i] in self.checked
        self.df.at[i, col] = _typed(new, None) if new else None
        self.text.at[i, col] = new
        nk = row_keys(self.df.loc[[i]]).iloc[0]
        ok_key = self.keys.at[i]
        moved = {(nk, c): v for (k, c), v in self.cell_colors.items() if k == ok_key}
        if moved:
            self.cell_colors = {kc: v for kc, v in self.cell_colors.items() if kc[0] != ok_key}
            self.cell_colors.update(moved)
            self._save_cell_colors()
        if was:
            self.checked.discard(ok_key)
            self.keys.at[i] = nk
            self.set_checked(nk, True)
        else:
            self.keys.at[i] = nk
        self.last_mtime = os.path.getmtime(path)
        self.status.setText(tr("Записано в файл: {col} = {v}  (строка {r})", col=col_title(col), v=new[:60], r=xrow))
        return True

    def _before_write(self, path):
        if os.path.getmtime(path) != self.last_mtime:
            raise RuntimeError(tr("файл изменился после загрузки — нажмите «Обновить» и повторите"))
        if path not in self.backed_up:
            bdir = os.path.join(APP_DIR, "_backup")
            os.makedirs(bdir, exist_ok=True)
            base, ext = os.path.splitext(os.path.basename(path))
            shutil.copy2(path, os.path.join(bdir, f"{base}_{time.strftime('%Y%m%d_%H%M%S')}{ext}"))
            self.backed_up.add(path)

    def _excel_sheet(self, path, sheet):
        self.status.setText(tr("Excel: открываю файл…"))
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
            raise RuntimeError(tr("в файле колонка {x} называется «{h}», а не «{c}»", x=xcol, h=head, c=col))
        rng = ws.range(xrow, xcol)
        cur = rng.value
        cur_txt = "" if cur is None else str(cur)
        if cur_txt != old_txt and not (isinstance(cur, (int, float)) and old_txt and _same_num(cur, old_txt)):
            if not (hasattr(cur, "isoformat") and old_txt.startswith(str(cur)[:10])):
                raise RuntimeError(tr("в файле в этой ячейке «{a}», а в таблице «{b}» — нажмите «Обновить»",
                                      a=cur_txt, b=old_txt))
        rng.value = _typed(new, cur)
        self._excel_saved(b)

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
                QMessageBox.warning(self, tr("Обновление"), tr("Прошлое обновление не встало: {msg}", msg=msg))
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
        if QMessageBox.question(self, tr("Обновление"), tr(
                "Доступна версия {tag} (у вас {v}).\n\nСкачать и обновить сейчас? Программа закроется и откроется заново.",
                tag=tag, v=VERSION)) == QMessageBox.Yes:
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
            self.bus.message.emit(tr("Обновление"), tr("Не скачалось: {exc}", exc=exc))
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
