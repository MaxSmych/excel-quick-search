#!/usr/bin/env python3
"""
Excel Search — SSC UZ
Reads 维修日报数据.xlsx (or any xlsx), lets you pick columns and search live.
Settings saved to excel_search_settings.json.
"""
import hashlib
import json
import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Base dir for persistent files (settings, default excel): next to the EXE when frozen,
# next to the script otherwise. NOT sys._MEIPASS — that is a throwaway temp dir per run,
# which is why settings/checkmarks weren't kept in the built exe.
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(APP_DIR, "excel_search_settings.json")
DEFAULT_EXCEL  = os.path.join(APP_DIR, "维修日报数据.xlsx")

# Версия и обновления (стандарт GUI-программ, brain 04 §10в, с 1.1.0 — 21.10.2026).
# Репозиторий публичный (MIT) — сборка лежит в его же релизах, токен не нужен.
VERSION      = "1.1.2"
APP_NAME     = "Excel Search"
UPDATE_REPO  = "MaxSmych/excel-quick-search"
UPDATE_ASSET = "Excel-Search.zip"      # без пробела: GitHub меняет пробел в имени ассета на точку
API_LATEST   = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
HELP_TEXT = """Быстрый поиск по любому Excel (.xlsx / .xls / .xlsb).

1. «Обзор» — выбрать файл (или из списка недавних в поле «Файл»).
2. «Лист» / «Вся книга» — искать в одном листе или во всех сразу.
3. Слева — колонки: галочка показывает колонку, поле справа от неё — фильтр по этой колонке.
4. «Поиск» — ищет по всем видимым колонкам сразу, совпадения подсвечиваются.
5. Колонка ✓ — отметить строку «обработано»; отметки запоминаются для каждого файла.
6. Клик по заголовку — сортировка; клик по ячейке — значение копируется в буфер.
7. «Авто↺» — сама перечитывает файл, когда он изменился.
8. «Экспорт» — сохранить найденное в .xlsx или .csv.

Если что-то не так:
• файл не открывается — закройте его в Excel и нажмите «Загрузить»;
• обновление программа предлагает сама при запуске, если на GitHub есть новая версия."""


def _parse_ver(tag):
    """'v1.2.3' -> (1, 2, 3): сравнение версий числами, а не строками."""
    out = []
    for part in str(tag).lstrip("vV").split(".")[:3]:
        num = "".join(ch for ch in part if ch.isdigit())
        out.append(int(num) if num else 0)
    return tuple(out + [0] * (3 - len(out)))


def _resource(name):
    # works both from source and from a PyInstaller onefile bundle
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)

# ── auto-install dependencies ────────────────────────────────────────────────
def _ensure(*packages):
    for pkg in packages:
        mod = pkg.split("[")[0].replace("-", "_")
        try:
            __import__(mod)
        except ImportError:
            print(f"Installing {pkg}…")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

_ensure("pandas", "openpyxl", "tksheet")

import pandas as pd          # noqa: E402
from tksheet import Sheet    # noqa: E402


# ── Settings ────────────────────────────────────────────────────────────────
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_settings(data):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ── Clipboard shortcuts (работает на русской/любой раскладке) ──────────────
# Windows keycodes: A=65, C=67, V=86, X=88
_KEYCODE_EVENT = {65: "<<SelectAll>>", 67: "<<Copy>>", 86: "<<Paste>>", 88: "<<Cut>>"}

def _bind_clipboard_shortcuts(widget):
    def handler(ev):
        if not (ev.state & 0x4):  # Ctrl not held
            return
        evt = _KEYCODE_EVENT.get(ev.keycode)
        if evt:
            if evt == "<<SelectAll>>":
                try:
                    widget.select_range(0, "end")
                    widget.icursor("end")
                except Exception:
                    pass
            else:
                widget.event_generate(evt)
            return "break"
    widget.bind("<Control-KeyPress>", handler)


# ── Catppuccin Mocha palette ─────────────────────────────────────────────────
C = {
    "base":    "#1e1e2e",
    "mantle":  "#181825",
    "surface0":"#313244",
    "surface1":"#45475a",
    "text":    "#cdd6f4",
    "blue":    "#89b4fa",
    "green":   "#a6e3a1",
    "overlay": "#6c7086",
}


# ── Main App ─────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # заставка со знаком SM — только в собранном exe (при разработке SPLASH=1)
        self._splash = None
        if getattr(sys, "frozen", False) or os.environ.get("SPLASH") == "1":
            import splash
            self.withdraw()
            self._splash = splash.Splash(self)
        self.title(f"Excel Search — SSC UZ  ·  {VERSION}")
        self.geometry("1280x720")
        self.minsize(800, 500)
        try:
            self.iconbitmap(_resource("icon.ico"))
        except Exception:
            pass

        self.settings   = load_settings()
        self.df          = None
        self._keys       = None   # pd.Series: df.index -> stable row hash
        self._display_keys = []   # row hash per displayed table row
        self._display_cols = []   # clean column names shown (без ✓), для маппинга клика
        self._sort_col   = None   # имя колонки сортировки (клик по заголовку)
        self._sort_asc   = True
        self.current_file = None
        self._auto_stop  = threading.Event()
        self._auto_thread = None
        self._last_mtime  = None

        self._apply_style()
        self._build_ui()

        startup_file = self.settings.get("last_file") or DEFAULT_EXCEL
        if startup_file and os.path.exists(startup_file):
            self.file_var.set(startup_file)
            self._load_file(startup_file)

        if self._splash is not None:
            self._splash.play(self.deiconify)
        self._make_shortcut()
        self.after(1500, self._check_update)

    # ── ttk style ────────────────────────────────────────────────────────────
    def _apply_style(self):
        self.configure(bg=C["base"])
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".",
            background=C["base"], foreground=C["text"],
            fieldbackground=C["surface0"], borderwidth=0)
        s.configure("TFrame",      background=C["base"])
        s.configure("TLabel",      background=C["base"], foreground=C["text"],
                    font=("Segoe UI", 10))
        s.configure("TButton",     background=C["surface0"], foreground=C["text"],
                    font=("Segoe UI", 10), padding=4, relief="flat")
        s.map("TButton",           background=[("active", C["surface1"])])
        s.configure("TCheckbutton",background=C["base"], foreground=C["text"],
                    font=("Segoe UI", 9))
        s.map("TCheckbutton",      background=[("active", C["base"])])
        s.configure("TEntry",      fieldbackground=C["surface0"], foreground=C["text"],
                    insertcolor=C["text"], font=("Segoe UI", 10))
        s.configure("TScrollbar",  background=C["surface0"], troughcolor=C["base"],
                    arrowcolor=C["text"], borderwidth=0)

    # ── UI ───────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # top bar
        top = ttk.Frame(self, padding=6)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Файл:").pack(side=tk.LEFT)
        self.file_var = tk.StringVar()
        self.file_cb = ttk.Combobox(top, textvariable=self.file_var, width=55,
                                    values=self.settings.get("recent_files", []))
        self.file_cb.pack(side=tk.LEFT, padx=4)
        self.file_cb.bind("<<ComboboxSelected>>", lambda e: self._on_recent_pick())
        ttk.Button(top, text="Обзор",    command=self._browse).pack(side=tk.LEFT)
        ttk.Button(top, text="Загрузить",command=self._on_load).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="↺ Обновить",command=self._on_refresh).pack(side=tk.LEFT)

        self.mode_var = tk.StringVar(value="sheet")
        ttk.Radiobutton(top, text="Лист", variable=self.mode_var, value="sheet",
                        command=self._on_mode_change).pack(side=tk.LEFT, padx=(8,0))
        ttk.Radiobutton(top, text="Вся книга", variable=self.mode_var, value="all",
                        command=self._on_mode_change).pack(side=tk.LEFT, padx=(2,4))

        ttk.Label(top, text="Лист:").pack(side=tk.LEFT)
        self.sheet_var = tk.StringVar()
        self.sheet_cb = ttk.Combobox(top, textvariable=self.sheet_var, width=20,
                                     state="readonly")
        self.sheet_cb.pack(side=tk.LEFT, padx=4)
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self._on_sheet_change())

        self.auto_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Авто↺", variable=self.auto_var,
                        command=self._toggle_auto).pack(side=tk.LEFT, padx=8)

        self.status_var = tk.StringVar(value="Файл не загружен")
        ttk.Label(top, textvariable=self.status_var,
                  foreground=C["green"]).pack(side=tk.RIGHT)

        # search bar
        bar = ttk.Frame(self, padding=(6, 0, 6, 6))
        bar.pack(fill=tk.X)

        ttk.Label(bar, text="Поиск:", font=("Segoe UI", 11)).pack(side=tk.LEFT)
        self.query_var = tk.StringVar()
        self.query_var.trace_add("write", lambda *_: self._on_search())
        e = ttk.Entry(bar, textvariable=self.query_var, width=40, font=("Segoe UI", 12))
        e.pack(side=tk.LEFT, padx=4)
        _bind_clipboard_shortcuts(e)
        e.focus()

        self.case_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar, text="Регистр", variable=self.case_var,
                        command=self._on_search).pack(side=tk.LEFT, padx=4)
        ttk.Button(bar, text="✕ Очистить", command=self._clear_search).pack(side=tk.LEFT)
        ttk.Button(bar, text="?", width=3, command=self._show_help).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(bar, text="Экспорт", command=self._export_csv).pack(side=tk.RIGHT)

        # paned window
        pane = tk.PanedWindow(self, orient=tk.HORIZONTAL, bg=C["base"],
                              sashwidth=5, sashrelief="flat")
        pane.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        # ── left: column checkboxes ──────────────────────────────────────────
        left = ttk.Frame(pane, padding=4)
        pane.add(left, width=self.settings.get("col_panel_width", 230), minsize=150)

        hdr = ttk.Frame(left)
        hdr.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(hdr, text="Колонки", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        ttk.Button(hdr, text="Все",  width=4,
                   command=lambda: self._toggle_all(True)).pack(side=tk.RIGHT)
        ttk.Button(hdr, text="Ноль", width=4,
                   command=lambda: self._toggle_all(False)).pack(side=tk.RIGHT, padx=2)

        scroll_frame = ttk.Frame(left)
        scroll_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(scroll_frame, bg=C["base"], highlightthickness=0)
        sb     = ttk.Scrollbar(scroll_frame, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.col_inner = ttk.Frame(canvas)
        win_id = canvas.create_window((0, 0), window=self.col_inner, anchor="nw")

        self.col_inner.bind("<Configure>",
                            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win_id, width=e.width))
        # scroll with mouse wheel only when cursor is over left panel
        canvas.bind("<Enter>", lambda e: canvas.bind_all(
            "<MouseWheel>", lambda ev: canvas.yview_scroll(-1*(ev.delta//120), "units")))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        self.col_vars = {}
        self.col_filters = {}

        # ── right: tksheet table ─────────────────────────────────────────────
        right = ttk.Frame(pane, padding=0)
        pane.add(right, minsize=400)

        self.sheet = Sheet(
            right,
            data=[],
            headers=[],
            show_row_index=False,
            theme="dark",
            # colours to match Catppuccin Mocha
            table_bg=C["mantle"],
            table_fg=C["text"],
            table_grid_fg=C["surface0"],
            table_selected_cells_bg=C["surface1"],
            table_selected_cells_fg=C["text"],
            header_bg=C["surface0"],
            header_fg=C["blue"],
            header_grid_fg=C["surface1"],
            header_selected_columns_bg=C["surface1"],
            header_selected_columns_fg=C["blue"],
            row_height=22,
            header_height=26,
            font=("Segoe UI", 9, "normal"),
            header_font=("Segoe UI", 9, "bold"),
            empty_horizontal=0,
            empty_vertical=0,
        )
        # enable: select, copy, column resize by drag AND double-click auto-fit
        self.sheet.enable_bindings(
            "single_select",
            "row_select",
            "column_select",
            "drag_select",
            "column_width_resize",
            "double_click_column_resize",
            "column_drag_and_drop",
            "arrowkeys",
            "copy",
            "rc_select",
            "right_click_popup_menu",
        )
        self.sheet.bind("<<SheetSelect>>", self._on_cell_select)
        self.sheet.extra_bindings("column_select", self._on_col_header)
        self.sheet.extra_bindings("column_width_resize", self._on_col_resize)
        self.sheet.pack(fill=tk.BOTH, expand=True)

        # formula bar
        fbar = ttk.Frame(right, padding=(0, 2, 0, 0))
        fbar.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Label(fbar, text="fx", foreground=C["overlay"],
                  font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=(0, 4))
        self.fbar_var = tk.StringVar()
        fbar_entry = ttk.Entry(fbar, textvariable=self.fbar_var,
                               font=("Segoe UI", 10), state="normal")
        fbar_entry.pack(fill=tk.X, expand=True)
        fbar_entry.bind("<Key>", lambda e: None if (e.state & 0x4) else "break")
        _bind_clipboard_shortcuts(fbar_entry)

    # ── column checkboxes ────────────────────────────────────────────────────
    def _rebuild_col_checkboxes(self, columns):
        for w in self.col_inner.winfo_children():
            w.destroy()
        self.col_vars.clear()
        self.col_filters.clear()
        saved = self.settings.get("col_visibility", {})
        for col in columns:
            var = tk.BooleanVar(value=saved.get(col, True))
            var.trace_add("write", lambda *_, c=col: self._on_col_toggle(c))
            self.col_vars[col] = var
            fvar = tk.StringVar(value=self.settings.get("col_filters", {}).get(col, ""))
            fvar.trace_add("write", lambda *_: self._on_filter_change())
            self.col_filters[col] = fvar
            row = ttk.Frame(self.col_inner)
            row.pack(fill=tk.X)
            ttk.Checkbutton(row, text=col, variable=var).pack(side=tk.LEFT)
            ttk.Entry(row, textvariable=fvar, width=17,  # 15 знаков SN заглавными
                      font=("Segoe UI", 8)).pack(side=tk.RIGHT, padx=(2, 0))

    def _on_col_toggle(self, _col):
        self._save_col_settings()
        self._on_search()

    def _toggle_all(self, val):
        for v in self.col_vars.values():
            v.set(val)
        self._save_col_settings()
        self._on_search()

    def _save_col_settings(self):
        self.settings["col_visibility"] = {c: v.get() for c, v in self.col_vars.items()}
        save_settings(self.settings)

    def _on_filter_change(self):
        self.settings["col_filters"] = {
            c: fv.get() for c, fv in self.col_filters.items() if fv.get().strip()}
        save_settings(self.settings)
        self._on_search()

    # ── row check-marks (per-file, keyed by full row content) ─────────────────
    def _compute_row_keys(self):
        # Stable hash of the whole source row (ignoring the synthetic _Лист_ col),
        # so a mark stays on its row regardless of what the row holds (numbers, names…).
        cols = [c for c in self.df.columns if c != "_Лист_"]
        # pandas 3: fillna("") не трогает пустые даты (NaT), astype(str) делает из них NaN-float —
        # join падал «sequence item 2: expected str instance, float found» (21.10.2026, выгрузка бэдов).
        # where(notna) обнуляет только пустые ячейки — ключи остальных строк (отметки ✓) прежние.
        sub = self.df[cols].fillna("").astype(str).where(self.df[cols].notna(), "")
        self._keys = sub.apply(
            lambda r: hashlib.md5("\x1f".join(r.values).encode("utf-8")).hexdigest(),
            axis=1)

    def _checked_set(self):
        store = self.settings.setdefault("checked", {})
        return set(store.get(self.current_file or "", []))

    def _set_checked(self, key, on):
        store = self.settings.setdefault("checked", {})
        keys = set(store.get(self.current_file or "", []))
        if on:
            keys.add(key)
        else:
            keys.discard(key)
        store[self.current_file or ""] = sorted(keys)
        save_settings(self.settings)

    def _on_check(self, event):
        try:
            key = self._display_keys[event.row]
        except (IndexError, AttributeError):
            return
        self._set_checked(key, bool(event.value))

    # ── recent files (quick switch in one window) ─────────────────────────────
    def _add_recent(self, path):
        lst = [p for p in self.settings.get("recent_files", []) if p != path]
        lst.insert(0, path)
        self.settings["recent_files"] = lst[:15]
        save_settings(self.settings)
        self.file_cb["values"] = self.settings["recent_files"]

    def _on_recent_pick(self):
        path = self.file_var.get().strip()
        if path:
            self._load_file(path)

    # ── file loading ─────────────────────────────────────────────────────────
    def _browse(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsb"), ("All", "*.*")])
        if path:
            self.file_var.set(path)
            self._load_file(path)

    def _on_load(self):
        path = self.file_var.get().strip()
        if path:
            self._load_file(path)

    def _on_refresh(self):
        if self.current_file:
            self._load_file(self.current_file)
        else:
            self.status_var.set("Нет загруженного файла")

    def _load_file(self, path):
        self.status_var.set("Загрузка…")
        self.update_idletasks()
        try:
            ext = os.path.splitext(path)[1].lower()
            engine = "pyxlsb" if ext == ".xlsb" else None
            xl = pd.ExcelFile(path, engine=engine)
            sheets = xl.sheet_names
            self.sheet_cb["values"] = sheets
            if not self.sheet_var.get() or self.sheet_var.get() not in sheets:
                self.sheet_var.set(sheets[0])
            self.current_file = path
            self._last_mtime  = os.path.getmtime(path)
            self.settings["last_file"] = path
            self._add_recent(path)
            self._load_sheet(xl)
        except Exception as exc:
            self.status_var.set(f"Ошибка: {exc}")

    def _load_sheet(self, xl=None):
        try:
            if xl is None:
                ext = os.path.splitext(self.current_file)[1].lower()
                xl = pd.ExcelFile(self.current_file, engine="pyxlsb" if ext == ".xlsb" else None)
            if self.mode_var.get() == "all":
                frames = []
                for name in xl.sheet_names:
                    df = xl.parse(name)
                    df.insert(0, "_Лист_", name)
                    frames.append(df)
                self.df = pd.concat(frames, ignore_index=True)
                self.sheet_cb.config(state="disabled")
            else:
                self.df = xl.parse(self.sheet_var.get())
                self.sheet_cb.config(state="readonly")
            self._compute_row_keys()
            self._rebuild_col_checkboxes(list(self.df.columns))
            self._on_search()
            self.status_var.set(
                f"{len(self.df)} строк · {len(self.df.columns)} колонок  |  "
                f"{os.path.basename(self.current_file)}")
        except Exception as exc:
            self.status_var.set(f"Ошибка: {exc}")

    def _on_sheet_change(self):
        if self.current_file:
            self._load_sheet()

    def _on_mode_change(self):
        if self.current_file:
            self._load_sheet()

    # ── search ────────────────────────────────────────────────────────────────
    def _clear_search(self):
        self.query_var.set("")

    def _on_search(self):
        if self.df is None:
            return
        cols = [c for c, v in self.col_vars.items() if v.get()]
        if not cols:
            self._show_table([], [])
            return

        q    = self.query_var.get().strip()
        case = self.case_var.get()
        df   = self.df
        mask = pd.Series(True, index=df.index)

        if q:
            mask &= df[cols].fillna("").apply(
                lambda col: col.astype(str).str.contains(q, case=case, na=False, regex=False)
            ).any(axis=1)

        for col, fvar in self.col_filters.items():
            fq = fvar.get().strip()
            if fq and col in df.columns:
                mask &= df[col].fillna("").astype(str).str.contains(
                    fq, case=case, na=False, regex=False)

        matched = df.loc[mask, cols]
        checked = self._checked_set()
        if self._sort_col == "✓" and self._keys is not None:
            is_checked = self._keys.loc[matched.index].isin(checked)
            matched = matched.loc[
                is_checked.sort_values(ascending=self._sort_asc, kind="stable").index]
        elif self._sort_col and self._sort_col in matched.columns:
            matched = matched.sort_values(
                self._sort_col, ascending=self._sort_asc, kind="stable", na_position="last",
                key=lambda s: s.astype(str).str.lower() if s.dtype == object else s)
        result = matched.head(1000)
        keys = self._keys.loc[result.index].tolist() if self._keys is not None else [""] * len(result)
        self._display_keys = keys
        self._display_cols = cols
        data = result.fillna("").astype(str).where(result.notna(), "").values.tolist()   # пустые даты — не «nan»
        rows = [[k in checked, *row] for k, row in zip(keys, data)]
        check_hdr = f"✓ {'▲' if self._sort_asc else '▼'}" if self._sort_col == "✓" else "✓"
        headers = [check_hdr] + [
            f"{c} {'▲' if self._sort_asc else '▼'}" if c == self._sort_col else c for c in cols]
        self._show_table(headers, rows, q, case)
        self.status_var.set(
            f"{int(mask.sum())} совпадений  |  {os.path.basename(self.current_file or '')}")

    def _show_table(self, columns, rows, query="", case=False):
        has_check = bool(columns) and columns[0].startswith("✓")
        self.sheet.headers(columns)
        self.sheet.set_sheet_data([list(r) for r in rows])
        try:
            self.sheet.set_all_column_widths()      # сначала — под содержимое
        except Exception:
            pass
        if has_check:
            try:
                self.sheet.checkbox("A", check_function=self._on_check, redraw=False)
                self.sheet.column_width(column=0, width=34, redraw=False)
            except Exception:
                pass
        # поверх — ширины, заданные руками/двойным кликом (21.10.2026: каждый поиск и сортировка
        # заново подгоняли все колонки, и ручная ширина тут же пропадала)
        saved = self._saved_widths()
        for i, name in enumerate(self._display_cols):
            w = saved.get(name)
            if w:
                try:
                    self.sheet.column_width(column=i + (1 if has_check else 0), width=int(w), redraw=False)
                except Exception:
                    pass
        self.sheet.redraw()
        self._highlight_matches(rows, query, case, skip_first=has_check)

    # ── ширина колонок: подгон под содержимое + ручная, запоминается по файлу ───
    def _saved_widths(self):
        return self.settings.get("col_widths", {}).get(self.current_file or "", {})

    def _on_col_resize(self, event=None):
        """Перетаскивание границы или двойной клик (подгон) — запомнить ширину ИЗМЕНЁННЫХ колонок
        по имени колонки, отдельно для каждого файла; остальные подгоняются под содержимое."""
        try:
            widths = self.sheet.get_column_widths()
            changed = dict((event or {}).get("resized", {}).get("columns", {}) or {})
        except Exception:
            return
        off = 1 if len(widths) == len(self._display_cols) + 1 else 0   # колонка ✓ не запоминается
        store = self.settings.setdefault("col_widths", {}).setdefault(self.current_file or "", {})
        for idx in (changed or range(len(widths))):
            i = int(idx) - off
            if 0 <= i < len(self._display_cols) and int(idx) < len(widths):
                store[self._display_cols[i]] = int(widths[int(idx)])
        save_settings(self.settings)

    def _on_col_header(self, event=None):
        # клик по заголовку колонки -> сортировка по ней (toggle asc/desc), на уровне pandas
        try:
            sel = self.sheet.get_currently_selected()
            if not sel or sel.type_ != "columns":
                return
            c = sel.column
            if c is None:
                return
            if c == 0:
                col = "✓"          # клик по заголовку ✓ -> сортировка по отметкам
            elif c - 1 < len(self._display_cols):
                col = self._display_cols[c - 1]
            else:
                return  # вне диапазона

            if self._sort_col == col:
                self._sort_asc = not self._sort_asc
            else:
                self._sort_col, self._sort_asc = col, True
            self._on_search()
        except Exception:
            pass

    def _on_cell_select(self, event=None):
        try:
            cell = self.sheet.get_currently_selected()
            if not cell:
                return
            if cell.type_ == "columns":   # клик по заголовку — это сортировка, не копирование
                return
            if cell[1] == 0:   # ✓ column — handled by checkbox, don't copy
                return
            val = self.sheet.get_cell_data(cell[0], cell[1])
            text = "" if val is None else str(val)
            self.fbar_var.set(text)
            if text:
                self.clipboard_clear()
                self.clipboard_append(text)
                self.status_var.set(f"Скопировано: {text[:80]}")
        except Exception:
            pass

    def _highlight_matches(self, rows, query, case, skip_first=False):
        try:
            self.sheet.dehighlight_all()
            if not query or len(rows) > 300:
                return
            q = query if case else query.lower()
            for r, row in enumerate(rows):
                for c, val in enumerate(row):
                    if skip_first and c == 0:
                        continue
                    v = val if case else val.lower()
                    if q in v:
                        self.sheet.highlight_cells(row=r, column=c,
                                                   bg="#f9e2af", fg="#1e1e2e")
        except Exception:
            pass

    # ── auto reload ───────────────────────────────────────────────────────────
    def _toggle_auto(self):
        if self.auto_var.get():
            if not self.current_file:
                self.auto_var.set(False)
                return
            self._auto_stop.clear()
            self._auto_thread = threading.Thread(target=self._auto_loop, daemon=True)
            self._auto_thread.start()
        else:
            self._auto_stop.set()

    def _auto_loop(self):
        while not self._auto_stop.is_set():
            time.sleep(5)
            if self.current_file and os.path.exists(self.current_file):
                mtime = os.path.getmtime(self.current_file)
                if mtime != self._last_mtime:
                    self.after(0, lambda: self._load_file(self.current_file))

    # ── export ────────────────────────────────────────────────────────────────
    def _export_csv(self):
        if self.df is None:
            return
        cols = [c for c, v in self.col_vars.items() if v.get()]
        q    = self.query_var.get().strip()
        case = self.case_var.get()
        df   = self.df
        mask = pd.Series(True, index=df.index)
        if q:
            mask &= df[cols].fillna("").apply(
                lambda col: col.astype(str).str.contains(q, case=case, na=False, regex=False)
            ).any(axis=1)
        for col, fvar in self.col_filters.items():
            fq = fvar.get().strip()
            if fq and col in df.columns:
                mask &= df[col].fillna("").astype(str).str.contains(
                    fq, case=case, na=False, regex=False)
        result = df.loc[mask, cols]
        if self._keys is not None:
            checked = self._checked_set()
            marks = self._keys.loc[result.index].isin(checked).map({True: "✓", False: ""})
            result = result.copy()
            result.insert(0, "✓", marks.values)
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All", "*.*")])
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                result.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                result.to_excel(path, index=False)
            messagebox.showinfo("Экспорт", f"Сохранено {len(result)} строк\n{path}")
        except Exception as exc:
            messagebox.showerror("Экспорт", f"Ошибка: {exc}")

    # ── справка, ярлык, обновление (стандарт GUI, 21.10.2026) ────────────────
    def _show_help(self):
        messagebox.showinfo(f"Excel Search {VERSION}", HELP_TEXT, parent=self)

    def _make_shortcut(self):
        """Один раз — ярлык «Excel Search» с иконкой программы на рабочем столе (только exe)."""
        if not getattr(sys, "frozen", False) or self.settings.get("shortcut_done"):
            return
        # рабочий стол — по Windows (бывает перенаправлен в OneDrive), не ~\Desktop
        exe = sys.executable
        ps = ("$l=Join-Path ([Environment]::GetFolderPath('Desktop')) '{0}.lnk';"
              "if(-not (Test-Path -LiteralPath $l)){{"
              "$s=(New-Object -ComObject WScript.Shell).CreateShortcut($l);"
              "$s.TargetPath='{1}';$s.WorkingDirectory='{2}';$s.IconLocation='{1},0';$s.Save()}}"
              ).format(*(p.replace("'", "''") for p in (APP_NAME, exe, APP_DIR)))
        try:
            subprocess.Popen(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
                             creationflags=0x08000000)
            self.settings["shortcut_done"] = True
            save_settings(self.settings)
        except OSError:
            pass

    def _update_fail_path(self):
        return os.path.join(APP_DIR, ".excel_search.update-failed.txt")

    def _check_update(self):
        """Тихо спросить GitHub о новой версии. Нет сети — молчим."""
        why = self._update_fail_path()
        if os.path.exists(why):
            try:
                msg = open(why, encoding="utf-8").read().strip()
                os.remove(why)
                messagebox.showwarning("Обновление", f"Прошлое обновление не встало: {msg}", parent=self)
            except OSError:
                pass

        def worker():
            import urllib.request
            try:
                req = urllib.request.Request(API_LATEST, headers={
                    "Accept": "application/vnd.github+json", "User-Agent": "excel-search"})
                data = json.load(urllib.request.urlopen(req, timeout=6))
            except Exception:
                return
            tag = str(data.get("tag_name", ""))
            if _parse_ver(tag) <= _parse_ver(VERSION):
                return
            asset = next((a for a in data.get("assets", []) if a.get("name") == UPDATE_ASSET), None)
            if asset and getattr(sys, "frozen", False):
                self.after(0, lambda: self._offer_update(tag, asset["browser_download_url"],
                                                         int(asset.get("size") or 0)))
        threading.Thread(target=worker, daemon=True, name="UpdateCheck").start()

    def _offer_update(self, tag, url, size):
        if not messagebox.askyesno("Обновление", f"Доступна версия {tag} (у вас {VERSION}).\n\n"
                                   "Скачать и обновить сейчас? Программа закроется и откроется заново.",
                                   parent=self):
            return
        threading.Thread(target=self._update_worker, args=(url, size), daemon=True).start()

    def _update_worker(self, url, size):
        """Сборка onedir (exe + _internal): архив скачивается во временную папку, после выхода
        программы скрытый PowerShell распаковывает его поверх папки программы и запускает exe.
        Настройки (excel_search_settings.json) в архиве нет — они не трогаются."""
        import shutil, tempfile, urllib.request
        zp = os.path.join(tempfile.gettempdir(), UPDATE_ASSET)
        try:
            with urllib.request.urlopen(url, timeout=180) as r, open(zp, "wb") as f:
                shutil.copyfileobj(r, f)
            if size and os.path.getsize(zp) != size:
                raise OSError(f"скачано {os.path.getsize(zp)} байт из {size}")
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror("Обновление", f"Не скачалось: {exc}", parent=self))
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
            f"if(-not $ok){{Set-Content -LiteralPath {q(self._update_fail_path())} "
            "-Value 'файлы программы были заняты' -Encoding UTF8}\r\n"
            f"Remove-Item -LiteralPath {q(zp)}\r\n"
            f"Start-Process -FilePath {q(sys.executable)}\r\n"
            "Remove-Item -LiteralPath $MyInvocation.MyCommand.Path\r\n")
        ps1 = os.path.join(tempfile.gettempdir(), "excel_search_update.ps1")
        with open(ps1, "w", encoding="utf-8-sig") as f:      # BOM: PS 5.1 иначе бьёт кириллицу в пути
            f.write(script)
        env = {k: v for k, v in os.environ.items() if not k.startswith(("_MEI", "_PYI"))}
        cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", ps1]
        try:
            subprocess.Popen(cmd, creationflags=0x08000000 | 0x01000000, env=env)
        except OSError:
            subprocess.Popen(cmd, creationflags=0x08000000, env=env)
        self.after(300, self.destroy)

    # ── close ─────────────────────────────────────────────────────────────────
    def destroy(self):
        self._auto_stop.set()
        save_settings(self.settings)
        super().destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
