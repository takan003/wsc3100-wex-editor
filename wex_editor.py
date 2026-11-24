"""廣播時程編輯器（WSC3100 EDIT 相容介面）。

用法：
  python wex_editor.py                 # 啟動後手動匯入 .WEX
  python wex_editor.py 檔案.WEX        # 指定要開啟的檔案
"""
import json
import os
import re
import sys
import traceback
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

from tksheet import Sheet

from wex_io import (
    WexProject, NO_DATE, WEEK_LABELS, MODE_LABELS,
    build_update_row, build_insert_row, build_delete_row,
    row_duration_changes, export_wex, duration_to_clock_text, sql_quote,
)

GRID_HEADERS = ["序號", "名稱註記"] + WEEK_LABELS + ["啟動日期", "播放模式", "啟動時間", "終止時間", "狀態"]
COL_NO, COL_NAME = 0, 1
COL_WEEK0 = 2
COL_DATE, COL_MODE, COL_START, COL_END, COL_STATUS = 9, 10, 11, 12, 13

def lock_sheet_format(sheet):
    """鎖定官方欄位格式：禁止增刪欄位、欄位重排、欄名編輯與任何會跨欄搬移值的排序。"""
    try:
        sheet.disable_bindings(
            "rc_insert_column", "insert_columns", "add_columns",
            "delete_columns", "del_columns",
            "sort_cells", "sort_row", "sort_column", "sort_columns",
        )
    except Exception:
        mt = getattr(sheet, "MT", None)
        if mt is not None:
            for attr in ("rc_insert_column_enabled", "rc_delete_column_enabled",
                         "rc_sort_column_enabled", "rc_sort_columns_enabled",
                         "rc_sort_row_enabled", "rc_sort_cells_enabled"):
                try:
                    setattr(mt, attr, False)
                except Exception:
                    pass
    mt = getattr(sheet, "MT", None)
    if mt is not None:
        try:
            mt.header_edit_cell_enabled = lambda: False
            mt.CH.edit_cell_enabled = False
        except Exception:
            pass
    for owner in (getattr(sheet, "ops", None), mt):
        if owner is None:
            continue
        try:
            owner.column_drag_and_drop_perform = False
        except Exception:
            pass


SHEET_MENU_ZH = {
    "edit_cell_label": "編輯儲存格",
    "edit_header_label": "編輯標題",
    "edit_index_label": "編輯列號",
    "cut_label": "剪下",
    "copy_label": "複製",
    "copy_plain_label": "複製純文字",
    "paste_label": "貼上",
    "delete_label": "刪除",
    "clear_contents_label": "清除內容",
    "delete_rows_label": "刪除列",
    "delete_columns_label": "刪除欄",
    "insert_rows_above_label": "上方插入列",
    "insert_rows_below_label": "下方插入列",
    "insert_row_label": "插入列",
    "insert_column_label": "插入欄",
    "insert_columns_left_label": "左方插入欄",
    "insert_columns_right_label": "右方插入欄",
    "sort_cells_label": "遞增排序",
    "sort_cells_reverse_label": "遞減排序",
    "sort_cells_x_label": "列向遞增排序",
    "sort_cells_x_reverse_label": "列向遞減排序",
    "sort_row_label": "值由小到大排序",
    "sort_row_reverse_label": "值由大到小排序",
    "sort_column_label": "值由小到大排序",
    "sort_column_reverse_label": "值由大到小排序",
    "sort_rows_label": "列由小到大排序",
    "sort_rows_reverse_label": "列由大到小排序",
    "sort_columns_label": "欄位由小到大排序",
    "sort_columns_reverse_label": "欄位由大到小排序",
    "select_all_label": "全選",
    "undo_label": "復原",
    "redo_label": "重做",
}


def localize_sheet_menu(sheet):
    for key, zh in SHEET_MENU_ZH.items():
        try:
            sheet.ops[key] = zh
        except Exception:
            try:
                setattr(sheet.ops, key, zh)
            except Exception:
                pass


def configure_group_tab_style():
    style = ttk.Style()
    if "Schedule.tab" not in style.element_names():
        style.element_create("Schedule.tab", "from", "clam", "Notebook.tab")
    style.layout("Schedule.TNotebook.Tab", [
        ("Schedule.tab", {"sticky": "nswe", "children": [
            ("Notebook.padding", {"side": "top", "sticky": "nswe", "children": [
                ("Notebook.focus", {"side": "top", "sticky": "nswe", "children": [
                    ("Notebook.label", {"side": "top", "sticky": ""}),
                ]}),
            ]}),
        ]}),
    ])
    style.configure("Schedule.TNotebook", background="#e8eff7", borderwidth=0)
    style.configure(
        "Schedule.TNotebook.Tab", font=("Microsoft JhengHei UI", 10, "bold"),
        padding=(12, 8), background="#edf2f8", foreground="#34465a",
        bordercolor="#afc1d3", lightcolor="#ffffff", darkcolor="#afc1d3")
    style.map(
        "Schedule.TNotebook.Tab",
        background=[("selected", "#205e9c"), ("active", "#d4e5f6")],
        foreground=[("selected", "#ffffff")],
        bordercolor=[("selected", "#205e9c")],
        lightcolor=[("selected", "#205e9c")],
        darkcolor=[("selected", "#205e9c")])


PERIOD_SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "課表作息.json")
PREFS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "喜好設定.json")
DEFAULT_PERIODS = [
    [1, "第1節", "08:00:00"],
    [2, "第2節", "09:20:00"],
    [3, "第3節", "10:20:00"],
    [4, "第4節", "11:20:00"],
    [5, "午休", "12:40:00"],
    [6, "第5節", "13:20:00"],
    [7, "第6節", "14:20:00"],
    [8, "第7節", "15:20:00"],
    [9, "第8節", "16:20:00"],
]
PERIOD_HEADERS = ["序號", "節次名稱註記", "啟動時間"]


def fmt_date(d):
    if not d:
        return ""
    return f"{d[0]:04d}-{d[1]:02d}-{d[2]:02d}"


def parse_date(s):
    s = (s or "").strip()
    if not s:
        return None
    parts = s.replace("/", "-").split("-")
    if len(parts) != 3:
        raise ValueError("日期格式應為 YYYY-MM-DD")
    y, m, d = (int(p) for p in parts)
    return y, m, d


def fmt_start(hm):
    return f"{hm[0]:02d}:{hm[1]:02d}"


def parse_start(s):
    s = (s or "").strip()
    if not s:
        raise ValueError("啟動時間不可為空")
    parts = s.replace("：", ":").split(":")
    if len(parts) == 2:
        h, m = int(parts[0]), int(parts[1])
    elif len(parts) == 3:
        h, m = int(parts[0]), int(parts[1])
    else:
        raise ValueError("時間格式應為 HH:MM")
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ValueError("時間超出範圍")
    return h, m


def fmt_status(on):
    return "已排入時程" if on else "未排入"


def parse_status(s):
    return 1 if s != "未排入" else 0


class ScheduleImportDialog(simpledialog.Dialog):
    def __init__(self, parent, groups, current_tab, new_id, max_name_units):
        self.groups = groups
        self.current_tab = current_tab
        self.new_id = new_id
        self.max_name_units = max_name_units
        super().__init__(parent, title="選擇匯入目標")

    def body(self, master):
        self.mode = tk.StringVar(value="replace" if self.groups else "new")
        self.new_name = tk.StringVar()
        ttk.Label(master, text="選擇這次監考排程要寫入的鐘聲組別：").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        self.replace_radio = ttk.Radiobutton(
            master, text="覆蓋既有工作表", variable=self.mode,
            value="replace", command=self._update_controls)
        self.replace_radio.grid(row=1, column=0, sticky="w", padx=(0, 8))
        if not self.groups:
            self.replace_radio.state(["disabled"])
        self.group_combo = ttk.Combobox(
            master, state="readonly", width=36,
            values=[f"{sid}. {name}" for sid, name in self.groups])
        self.group_combo.grid(row=1, column=1, sticky="ew", pady=4)
        if self.groups:
            ids = [sid for sid, _ in self.groups]
            self.group_combo.current(ids.index(self.current_tab) if self.current_tab in ids else 0)
        ttk.Label(master, text="覆蓋會取代該組全部原有時程，其他組別不受影響。",
                  foreground="#8a4815").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Radiobutton(master, text="新增工作表", variable=self.mode,
                        value="new", command=self._update_controls).grid(
            row=3, column=0, sticky="w", padx=(0, 8))
        self.name_entry = ttk.Entry(master, textvariable=self.new_name, width=38)
        self.name_entry.grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Label(master, text="輸入新工作表名稱；新增後需另行匯出 WEX。").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(0, 6))
        if self.new_id > 8:
            ttk.Label(master, foreground="#a23620", wraplength=480,
                      text=f"新工作表將使用第 {self.new_id} 組；"
                           "官方軟體是否支援第 9 組以上尚未確認。").grid(
                row=5, column=0, columnspan=2, sticky="w")
        master.columnconfigure(1, weight=1)
        self._update_controls()
        return self.group_combo if self.groups else self.name_entry

    def _update_controls(self):
        replacing = self.mode.get() == "replace"
        self.group_combo.configure(state="readonly" if replacing else "disabled")
        self.name_entry.configure(state="disabled" if replacing else "normal")

    def buttonbox(self):
        box = ttk.Frame(self)
        ttk.Button(box, text="確定", command=self.ok).pack(side="left", padx=5, pady=5)
        ttk.Button(box, text="取消", command=self.cancel).pack(side="left", padx=5, pady=5)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        box.pack()

    def validate(self):
        if self.mode.get() == "replace":
            index = self.group_combo.current()
            if not 0 <= index < len(self.groups):
                messagebox.showerror("請選擇工作表", "請選擇要覆蓋的鐘聲組別。", parent=self)
                return False
            sid, name = self.groups[index]
            return messagebox.askyesno(
                "確認覆蓋",
                f"確定取代「{sid}. {name}」的全部時程？\n\n"
                "包括尚未匯出的修改及已加入的學校作息；其他組別不受影響。",
                parent=self)
        name = self.new_name.get().strip()
        if not name:
            messagebox.showerror("缺少名稱", "請輸入新的工作表名稱。", parent=self)
            return False
        if len(name.encode("utf-16-le")) // 2 > self.max_name_units:
            messagebox.showerror("名稱過長",
                                 f"工作表名稱最多 {self.max_name_units} 個字元（部分特殊符號佔兩個）。",
                                 parent=self)
            return False
        if any(name.casefold() == existing.strip().casefold() for _, existing in self.groups):
            messagebox.showerror("名稱重複", "已有相同名稱的鐘聲組別，請使用其他名稱。", parent=self)
            return False
        if self.new_id > 8:
            return messagebox.askyesno(
                "官方相容性未確認",
                f"新工作表將使用第 {self.new_id} 組。\n"
                "官方控制軟體是否支援第 9 組以上尚未確認，可能無法顯示或匯入。\n\n"
                "仍要新增嗎？", parent=self)
        return True

    def apply(self):
        if self.mode.get() == "replace":
            sid, name = self.groups[self.group_combo.current()]
            self.result = ("replace", sid, name)
        else:
            self.result = ("new", self.new_id, self.new_name.get().strip())


class Workspace(ttk.Frame):
    """單一工作區：工具列、鐘聲組別頁籤與時程表格。"""

    def __init__(self, master, shell):
        super().__init__(master)
        self.shell = shell
        self.proj = None
        self.path = None
        self.tab_rows = {}
        self.tab_new_seq = {}
        self.tab_names = {}
        self.tab_deleted = {}
        self.tab_backup = {}
        self.orig_tab_names = {}
        self.visible = []
        self.current_tab = None
        self.dirty = set()
        self.detail_win = None
        self._detail_row = None
        self._period_win = None

        self._build_toolbar()
        ttk.Label(self, text="鐘聲時程組別  ·  點選切換，拖曳下方捲軸查看更多",
                  font=("Microsoft JhengHei UI", 10, "bold"),
                  foreground="#34465a").pack(anchor="w", padx=8, pady=(4, 3))
        configure_group_tab_style()
        self.tab_canvas = tk.Canvas(self, height=40, background="#e8eff7", highlightthickness=0)
        self.tab_canvas.pack(fill="x", padx=6)
        self.tab_scrollbar = ttk.Scrollbar(
            self, orient="horizontal", command=self.tab_canvas.xview)
        self.tab_scrollbar.pack(fill="x", padx=6)
        self.tab_canvas.configure(xscrollcommand=self.tab_scrollbar.set)
        self.group_tabs = ttk.Notebook(self.tab_canvas, style="Schedule.TNotebook")
        self._tabs_window = self.tab_canvas.create_window(
            0, 0, anchor="nw", window=self.group_tabs)
        self.tab_canvas.bind("<Configure>", self._resize_group_tabs)
        self._tab_ids = {}
        self.group_tabs.bind("<<NotebookTabChanged>>", self._on_group_selected)

        self.sheet = Sheet(self, width=1040, height=560,
                           header_font=("Microsoft JhengHei UI", 10, "bold"),
                           table_font=("Microsoft JhengHei UI", 10, "normal"))
        self.sheet.pack(fill="both", expand=True, padx=6, pady=(0, 2))
        localize_sheet_menu(self.sheet)
        self.sheet.headers(GRID_HEADERS)
        self._apply_column_widths()
        self.sheet.enable_bindings()
        lock_sheet_format(self.sheet)
        self.sheet.extra_bindings("end_edit_cell", self._on_cell_edited)
        self.sheet.extra_bindings("double_click", self._on_double_click)

    @property
    def root(self):
        return self.shell.root

    @property
    def status(self):
        return self.shell.status

    def _build_toolbar(self):
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=6, pady=4)
        ttk.Button(bar, text="加入學校作息", command=self.embed_school_periods).pack(side="left", padx=(0, 12))
        ttk.Button(bar, text="匯入監考表排程", command=self.import_schedule_excel).pack(side="left", padx=(0, 4))
        ttk.Button(bar, text="條目整併", command=self.merge_entries).pack(side="left", padx=(0, 4))
        ttk.Button(bar, text="匯入舊WEX檔", command=self.import_old_wex).pack(side="left", padx=(0, 4))
        ttk.Button(bar, text="匯出WEX", command=self.export_dialog).pack(side="left", padx=4)

    def _period_rows_from_settings(self):
        """把『設定→課表作息』的節次轉成表格列：每節開始響上課鐘（啟動 + 上課.wav 秒數）。"""
        periods = self._load_periods()
        audio = "上課.wav"
        auds = self._audio_durations()

        rows = []
        for i, p in enumerate(periods, start=1):
            seq = int(p[0]) if str(p[0]).isdigit() else i
            name = str(p[1]).strip()
            if not name:
                continue
            start = str(p[2]).strip() or "08:00:00"
            parts = []
            for x in start.replace("：", ":").split(":"):
                try:
                    parts.append(int(x))
                except ValueError:
                    parts.append(0)
            while len(parts) < 2:
                parts.append(0)
            d = {
                "MUSIC": -seq, "Mode": 1, "Music_on": 1, "Allcall": True, "Vol": 80,
                "Week_day": False, "Channel": False,
                "Group1": 0, "Group2": 0, "Group3": 0,
                "Dir0": audio, "Time": 0, "Total_time": 0,
                "End_hour": "00", "End_min": "00", "End_sec": "00",
                "Music_no": seq, "Music_name": name,
                "Clock_year": "2015", "Clock_month": "03", "Clock_day": "14",
                "Clock_start_hour": f"{parts[0]:02d}",
                "Clock_start_min": f"{parts[1]:02d}",
            }
            for w in range(1, 8):
                d[f"Week{w}"] = w <= 5  # 週一~週五
            d.update(row_duration_changes(d, auds))
            rows.append(d)
        return rows

    def embed_school_periods(self):
        """加入學校作息（按鈕）：把課表作息節次追加到目前組（已存在的節次不重複加入）。"""
        new_rows = self._period_rows_from_settings()
        if not new_rows:
            messagebox.showinfo("加入學校作息", "「設定 → 課表作息」沒有可用的節次資料。")
            return
        self._commit_sheet_to_model()
        created = self.current_tab is None
        if created:
            self._create_group_from_rows(new_rows)
            self._reload_sheet()
            added = len(new_rows)
        else:
            rows = self.tab_rows.setdefault(self.current_tab, [])
            have = set()
            for r in rows:
                try:
                    h = int(r.get("Clock_start_hour") or 0)
                    m = int(r.get("Clock_start_min") or 0)
                except Exception:
                    h, m = -1, -1
                have.add((str(r.get("Music_name") or "").strip(), h, m))
            next_no = max((int(r.get("Music_no") or 0) for r in rows), default=0)
            added = 0
            for d in new_rows:
                key = (str(d.get("Music_name") or "").strip(),
                       int(d.get("Clock_start_hour") or 0),
                       int(d.get("Clock_start_min") or 0))
                if key in have:
                    continue
                next_no += 1
                d["Music_no"] = next_no
                rows.append(d)
                have.add(key)
                added += 1
            self.dirty.add(self.current_tab)
            self._reload_sheet()
        total = len(self.tab_rows.get(self.current_tab, []))
        if added:
            self.status.config(text=f"已加入學校作息（追加 {added} 節，共 {total} 列）")
        else:
            self.status.config(text=f"加入學校作息：節次皆已存在（共 {total} 列），未重複加入")

    def _switch_group(self, sid):
        if sid == self.current_tab:
            return
        self._commit_sheet_to_model()
        self.current_tab = sid
        self._reload_sheet()
        self.status.config(text=f"已切換到「{sid}.{self.tab_names.get(sid, '')}」")

    def import_schedule_excel(self):
        if self.proj is None:
            messagebox.showinfo("請先匯入 WEX",
                                "請先使用「匯入舊WEX檔」選擇包含音檔與鐘聲組別的完整 WEX。")
            return
        path = filedialog.askopenfilename(
            title="匯入監考表排程",
            filetypes=[("Excel 檔 (*.xlsx;*.xls)", "*.xlsx *.xls"), ("所有檔案", "*.*")],
        )
        if not path:
            return
        try:
            source = "範本"
            exam_rows = self._load_invigilation_rows(path)
            if exam_rows is not None:
                new_rows = exam_rows
                source = "監考表"
            else:
                raw = self._read_excel_rows(path)
                if raw is None:
                    return
                new_rows = self._parse_excel_rows(raw)
                if new_rows is None:
                    return
            if not new_rows:
                messagebox.showinfo("匯入監考表排程", "試算表中沒有可匯入的時程列。")
                return
            missing = sorted({str(d.get("Dir0") or "") for d in new_rows} - set(self.proj.audio_files))
            if missing:
                raise ValueError("所選 WEX 缺少排程需要的音檔：\n" + "\n".join(missing))

            groups = [(sid, self.tab_names.get(sid, "")) for sid in self.visible]
            new_id = next(i for i in range(1, max(self.visible, default=0) + 2)
                          if i not in self.visible)
            music_table = self.proj.db.read_table(self.proj._catalog["MUSIC"])
            self.proj.db.read_columns(music_table)
            name_column = next(c for c in music_table.columns if c.name == "Music_name")
            dialog = ScheduleImportDialog(
                self.root, groups, self.current_tab, new_id, name_column.col_size // 2)
            if dialog.result is None:
                return
            mode, sid, name = dialog.result
            self._commit_sheet_to_model()
            old_n = len(self.tab_rows.get(sid, []))
            for i, d in enumerate(new_rows, 1):
                d["MUSIC"] = -i
                d["Music_no"] = i
            self.tab_rows[sid] = new_rows
            self.tab_names[sid] = name
            self.tab_deleted[sid] = False
            self.tab_backup.pop(sid, None)
            self.tab_new_seq[sid] = -len(new_rows)
            if sid not in self.visible:
                self.visible.append(sid)
                self.visible.sort()
            self.current_tab = sid
            self.dirty.add(sid)
            self._refresh_group_tabs()
            self._reload_sheet()
            self.after_idle(self._show_selected_group_tab)
            self.status.config(
                text=f"已匯入監考表排程（{source}，{'覆蓋' if mode == 'replace' else '新增'}）："
                     f"{sid}.{name}（{old_n} → {len(new_rows)} 列）；請確認後匯出 WEX")
            if self._pref_auto_merge():
                self.merge_entries(auto=True)
        except ValueError as e:
            messagebox.showerror("匯入失敗", str(e))
        except Exception:
            tb = traceback.format_exc()
            messagebox.showerror("匯入失敗", tb[-1500:])

    def merge_entries(self, auto=False):
        """條目整併：同節次/同時間/同音檔的跨日列併成一列
        （後面併入前面：星期勾選取聯集、名稱日期合併如「10/28, 29-1-起始」、移除多餘列並重排序號）。"""
        if self.current_tab is None or not self.tab_rows.get(self.current_tab):
            if not auto:
                messagebox.showinfo("條目整併", "目前時程組沒有條目。")
            return
        self._commit_sheet_to_model()
        rows = self.tab_rows[self.current_tab]
        name_re = re.compile(r"^(\d{1,2}/\d{1,2}(?:,\s*(?:\d{1,2}/)?\d{1,2})*)(-.+)$")

        def parse_name(name):
            m = name_re.match(name)
            if not m:
                return None
            toks = m.group(1).split(", ")
            dates = []
            first_mo = None
            for j, t in enumerate(toks):
                if "/" in t:
                    mo, dd = t.split("/")
                    mo, dd = int(mo), int(dd)
                    if j == 0:
                        first_mo = mo
                else:
                    mo, dd = first_mo, int(t)
                dates.append((mo, dd))
            return dates, m.group(2)

        groups = {}
        for idx, d in enumerate(rows):
            parsed = parse_name(str(d.get("Music_name") or ""))
            if parsed is None:
                continue
            _, rest = parsed
            date = tuple(int(d.get(col) or 0) for col in ("Clock_year", "Clock_month", "Clock_day"))
            if date == NO_DATE or date[0] == 0:
                date = None
            key = (rest,
                   int(d.get("Clock_start_hour") or 0), int(d.get("Clock_start_min") or 0),
                   int(d.get("Total_time") or 0), str(d.get("Dir0") or ""),
                   int(d.get("Group1") or 0), int(d.get("Group2") or 0), int(d.get("Group3") or 0),
                   bool(d.get("Allcall")), int(d.get("Vol") or 0), date)
            groups.setdefault(key, []).append((idx, parsed[0], d))

        drop = set()
        merged_n = removed = 0
        for key, items in groups.items():
            if len(items) < 2:
                continue
            merged_n += 1
            _, keeper_dates, keeper = items[0]
            rest = key[0]
            for idx, dates, d in items[1:]:
                mo, dd = dates[0]
                if (mo, dd) not in keeper_dates:
                    keeper_dates.append((mo, dd))
                for w in range(1, 8):
                    if d.get(f"Week{w}"):
                        keeper[f"Week{w}"] = True
                drop.add(idx)
                removed += 1
            if not any(keeper.get(f"Week{w}") for w in range(1, 8)):
                try:
                    y = int(keeper.get("Clock_year") or 0)
                    if y and y != 2015:
                        wk = datetime(y, int(keeper.get("Clock_month") or 1),
                                      int(keeper.get("Clock_day") or 1)).weekday() + 1
                        keeper[f"Week{wk}"] = True
                except Exception:
                    pass
            if len(keeper_dates) > 1:
                first_mo = keeper_dates[0][0]
                parts = [f"{keeper_dates[0][0]}/{keeper_dates[0][1]}"]
                for mo, dd in keeper_dates[1:]:
                    parts.append(str(dd) if mo == first_mo else f"{mo}/{dd}")
                keeper["Music_name"] = ", ".join(parts) + rest

        if not merged_n:
            if not auto:
                self.status.config(text="條目整併：沒有可合併的條目（需同名稱格式、同啟動時間、同音檔）")
            return
        before = len(rows)
        rows[:] = [d for i, d in enumerate(rows) if i not in drop]
        for i, d in enumerate(rows, 1):
            d["Music_no"] = i
        self.dirty.add(self.current_tab)
        self._reload_sheet()
        self.status.config(
            text=f"條目整併：合併 {merged_n} 組、移除 {removed} 列（{before} → {len(rows)} 列）")

    FALLBACK_AUDIO_DUR = {
        "上課.wav": 53.0, "下課.wav": 55.0,
        "01-預備鈴.mp3": 20.0, "02-考試鈴.mp3": 19.0, "03-結束鈴.mp3": 17.0,
        "費玉青-晚安曲.mp3": 215.0,
        "Alarm1.mp3": 38.0, "Alarm2.mp3": 25.5, "Alarm4.mp3": 52.5,
    }

    def _audio_durations(self):
        """音檔時長：以基底實際量測為準，未載入基底時用內建近似值。"""
        d = dict(self.FALLBACK_AUDIO_DUR)
        if self.proj is not None:
            d.update(self.proj.audio_durations)
        return d
    WEEK_CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "日": 7, "天": 7}
    SESS_CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}

    def _load_invigilation_rows(self, path):
        """檔內有『監考表』工作表 → 解析成時程列；沒有 → 回傳 None（改走範本格式）。"""
        lower = path.lower()
        if lower.endswith(".xls"):
            import xlrd
            book = xlrd.open_workbook(path)
            names = book.sheet_names()
            if "監考表" not in names:
                return None
            sh = book.sheet_by_name("監考表")
            grid = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
        else:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            try:
                if "監考表" not in wb.sheetnames:
                    return None
                ws = wb["監考表"]
                grid = [list(row) for row in ws.iter_rows(values_only=True)]
            finally:
                wb.close()
        return self._parse_invigilation_grid(grid, path)

    @staticmethod
    def _invig_to_hm(v):
        if v is None:
            return None
        if hasattr(v, "hour") and hasattr(v, "minute"):
            return int(v.hour), int(v.minute)
        s = str(v).strip()
        m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", s)
        if not m:
            return None
        return int(m.group(1)), int(m.group(2))

    @staticmethod
    def _invig_ffill(row_vals, ncols):
        out, last = [], None
        for c in range(ncols):
            v = row_vals[c] if c < len(row_vals) else None
            if v is not None and str(v).strip() != "":
                last = v
            out.append(last)
        return out

    def _guess_exam_year(self, path, grid):
        base = os.path.basename(path)
        m = re.match(r"^(\d{2,3})[-_]", base)
        if m:
            roc = int(m.group(1))
            if 100 <= roc <= 200:
                return roc + 1911
        for row in grid[:3]:
            for cell in row:
                if cell is None:
                    continue
                m = re.search(r"(?:民國)?(\d{2,3})\s*[-–]", str(cell))
                if m:
                    roc = int(m.group(1))
                    if 100 <= roc <= 200:
                        return roc + 1911
        return datetime.now().year

    def _parse_invigilation_grid(self, grid, path):
        if not grid:
            raise ValueError("監考表工作表是空的")

        def s(v):
            return str(v).strip() if v is not None else ""

        rows_by_label = {}
        for r, row in enumerate(grid):
            a = s(row[0] if row else "").rstrip("：: ")
            if a:
                rows_by_label.setdefault(a, r)
        if "日期" not in rows_by_label:
            raise ValueError("監考表找不到『日期』列（A欄標籤）。")
        r_date = rows_by_label["日期"]
        r_week = rows_by_label.get("星期")
        r_sess = rows_by_label.get("考試節次")

        # 時間列：有就抓、沒有就跳過（定期考/模擬考標籤略有差異）
        time_rows = []
        seen_types = set()
        for label, r in rows_by_label.items():
            typ = None
            if label in ("預備鐘聲", "預備鈴聲") or ("預備" in label and ("鐘" in label or "鈴" in label)):
                typ = "預備"
            elif label == "起始時間":
                typ = "起始"
            else:
                m = re.fullmatch(r"結束時間(\d*)", label)
                if m:
                    d = m.group(1)
                    typ = "結束" if d in ("", "1") else f"結束{d}"
            if typ is not None and typ not in seen_types:
                seen_types.add(typ)
                time_rows.append((typ, r))
        time_rows.sort(key=lambda kv: kv[1])
        if not time_rows:
            raise ValueError(
                "監考表找不到任何時間列。\n"
                "支援標籤（有就抓、沒有就跳過）：\n"
                "預備鐘聲／預備鈴聲、起始時間、結束時間1~N")

        ncols = min(max(len(row) for row in grid), 120)
        date_filled = self._invig_ffill(grid[r_date], ncols)
        week_filled = self._invig_ffill(grid[r_week], ncols) if r_week is not None else None

        year = self._guess_exam_year(path, grid)
        write_date = self._pref_import_write_date()

        hay = os.path.basename(path) + " " + " ".join(s(c) for c in grid[0])
        is_mock = "模擬考" in hay
        if is_mock:
            grp = (3, 3, 9)
            allcall = False
        else:
            grp = (1, 1, 1)
            allcall = True

        auds = self._audio_durations()

        type_rows = list(time_rows)

        def parse_date(v):
            m = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日", s(v))
            if not m:
                return None
            return int(m.group(1)), int(m.group(2))

        def parse_week(v, month, day):
            m = re.search(r"星期([一二三四五六日天])", s(v))
            if m:
                return self.WEEK_CN[m.group(1)]
            try:
                return datetime(year, month, day).weekday() + 1
            except Exception:
                return None

        session_cols = []
        if r_sess is not None:
            for c in range(1, ncols):
                sv = grid[r_sess][c] if c < len(grid[r_sess]) else None
                if sv is not None and str(sv).strip() != "":
                    session_cols.append(c)
        else:
            for c in range(1, ncols):
                if parse_date(date_filled[c] if c < len(date_filled) else None) is None:
                    continue
                if any(self._invig_to_hm(grid[r][c] if c < len(grid[r]) else None)
                       for _, r in time_rows):
                    session_cols.append(c)
        if not session_cols:
            raise ValueError("監考表找不到可用的時間欄（有時間列但欄位裡沒有時間值）")

        rows = []
        seq = 0
        last_month = last_day = None
        for c in session_cols:
            md = parse_date(date_filled[c]) if c < len(date_filled) else None
            if md is None:
                continue
            month, day = md
            if (month, day) != (last_month, last_day):
                last_month, last_day = month, day
                day_no = None
            week_n = parse_week(
                (week_filled[c] if week_filled else None) or date_filled[c], month, day)

            if r_sess is not None:
                sess_raw = s(grid[r_sess][c] if c < len(grid[r_sess]) else None)
                if sess_raw in self.SESS_CN:
                    sess_no = self.SESS_CN[sess_raw]
                elif sess_raw.isdigit():
                    sess_no = int(sess_raw)
                else:
                    sess_no = (day_no or 0) + 1
            else:
                sess_no = (day_no or 0) + 1
            day_no = sess_no if day_no is None else max(day_no, sess_no)

            for typ, rws in type_rows:
                hm = self._invig_to_hm(grid[rws][c] if c < len(grid[rws]) else None)
                if hm is None:
                    continue
                total_min = hm[0] * 60 + hm[1]
                sh, sm = divmod(total_min, 60)
                sh %= 24
                seq += 1
                if typ == "預備":
                    audio = "01-預備鈴.mp3"
                elif typ == "起始":
                    audio = "02-考試鈴.mp3"
                else:
                    audio = "03-結束鈴.mp3"
                d = {
                    "MUSIC": -seq, "Mode": 1, "Music_on": 1, "Allcall": allcall, "Vol": 80,
                    "Week_day": False, "Channel": False,
                    "Group1": grp[0], "Group2": grp[1], "Group3": grp[2],
                    "Dir0": audio, "Time": 0, "Total_time": 0,
                    "End_hour": "00", "End_min": "00", "End_sec": "00",
                    "Music_no": seq,
                    "Music_name": f"{month}/{day}-{sess_no}-{typ}",
                    "Clock_year": str(year) if write_date else "2015",
                    "Clock_month": f"{month:02d}" if write_date else "03",
                    "Clock_day": f"{day:02d}" if write_date else "14",
                    "Clock_start_hour": f"{sh:02d}",
                    "Clock_start_min": f"{sm:02d}",
                }
                for w in range(1, 8):
                    d[f"Week{w}"] = (week_n == w)
                d.update(row_duration_changes(d, auds))
                rows.append(d)
        return rows


    def import_old_wex(self):
        path = filedialog.askopenfilename(
            title="匯入舊WEX檔",
            filetypes=[("廣播時程檔 (*.WEX)", "*.WEX"), ("所有檔案", "*.*")],
        )
        if not path:
            return
        self.shell.open_path(path)

    def clear_workspace(self):
        self.proj = None
        self.path = None
        self.tab_rows = {}
        self.tab_new_seq = {}
        self.tab_names = {}
        self.tab_deleted = {}
        self.tab_backup = {}
        self.orig_tab_names = {}
        self.visible = []
        self.current_tab = None
        self.dirty = set()
        self._detail_row = None
        if self.detail_win is not None:
            try:
                if self.detail_win.winfo_exists():
                    self.detail_win.destroy()
            except Exception:
                pass
            self.detail_win = None
        self._refresh_group_tabs()
        self._reload_sheet()
        self.status.config(
            text="工作區已清空 — 請匯入 WEX 檔，再編輯或匯入排程")

    def open_detail_dialog(self, r=None):
        if r is None:
            try:
                r = self.sheet.get_currently_selected()[0].row
            except Exception:
                messagebox.showinfo("提示", "請先選取一列（或雙擊列）")
                return
        rows = self.tab_rows.get(self.current_tab, [])
        if r is None or r >= len(rows):
            return
        self._commit_sheet_to_model()
        if self.detail_win is not None and self.detail_win.winfo_exists():
            self.detail_win.destroy()
        win = tk.Toplevel(self.root)
        win.title(f"列設定 - {rows[r].get('Music_name') or ''}")
        win.transient(self.root)
        win.resizable(False, False)
        self.detail_win = win
        self._detail_row = r
        d = rows[r]
        proj = self.proj
        audio_files = list(proj.audio_files) if proj is not None else list(self.FALLBACK_AUDIO_DUR)
        group_names = proj.group_names if proj is not None else {}

        ttk.Label(win, text="音檔:").grid(row=0, column=0, sticky="e", padx=6, pady=4)
        self.var_dir = tk.StringVar()
        self.cmb_dir = ttk.Combobox(win, textvariable=self.var_dir, width=26, state="readonly")
        self.cmb_dir["values"] = [""] + audio_files
        cur_dir = str(d.get("Dir0") or "")
        self.var_dir.set(cur_dir if cur_dir in audio_files else "")
        self.cmb_dir.grid(row=0, column=1, columnspan=2, sticky="w", padx=6)
        dur = self._audio_durations().get(cur_dir)
        self.var_dir_dur = ttk.Label(win, text=f"（{duration_to_clock_text(dur)}）" if dur else "")
        self.var_dir_dur.grid(row=0, column=3, sticky="w", padx=4)

        gitems = [""]
        for gid in sorted(group_names):
            label = group_names[gid] or f"群組{gid}"
            gitems.append(f"{gid}:{label}")
        for i, label in enumerate(("群組1:", "群組2:", "群組3:"), start=1):
            ttk.Label(win, text=label).grid(row=1, column=i * 2 - 2, sticky="e", padx=4)
            var = tk.StringVar()
            cmb = ttk.Combobox(win, textvariable=var, width=14, state="readonly")
            cmb["values"] = gitems
            gid = int(d.get(f"Group{i}") or 0)
            var.set(f"{gid}:{group_names[gid]}" if gid and gid in group_names else "")
            cmb.grid(row=1, column=i * 2 - 1, sticky="w", padx=4)
            setattr(self, f"var_group{i}", var)
            setattr(self, f"cmb_group{i}", cmb)

        self.var_allcall = tk.BooleanVar(value=bool(d.get("Allcall")) if d.get("Allcall") is not None else True)
        ttk.Checkbutton(win, text="全體廣播", variable=self.var_allcall, command=self._detail_changed).grid(
            row=2, column=0, columnspan=2, sticky="w", padx=6, pady=3)
        ttk.Label(win, text="音量:").grid(row=2, column=2, sticky="e", padx=4)
        v = d.get("Vol")
        self.var_vol = tk.StringVar(value=str(int(v)) if v is not None else "80")
        self.spin_vol = ttk.Spinbox(win, from_=0, to=100, width=5, textvariable=self.var_vol,
                                    command=self._detail_changed)
        self.spin_vol.grid(row=2, column=3, sticky="w", padx=4)
        self.spin_vol.bind("<FocusOut>", lambda e: self._detail_changed())
        self.spin_vol.bind("<Return>", lambda e: self._detail_changed())

        gid0 = int(d.get("Group1") or 0)
        members = proj.area_members(gid0) if (gid0 and proj is not None) else []
        self.var_members = ttk.Label(win, text=("群組成員: " + "、".join(members[:12]) + ("…" if len(members) > 12 else "")) if members else "")
        self.var_members.grid(row=3, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 6))

        self.cmb_dir.bind("<<ComboboxSelected>>", lambda e: self._detail_changed())
        self.cmb_group1.bind("<<ComboboxSelected>>", lambda e: self._detail_changed())
        self.cmb_group2.bind("<<ComboboxSelected>>", lambda e: self._detail_changed())
        self.cmb_group3.bind("<<ComboboxSelected>>", lambda e: self._detail_changed())
        ttk.Button(win, text="關閉", command=win.destroy).grid(row=4, column=3, sticky="e", padx=6, pady=4)
        win.grab_set()

    def open_dialog(self):
        path = filedialog.askopenfilename(
            title="選擇廣播時程檔",
            filetypes=[("廣播時程檔 (*.WEX)", "*.WEX"), ("所有檔案", "*.*")],
        )
        if path:
            self.load_path(path)

    def load_path(self, path):
        try:
            if not os.path.exists(path) and not os.path.isabs(path):
                alt = os.path.join(os.path.dirname(os.path.abspath(__file__)), path)
                if os.path.exists(alt):
                    path = alt
            self.status.config(text=f"載入中: {path}")
            self.root.update_idletasks()
            proj = WexProject(path)
            tabs = proj.tabs
            self.proj = proj
            self.path = path
            self.tab_rows = {}
            self.tab_new_seq = {}
            self.tab_names = {}
            self.tab_deleted = {}
            self.tab_backup = {}
            self.orig_tab_names = dict(proj.tab_names)
            self.visible = []
            self.dirty = set()
            for idx, name, rows in tabs:
                self.tab_rows[idx] = [dict(r.data) for r in rows]
                self.tab_new_seq[idx] = 0
                self.tab_names[idx] = name
                self.tab_deleted[idx] = False
                self.visible.append(idx)
            self.current_tab = self.visible[0] if self.visible else None
            self._refresh_group_tabs()
            self._reload_sheet()
            if hasattr(self, "file_label"):
                self.file_label.config(text=os.path.basename(path))
            total = sum(len(v) for v in self.tab_rows.values())
            self.status.config(
                text=f"{path} | {len(self.visible)} 組鐘聲, 共 {total} 列 | "
                     "點選頁籤切換組別，雙擊列可設定音檔/群組")
            self.shell.update_title()
            return True
        except Exception:
            tb = traceback.format_exc()
            messagebox.showerror("載入失敗", tb[-1500:])
            self.status.config(text="載入失敗")
            return False

    def _refresh_group_tabs(self, select=None):
        if select is None:
            select = self.current_tab
        if select not in self.visible:
            select = self.visible[0] if self.visible else None
        self.current_tab = select
        for tab in self.group_tabs.tabs():
            self.group_tabs.nametowidget(tab).destroy()
        self._tab_ids = {}
        for sid in self.visible:
            frame = ttk.Frame(self.group_tabs, height=1)
            self.group_tabs.add(frame, text=f"{sid}.{self.tab_names.get(sid, '')}")
            self._tab_ids[str(frame)] = sid
            if sid == select:
                self.group_tabs.select(frame)
        self.after_idle(self._resize_group_tabs)

    def _resize_group_tabs(self, event=None):
        width = max(self.tab_canvas.winfo_width(), self.group_tabs.winfo_reqwidth())
        height = self.group_tabs.winfo_reqheight()
        self.tab_canvas.itemconfigure(self._tabs_window, width=width, height=height)
        self.tab_canvas.configure(height=height, scrollregion=(0, 0, width, height))

    def _show_selected_group_tab(self):
        selected = self.group_tabs.select()
        if not selected:
            return
        index = self.group_tabs.index(selected)
        width = self.group_tabs.winfo_width()
        start = end = None
        for x in range(width):
            try:
                matches = self.group_tabs.index(f"@{x},10") == index
            except tk.TclError:
                matches = False
            if matches:
                if start is None:
                    start = x
                end = x + 1
            elif start is not None:
                break
        if start is None:
            return
        left = self.tab_canvas.canvasx(0)
        visible_width = self.tab_canvas.winfo_width()
        if start < left:
            self.tab_canvas.xview_moveto(start / width)
        elif end > left + visible_width:
            self.tab_canvas.xview_moveto(max(start, end - visible_width) / width)

    def _on_group_selected(self, event=None):
        sid = self._tab_ids.get(self.group_tabs.select())
        if sid is not None:
            self._switch_group(sid)

    def _rename_current_group(self):
        if not self.proj or self.current_tab is None:
            return
        sid = self.current_tab
        name = simpledialog.askstring(
            "重新命名", f"第 {sid} 組時程名稱:",
            initialvalue=self.tab_names.get(sid, ""), parent=self.root)
        if name is None:
            return
        name = name.strip()
        if not name:
            return
        self.tab_names[sid] = name
        self.dirty.add(sid)
        self._refresh_group_tabs()
        self.status.config(text=f"已改名為「{name}」（匯出時寫入檔案）")

    def _delete_current_group(self):
        if not self.proj or self.current_tab is None:
            return
        sid = self.current_tab
        if len(self.visible) <= 1:
            messagebox.showinfo("提示", "至少保留一組")
            return
        rows_n = len(self.tab_rows.get(sid, []))
        if not messagebox.askyesno(
                "刪除鐘聲組",
                f"刪除「{self.tab_names.get(sid)}」？\n"
                f"該組共 {rows_n} 列，將於匯出時從檔案刪除。\n\n"
                "（匯出前可按「新增組」復原）"):
            return
        self._commit_sheet_to_model()
        pos = self.visible.index(sid)
        self.tab_backup[sid] = list(self.tab_rows.get(sid, []))
        self.tab_rows[sid] = []
        self.tab_deleted[sid] = True
        self.visible.remove(sid)
        self.dirty.add(sid)
        self.current_tab = self.visible[min(pos, len(self.visible) - 1)]
        self._refresh_group_tabs()
        self._reload_sheet()
        self.status.config(text=f"已刪除「{sid}」（匯出時生效，可由「新增組」復原）")

    def _add_tab(self):
        if not self.proj:
            return
        free = [i for i in sorted(set(self.tab_names) | set(self.tab_backup) | set(self.tab_deleted))
                if self.tab_deleted.get(i)]
        if free:
            sid = free[0]
            self.tab_deleted[sid] = False
            self.tab_rows[sid] = self.tab_backup.pop(sid, [])
            if not self.tab_names.get(sid):
                self.tab_names[sid] = f"時程組{sid}"
        else:
            used = set(self.tab_names) | set(self.tab_rows) | set(self.visible) | set(self.tab_deleted)
            sid = max(used | {0}) + 1
            self.tab_names[sid] = f"時程組{sid}"
            self.tab_rows[sid] = []
            self.tab_deleted[sid] = False
            self.tab_new_seq[sid] = 0
        if sid not in self.visible:
            self.visible.append(sid)
            self.visible.sort()
        self._commit_sheet_to_model()
        self.current_tab = sid
        self.dirty.add(sid)
        self._refresh_group_tabs()
        self._reload_sheet()
        self.status.config(text=f"已新增「{sid}.{self.tab_names[sid]}」")

    def _on_double_click(self, event=None):
        try:
            r = event.row
        except Exception:
            return
        self.open_detail_dialog(r)

    def _row_to_grid(self, d):
        weeks = ["✓" if d.get(f"Week{i}") else "" for i in range(1, 8)]
        y = int(d.get("Clock_year") or 0)
        mo = int(d.get("Clock_month") or 0)
        dd = int(d.get("Clock_day") or 0)
        if (y, mo, dd) == NO_DATE or y == 0:
            date_s = ""
        else:
            date_s = f"{y:04d}-{mo:02d}-{dd:02d}"
        try:
            h = int(d.get("Clock_start_hour") or 0)
            m = int(d.get("Clock_start_min") or 0)
            start_s = f"{h:02d}:{m:02d}:00"
        except Exception:
            h, m = 0, 0
            start_s = "00:00:00"
        secs = int(d.get("Total_time") or 0)
        total_sec = h * 3600 + m * 60 + secs
        eh, er = divmod(total_sec, 3600)
        em, es = divmod(er, 60)
        end_s = f"{eh % 24:02d}:{em:02d}:{es:02d}"
        mode = int(d.get("Mode") or 1)
        on = d.get("Music_on")
        on_val = bool(on) if on is not None else True
        return [
            int(d.get("Music_no") or 0),
            str(d.get("Music_name") or ""),
            *weeks,
            date_s,
            MODE_LABELS.get(mode, str(mode)),
            start_s,
            end_s,
            fmt_status(on_val),
        ]

    @staticmethod
    def _text_units(s):
        s = str(s or "")
        return sum(2 if ord(ch) > 0x2E80 else 1 for ch in s)

    def _apply_column_widths(self):
        """欄寬隨文字長度自動調整（標題 + 內容）。"""
        try:
            data = self.sheet.get_sheet_data()
        except Exception:
            data = []
        if not isinstance(data, list):
            data = []
        for c, h in enumerate(GRID_HEADERS):
            w = self._text_units(h) * 8 + 22
            for row in data[:500]:
                if c < len(row):
                    w = max(w, self._text_units(row[c]) * 8 + 16)
            self.sheet.column_width(c, max(30, min(w, 380)))

    def _reload_sheet(self):
        rows = self.tab_rows.get(self.current_tab, [])
        data = [self._row_to_grid(d) for d in rows]
        self.sheet.set_sheet_data(data=data if data else [[""] * len(GRID_HEADERS)])
        self.sheet.readonly_columns([COL_NO, COL_END])
        self._apply_column_widths()
        for c in range(COL_WEEK0, COL_WEEK0 + 7):
            self.sheet.checkbox_column(c, checked=False, state="normal", text="")
        # 重新套用每列勾選狀態
        for r, d in enumerate(rows):
            for w in range(7):
                if d.get(f"Week{w + 1}"):
                    self.sheet.checkbox_cell(r, COL_WEEK0 + w, checked=True)
        self.sheet.dropdown((0, COL_STATUS, len(data), COL_STATUS),
                             values=["已排入時程", "未排入"], state="readonly")
        self.sheet.dropdown((0, COL_MODE, len(data), COL_MODE),
                            values=list(MODE_LABELS.values()), state="readonly")
        self._detail_row = None

    def _on_cell_edited(self, event=None):
        try:
            r, c = event.row, event.column
        except Exception:
            return
        if self.current_tab is None:
            self._commit_sheet_to_model()
            return
        rows = self.tab_rows.get(self.current_tab, [])
        if r >= len(rows):
            self._commit_sheet_to_model()
            return
        d = rows[r]
        try:
            sheet_data = self.sheet.get_sheet_data()[r]
            self._apply_grid_to_dict(d, sheet_data)
            self.dirty.add(self.current_tab)
        except Exception as ex:
            messagebox.showerror("輸入錯誤", str(ex))
            self._reload_sheet()

    def _apply_grid_to_dict(self, d, grid):
        d["Music_no"] = int(grid[COL_NO] or 0)
        d["Music_name"] = str(grid[COL_NAME] or "")
        for w in range(7):
            d[f"Week{w + 1}"] = str(grid[COL_WEEK0 + w]) in ("True", "✓", "1", "true")
        date_s = str(grid[COL_DATE] or "").strip()
        parsed = parse_date(date_s) if date_s else None
        if parsed:
            d["Clock_year"], d["Clock_month"], d["Clock_day"] = (f"{v:04d}" if i == 0 else f"{v:02d}"
                                                                 for i, v in enumerate(parsed))
        else:
            d["Clock_year"], d["Clock_month"], d["Clock_day"] = "2015", "03", "14"
        mode_s = str(grid[COL_MODE] or "單曲")
        mode = next((k for k, v in MODE_LABELS.items() if v == mode_s), 1)
        d["Mode"] = mode
        h, m = parse_start(str(grid[COL_START] or ""))
        d["Clock_start_hour"] = f"{h:02d}"
        d["Clock_start_min"] = f"{m:02d}"
        d["Music_on"] = parse_status(str(grid[COL_STATUS] or "已排入時程"))

    @staticmethod
    def _cell_meaningful(v):
        if v is True or v == "✓":
            return True
        if v is None or v is False:
            return False
        return bool(str(v).strip())

    def _meaningful_sheet_rows(self, data=None):
        if data is None:
            try:
                data = self.sheet.get_sheet_data()
            except Exception:
                return []
        return [row for row in data if any(self._cell_meaningful(v) for v in row)]

    def _row_dict_from_grid(self, grid, seq):
        d = {
            "MUSIC": -seq, "Mode": 1, "Music_on": 1, "Allcall": True, "Vol": 80,
            "Week_day": False, "Channel": False,
            "Group1": 0, "Group2": 0, "Group3": 0,
            "Dir0": "上課.wav", "Time": 0, "Total_time": 0,
            "End_hour": "00", "End_min": "00", "End_sec": "00",
        }
        try:
            self._apply_grid_to_dict(d, grid)
        except Exception:
            pass
        if not d.get("Music_no"):
            d["Music_no"] = seq
        d.update(row_duration_changes(d, self._audio_durations()))
        return d

    def _create_group_from_rows(self, rows):
        sid = 1
        name = "新時程組"
        for i, d in enumerate(rows, start=1):
            d["Music_no"] = i
        self.tab_names[sid] = name
        self.tab_rows[sid] = rows
        self.tab_deleted[sid] = False
        self.tab_new_seq[sid] = -len(rows)
        if sid not in self.visible:
            self.visible.append(sid)
            self.visible.sort()
        self.current_tab = sid
        self.dirty.add(sid)
        self._refresh_group_tabs()

    def _commit_sheet_to_model(self):
        if getattr(self, "sheet", None) is None:
            return
        try:
            data = self.sheet.get_sheet_data()
        except Exception:
            return
        if not isinstance(data, list):
            return
        if self.current_tab is not None and self._has_unsaved_changes():
            self.dirty.add(self.current_tab)
        if self.current_tab is None:
            rows = [self._row_dict_from_grid(row, i + 1)
                    for i, row in enumerate(self._meaningful_sheet_rows(data))]
            if rows:
                self._create_group_from_rows(rows)
            return
        model = self.tab_rows.setdefault(self.current_tab, [])
        if not model:
            rows = [self._row_dict_from_grid(row, i + 1)
                    for i, row in enumerate(self._meaningful_sheet_rows(data))]
            if rows:
                model.extend(rows)
                self.tab_new_seq[self.current_tab] = -len(rows)
                self.dirty.add(self.current_tab)
            return
        if len(data) > len(model):
            k = self.tab_new_seq.get(self.current_tab, 0)
            for i in range(len(model), len(data)):
                k -= 1
                model.append(self._row_dict_from_grid(data[i], -k))
            self.tab_new_seq[self.current_tab] = k
            self.dirty.add(self.current_tab)
        elif len(data) < len(model):
            del model[len(data):]
            self.dirty.add(self.current_tab)
        for r, d in enumerate(model):
            if r < len(data):
                try:
                    self._apply_grid_to_dict(d, data[r])
                except Exception:
                    pass

    @staticmethod
    def _norm_cell(v):
        if v is True or v == "✓":
            return "✓"
        if v is None or v is False:
            return ""
        return str(v).strip()

    def _has_unsaved_changes(self):
        if self.dirty:
            return True
        try:
            data = self.sheet.get_sheet_data()
        except Exception:
            data = []
        if self.current_tab is None or self.current_tab not in self.tab_rows:
            if self._meaningful_sheet_rows(data):
                return True
        else:
            model = self.tab_rows[self.current_tab]
            if not model and not self._meaningful_sheet_rows(data):
                return False
            if len(data) != len(model):
                return True
            for r, d in enumerate(model):
                if r >= len(data):
                    return True
                if [self._norm_cell(v) for v in self._row_to_grid(d)] != \
                        [self._norm_cell(v) for v in data[r]]:
                    return True
        return False

    def _detail_changed(self):
        r = self._detail_row
        rows = self.tab_rows.get(self.current_tab, [])
        if r is None or r >= len(rows):
            return
        d = rows[r]
        proj = self.proj
        try:
            new_dir = self.var_dir.get()
        except Exception:
            return
        if str(d.get("Dir0") or "") != new_dir:
            d["Dir0"] = new_dir
            auds = self._audio_durations()
            d.update(row_duration_changes(d, auds))
            dur = auds.get(new_dir)
            try:
                self.var_dir_dur.config(text=f"（{duration_to_clock_text(dur)}）" if dur else "")
            except Exception:
                pass
        for i in range(1, 4):
            var = getattr(self, f"var_group{i}")
            s = var.get()
            gid = int(s.split(":", 1)[0]) if s else 0
            d[f"Group{i}"] = gid
        d["Allcall"] = bool(self.var_allcall.get())
        try:
            d["Vol"] = int(self.var_vol.get())
        except Exception:
            d["Vol"] = 80
        self.dirty.add(self.current_tab)
        try:
            gid0 = int(d.get("Group1") or 0)
            members = proj.area_members(gid0) if (gid0 and proj is not None) else []
            self.var_members.config(text=("群組成員: " + "、".join(members[:12]) + ("…" if len(members) > 12 else "")) if members else "")
        except Exception:
            pass
        self._reload_sheet_row(r)

    def _reload_sheet_row(self, r):
        rows = self.tab_rows.get(self.current_tab, [])
        if r >= len(rows):
            return
        for c, v in enumerate(self._row_to_grid(rows[r])):
            self.sheet.set_cell_data(r, c, v)
        for w in range(7):
            self.sheet.checkbox_cell(r, COL_WEEK0 + w, checked=bool(rows[r].get(f"Week{w + 1}")))

    def add_row(self):
        if self.current_tab is None:
            messagebox.showinfo("提示", "請先匯入時程或建立鐘聲組別")
            return
        self._commit_sheet_to_model()
        rows = self.tab_rows.setdefault(self.current_tab, [])
        template = None
        for idx in range(1, 9):
            if self.tab_rows.get(idx):
                template = dict(self.tab_rows[idx][0])
                break
        if template is None:
            template = {"Mode": 1, "Music_on": 1, "Allcall": True, "Vol": 80, "Dir0": "上課.wav",
                        "Time": 53, "Total_time": 53, "End_hour": "00", "End_min": "00", "End_sec": "53",
                        "Week_day": False, "Channel": False, "Group1": 0, "Group2": 0, "Group3": 0}
        self.tab_new_seq[self.current_tab] = self.tab_new_seq.get(self.current_tab, 0) - 1
        template["MUSIC"] = self.tab_new_seq[self.current_tab]
        template["Music_no"] = max((int(r.get("Music_no") or 0) for r in rows), default=0) + 1
        template["Music_name"] = "新增項目"
        for w in range(1, 8):
            template[f"Week{w}"] = False
        template["Clock_year"], template["Clock_month"], template["Clock_day"] = "2015", "03", "14"
        template["Clock_start_hour"] = "08"
        template["Clock_start_min"] = "00"
        rows.append(template)
        self.dirty.add(self.current_tab)
        self._reload_sheet()
        self.status.config(text=f"已新增列（尚未匯出）- 分頁 {self.current_tab}")
        if self._pref_auto_merge():
            self.merge_entries(auto=True)

    def delete_row(self):
        if self.current_tab is None:
            return
        try:
            r = self.sheet.get_currently_selected()[0].row
        except Exception:
            messagebox.showinfo("提示", "請先選取要刪除的列")
            return
        rows = self.tab_rows.get(self.current_tab, [])
        if r >= len(rows):
            return
        name = str(rows[r].get("Music_name") or "")
        if not messagebox.askyesno("確認刪除", f"刪除第 {r + 1} 列「{name}」？"):
            return
        rows.pop(r)
        self.dirty.add(self.current_tab)
        self._reload_sheet()

    def duplicate_row(self):
        if self.current_tab is None:
            return
        try:
            r = self.sheet.get_currently_selected()[0].row
        except Exception:
            messagebox.showinfo("提示", "請先選取要複製的列")
            return
        self._commit_sheet_to_model()
        rows = self.tab_rows.get(self.current_tab, [])
        if r >= len(rows):
            return
        clone = dict(rows[r])
        self.tab_new_seq[self.current_tab] = self.tab_new_seq.get(self.current_tab, 0) - 1
        clone["MUSIC"] = self.tab_new_seq[self.current_tab]
        clone["Music_no"] = max((int(x.get("Music_no") or 0) for x in rows), default=0) + 1
        rows.insert(r + 1, clone)
        self.dirty.add(self.current_tab)
        self._reload_sheet()

    EXCEL_HEADERS = ["序號", "名稱註記", "一", "二", "三", "四", "五", "六", "日",
                     "啟動日期", "播放模式", "啟動時間", "終止時間", "狀態", "音檔"]

    def export_excel(self):
        if not self.proj:
            messagebox.showinfo("提示", "請先開啟 .WEX 檔案")
            return
        self._commit_sheet_to_model()
        rows = self.tab_rows.get(self.current_tab, [])
        group_name = self.tab_names.get(self.current_tab, "")
        path = filedialog.asksaveasfilename(
            title="匯出Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel (*.xlsx)", "*.xlsx"), ("CSV (*.csv)", "*.csv")],
            initialfile=f"{group_name}_時程.xlsx",
        )
        if not path:
            return
        try:
            data = []
            for d in rows:
                weeks = ["✓" if d.get(f"Week{i}") else "" for i in range(1, 8)]
                y = int(d.get("Clock_year") or 0)
                mo = int(d.get("Clock_month") or 0)
                dd = int(d.get("Clock_day") or 0)
                if (y, mo, dd) == NO_DATE or y == 0:
                    date_s = ""
                else:
                    date_s = f"{y:04d}-{mo:02d}-{dd:02d}"
                try:
                    h = int(d.get("Clock_start_hour") or 0)
                    m = int(d.get("Clock_start_min") or 0)
                except Exception:
                    h, m = 0, 0
                start_s = f"{h:02d}:{m:02d}:00"
                secs = int(d.get("Total_time") or 0)
                total_sec = h * 3600 + m * 60 + secs
                eh, er = divmod(total_sec, 3600)
                em, es = divmod(er, 60)
                end_s = f"{eh % 24:02d}:{em:02d}:{es:02d}"
                on = d.get("Music_on")
                on_val = bool(on) if on is not None else True
                data.append([
                    int(d.get("Music_no") or 0),
                    str(d.get("Music_name") or ""),
                    *weeks,
                    date_s,
                    "單曲",
                    start_s,
                    end_s,
                    fmt_status(on_val),
                    str(d.get("Dir0") or ""),
                ])
            if path.lower().endswith(".csv"):
                import csv
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f)
                    w.writerow(self.EXCEL_HEADERS)
                    w.writerows(data)
            else:
                import openpyxl
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = group_name[:31] or "時程"
                ws.append(self.EXCEL_HEADERS)
                for row in data:
                    ws.append(row)
                for i, h in enumerate(self.EXCEL_HEADERS, 1):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(10, len(h) * 2 + 4)
                wb.save(path)
            self.status.config(text=f"已匯出Excel: {path}（{len(data)} 列）")
            messagebox.showinfo("匯出完成",
                                f"已產生：\n{path}\n\n"
                                f"共 {len(data)} 列。\n"
                                "在 Excel 中編輯後，按「匯入Excel」載回。")
        except Exception:
            tb = traceback.format_exc()
            messagebox.showerror("匯出Excel失敗", tb[-1500:])

    def import_excel(self):
        if not self.proj:
            messagebox.showinfo("提示", "請先開啟 .WEX 檔案")
            return
        path = filedialog.askopenfilename(
            title="匯入Excel",
            filetypes=[("Excel (*.xlsx)", "*.xlsx"), ("CSV (*.csv)", "*.csv"), ("所有檔案", "*.*")],
        )
        if not path:
            return
        try:
            raw = self._read_excel_rows(path)
            if raw is None:
                return
            new_rows = self._parse_excel_rows(raw)
            if new_rows is None:
                return
            old_n = len(self.tab_rows.get(self.current_tab, []))
            self.tab_rows[self.current_tab] = new_rows
            self.tab_new_seq[self.current_tab] = -1
            self.dirty.add(self.current_tab)
            self._reload_sheet()
            self.status.config(text=f"已匯入Excel: {path}（{old_n} 列 → {len(new_rows)} 列）")
            messagebox.showinfo("匯入完成",
                                f"已匯入 {len(new_rows)} 列（原 {old_n} 列）。\n\n"
                                "確認無誤後按「匯出 .WEX」寫入檔案。")
        except Exception:
            tb = traceback.format_exc()
            messagebox.showerror("匯入Excel失敗", tb[-1500:])

    def save_template(self):
        path = filedialog.asksaveasfilename(
            title="儲存Excel範本",
            defaultextension=".xlsx",
            filetypes=[("Excel (*.xlsx)", "*.xlsx")],
            initialfile="時程Excel範本.xlsx",
        )
        if not path:
            return
        try:
            from wex_make_template import make_template
            make_template(path)
            self.status.config(text=f"已產生Excel範本: {path}")
            messagebox.showinfo("完成",
                                f"範本已存至：\n{path}\n\n"
                                "在 Excel 中填入時程後，\n"
                                "按「匯入Excel」載入編輯器。")
        except Exception:
            tb = traceback.format_exc()
            messagebox.showerror("產生範本失敗", tb[-1500:])

    def _load_periods(self):
        try:
            with open(PERIOD_SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list) and data:
                cleaned = []
                for p in data:
                    if isinstance(p, (list, tuple)) and len(p) >= 3:
                        cleaned.append([p[0], str(p[1]), str(p[2])])
                if cleaned:
                    return cleaned
        except Exception:
            pass
        return [list(p) for p in DEFAULT_PERIODS]

    def _save_periods(self, periods):
        with open(PERIOD_SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(periods, f, ensure_ascii=False, indent=2)

    def _load_prefs(self):
        try:
            with open(PREFS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except Exception:
            pass
        return {}

    def _save_prefs(self, prefs):
        with open(PREFS_FILE, "w", encoding="utf-8") as f:
            json.dump(prefs, f, ensure_ascii=False, indent=2)

    def _pref_import_write_date(self):
        return bool(self._load_prefs().get("import_write_date", False))

    def _pref_auto_merge(self):
        return bool(self._load_prefs().get("auto_merge", True))

    def open_prefs_dialog(self):
        if getattr(self, "_prefs_win", None) is not None and self._prefs_win.winfo_exists():
            self._prefs_win.lift()
            return
        win = tk.Toplevel(self.root)
        win.title("喜好設定")
        win.geometry("440x300")
        win.minsize(400, 260)
        win.transient(self.root)
        self._prefs_win = win

        prefs = self._load_prefs()
        var_date = tk.BooleanVar(value=bool(prefs.get("import_write_date", False)))
        var_merge = tk.BooleanVar(value=bool(prefs.get("auto_merge", True)))

        ttk.Label(win, text="匯入監考表排程的行為：",
                  font=("Microsoft JhengHei UI", 10, "bold")).pack(anchor="w", padx=12, pady=(12, 6))
        ttk.Checkbutton(
            win,
            text='匯入監考表排程時每則寫入「啟動日期」',
            variable=var_date).pack(anchor="w", padx=20, pady=4)
        ttk.Label(
            win, foreground="#666", justify="left",
            text=("• 勾選：把考程日期寫入每則的啟動日期（該列僅當天執行）\n"
                  "• 不勾（預設）：啟動日期留空，只設星期勾選（每週該星期重複）")).pack(
            anchor="w", padx=20, pady=(0, 8))

        ttk.Label(win, text="條目整併：",
                  font=("Microsoft JhengHei UI", 10, "bold")).pack(anchor="w", padx=12, pady=(4, 2))
        ttk.Checkbutton(
            win,
            text='匯入或新增時程條目時，自動整併相同條目（預設開）',
            variable=var_merge).pack(anchor="w", padx=20, pady=4)
        ttk.Label(
            win, foreground="#666", justify="left",
            text=("• 整併：同名稱格式＋同啟動時間＋同音檔的條目併成一列\n"
                  "  （後面併入前面：星期取聯集、名稱併日期如「10/28, 29-1-起始」）\n"
                  "• 不勾：僅可按工作列「條目整併」手動執行")).pack(
            anchor="w", padx=20, pady=(0, 8))

        btn_bar = ttk.Frame(win)
        btn_bar.pack(fill="x", padx=12, pady=(0, 10))

        def save_prefs():
            self._save_prefs({
                "import_write_date": bool(var_date.get()),
                "auto_merge": bool(var_merge.get()),
            })
            s_date = "已啟用" if var_date.get() else "已關閉"
            s_merge = "已啟用" if var_merge.get() else "已關閉"
            self.status.config(
                text=f"喜好設定已儲存（匯入寫入啟動日期：{s_date}；自動整併：{s_merge}）")
            win.destroy()

        ttk.Button(btn_bar, text="儲存", command=save_prefs).pack(side="right", padx=4)
        ttk.Button(btn_bar, text="取消", command=win.destroy).pack(side="right", padx=4)
        win.grab_set()

    def open_period_settings(self):
        if getattr(self, "_period_win", None) is not None and self._period_win.winfo_exists():
            self._period_win.lift()
            return
        win = tk.Toplevel(self.root)
        win.title("設定課表作息")
        win.geometry("470x480")
        win.minsize(420, 360)
        win.transient(self.root)
        self._period_win = win

        ttk.Label(win, text="設定學校正常上課的課程節數（匯出 .WEX 後由主機套用）",
                  foreground="#555").pack(padx=8, pady=(8, 2), anchor="w")

        sheet = Sheet(win, width=430, height=380,
                      header_font=("Microsoft JhengHei UI", 10, "bold"),
                      table_font=("Microsoft JhengHei UI", 10, "normal"))
        sheet.pack(fill="both", expand=True, padx=8, pady=4)
        localize_sheet_menu(sheet)
        sheet.headers(PERIOD_HEADERS)
        sheet.enable_bindings()
        lock_sheet_format(sheet)
        try:
            sheet.MT.RI.drag_and_drop_enabled = True
            sheet.ops.row_drag_and_drop_perform = True
        except Exception:
            pass
        periods = self._load_periods()
        sheet.set_sheet_data(data=[list(p) for p in periods])
        sheet.readonly_columns([0])
        sheet.column_width(0, 50)
        sheet.column_width(1, 170)
        sheet.column_width(2, 120)

        btn_bar = ttk.Frame(win)
        btn_bar.pack(fill="x", padx=8, pady=(0, 8))

        def restyle():
            sheet.readonly_columns([0])
            sheet.column_width(0, 50)
            sheet.column_width(1, 170)
            sheet.column_width(2, 120)

        def add_period():
            data = sheet.get_sheet_data()
            next_no = max((int(r[0]) for r in data if str(r[0]).isdigit()), default=0) + 1
            data.append([next_no, f"第{next_no}節", "08:00:00"])
            sheet.set_sheet_data(data=data)
            restyle()

        def del_period():
            try:
                r = sheet.get_currently_selected()[0].row
            except Exception:
                messagebox.showinfo("提示", "請先選取要刪除的節次", parent=win)
                return
            data = sheet.get_sheet_data()
            if 0 <= r < len(data):
                data.pop(r)
                sheet.set_sheet_data(data=data)
                restyle()

        def reset_defaults():
            sheet.set_sheet_data(data=[list(p) for p in DEFAULT_PERIODS])
            restyle()

        def import_from_schedule():
            if not self.proj:
                messagebox.showinfo("提示", "目前工作區沒有載入時程", parent=win)
                return
            self._commit_sheet_to_model()
            rows = self.tab_rows.get(self.current_tab, [])
            if not rows:
                messagebox.showinfo("提示", "目前時程組沒有資料可匯入", parent=win)
                return
            sel = []
            try:
                sel = sorted(set(int(r) for r in self.sheet.MT.get_selected_rows()))
            except Exception:
                sel = []
            picked = [rows[i] for i in sel if 0 <= i < len(rows)] if sel else list(rows)
            new_data = []
            for d in picked:
                name = str(d.get("Music_name") or "").strip()
                if not name:
                    continue
                try:
                    h = int(d.get("Clock_start_hour") or 0)
                    m = int(d.get("Clock_start_min") or 0)
                except Exception:
                    h, m = 0, 0
                no = int(d.get("Music_no") or 0) or len(new_data) + 1
                new_data.append([no, name, f"{h:02d}:{m:02d}:00"])
            if not new_data:
                messagebox.showinfo("提示", "沒有可匯入的列（名稱為空）", parent=win)
                return
            cur = sheet.get_sheet_data()
            gname = self.tab_names.get(self.current_tab) or ""
            scope = f"選取的 {len(new_data)} 列" if sel else f"「{gname}」的 {len(new_data)} 列"
            if cur and not messagebox.askyesno(
                    "自時程匯入",
                    f"以目前時程 {scope} 取代現在的 {len(cur)} 節？",
                    parent=win):
                return
            sheet.set_sheet_data(data=new_data)
            restyle()
            src_hint = "（僅匯入選取列）" if sel else ""
            self.status.config(text=f"已自時程匯入課表作息：{len(new_data)} 節{src_hint} — 按「儲存」生效")

        def save_periods():
            data = sheet.get_sheet_data()
            cleaned = []
            for row in data:
                name = str(row[1] or "").strip()
                if not name:
                    continue
                cleaned.append([
                    int(row[0]) if str(row[0]).isdigit() else len(cleaned) + 1,
                    name,
                    str(row[2] or "08:00:00").strip(),
                ])
            self._save_periods(cleaned)
            self.status.config(text=f"課表作息已儲存（{len(cleaned)} 節）")
            messagebox.showinfo("已儲存", f"課表作息已儲存（{len(cleaned)} 節）", parent=win)
            win.destroy()

        ttk.Button(btn_bar, text="新增節次", command=add_period).pack(side="left", padx=(0, 4))
        ttk.Button(btn_bar, text="刪除節次", command=del_period).pack(side="left", padx=(0, 4))
        ttk.Button(btn_bar, text="恢復預設", command=reset_defaults).pack(side="left", padx=(0, 4))
        ttk.Button(btn_bar, text="自時程匯入", command=import_from_schedule).pack(side="left", padx=(0, 4))
        ttk.Button(btn_bar, text="儲存", command=save_periods).pack(side="right", padx=4)
        ttk.Button(btn_bar, text="取消", command=win.destroy).pack(side="right", padx=4)
        win.grab_set()

    def _read_excel_rows(self, path):
        if path.lower().endswith(".csv"):
            import csv
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.reader(f))
        elif path.lower().endswith(".xls"):
            import xlrd
            wb = xlrd.open_workbook(path)
            sh = wb.sheet_by_index(0)
            rows = []
            for r in range(sh.nrows):
                rows.append([
                    int(v) if isinstance(v, float) and v == int(v) else v
                    for v in sh.row_values(r)
                ])
        else:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        if not rows:
            messagebox.showerror("匯入失敗", "檔案是空的")
            return None
        header = [str(c or "").strip() for c in rows[0]]
        if "名稱註記" not in header:
            messagebox.showerror("匯入失敗",
                                 "找不到「名稱註記」欄位。\n"
                                 "請確認第一列是欄位名稱，且包含：\n"
                                 "序號、名稱註記、一~日、啟動日期、啟動時間、音檔")
            return None
        return rows

    def _parse_excel_rows(self, raw_rows):
        header = [str(c or "").strip() for c in raw_rows[0]]
        col = {h: i for i, h in enumerate(header)}
        need = ["名稱註記"]
        if any(n not in col for n in need):
            messagebox.showerror("匯入失敗", "缺少必要欄位：名稱註記")
            return None
        week_true = {"✓", "v", "V", "1", "√", "●", "x", "X", "y", "Y", "true", "TRUE", "True", "○"}
        mode_map = {"單曲": 1}
        audio_files = set(self._audio_durations())
        out = []
        seq = 0
        for ri, row in enumerate(raw_rows[1:], start=2):
            def cell(name):
                i = col.get(name)
                if i is None or i >= len(row):
                    return ""
                v = row[i]
                return "" if v is None else str(v).strip()

            name = cell("名稱註記")
            if not name:
                continue
            seq += 1
            d = {
                "MUSIC": -seq,
                "Music_no": int(cell("序號")) if cell("序號").isdigit() else seq,
                "Music_name": name,
                "Mode": mode_map.get(cell("播放模式"), 1),
                "Music_on": 0 if cell("狀態") == "未排入" else 1,
                "Allcall": True,
                "Vol": 80,
                "Week_day": False,
                "Channel": False,
                "Group1": 0, "Group2": 0, "Group3": 0,
            }
            for w, label in enumerate(("一", "二", "三", "四", "五", "六", "日"), start=1):
                d[f"Week{w}"] = cell(label) in week_true
            date_s = cell("啟動日期")
            if date_s:
                try:
                    parts = date_s.replace("/", "-").split("-")
                    y, m, dd = int(parts[0]), int(parts[1]), int(parts[2])
                    d["Clock_year"], d["Clock_month"], d["Clock_day"] = f"{y:04d}", f"{m:02d}", f"{dd:02d}"
                except Exception:
                    d["Clock_year"], d["Clock_month"], d["Clock_day"] = "2015", "03", "14"
            else:
                d["Clock_year"], d["Clock_month"], d["Clock_day"] = "2015", "03", "14"
            time_s = cell("啟動時間").replace("：", ":")
            try:
                tp = time_s.split(":")
                h, m = int(tp[0]), int(tp[1])
                d["Clock_start_hour"] = f"{h:02d}"
                d["Clock_start_min"] = f"{m:02d}"
            except Exception:
                d["Clock_start_hour"], d["Clock_start_min"] = "00", "00"
            dir0 = cell("音檔")
            if dir0 and dir0 in audio_files:
                d["Dir0"] = dir0
            else:
                d["Dir0"] = "上課.wav"
            d.update(row_duration_changes(d, self._audio_durations()))
            out.append(d)
        return out

    def export_dialog(self):
        self._commit_sheet_to_model()
        if self.proj is None:
            messagebox.showerror("無法匯出", "請先使用「匯入舊WEX檔」選擇完整的 WEX 檔案。")
            return False
        if not self.visible or self.current_tab is None:
            messagebox.showinfo("提示", "目前沒有任何時程組")
            return False
        path = filedialog.asksaveasfilename(
            title="匯出 .WEX",
            defaultextension=".WEX",
            filetypes=[("廣播時程檔 (*.WEX)", "*.WEX")],
            initialfile=os.path.splitext(os.path.basename(self.path))[0] + "_修改.WEX",
        )
        if not path:
            return False
        try:
            statements = self._build_statements()
            if not statements:
                messagebox.showinfo("無變更", "沒有偵測到變更，仍可匯出完整副本")
            export_wex(self.path, path, statements)
            self.dirty.clear()
            self._refresh_group_tabs()
            self._reload_sheet()
            extra_slots = [i for i in self.visible if i > 8 and not self.tab_deleted.get(i)]
            extra_note = ""
            if extra_slots:
                extra_note = ("\n\n注意：包含第9組以上的頁籤（" +
                              "、".join(str(i) for i in extra_slots) +
                              "），\n為本工具相容性擴充，業者 WSC3100 主機可能不顯示，\n請先以小量資料測試匯入。")
            self.status.config(text=f"已匯出: {path}（{len(statements)} 個 SQL 變更）")
            messagebox.showinfo("匯出完成",
                                f"已產生：\n{path}\n\n"
                                f"共 {len(statements)} 個資料庫變更。{extra_note}\n\n"
                                "請交由 WSC3100 主機匯入驗證。")
            return True
        except Exception:
            tb = traceback.format_exc()
            messagebox.showerror("匯出失敗", tb[-1500:])
            return False

    def _build_statements(self):
        proj = self.proj
        statements = []
        mirror = self._music_t_mirror()
        name_changed = False
        orig_ids = set(proj.schedule_ids)
        all_ids = sorted(orig_ids | set(self.tab_rows) | set(self.tab_names))
        template_table = f"MUSIC{max(orig_ids)}" if orig_ids else "MUSIC8"
        for idx in all_ids:
            table = f"MUSIC{idx}"
            deleted = bool(self.tab_deleted.get(idx))
            rows = [] if deleted else self.tab_rows.get(idx, [])
            table_existed = table in proj._catalog
            orig = {r.music_id: r.data for r in proj.load_rows(table)} if table_existed else {}
            is_mirror = mirror == table
            schedule_changed = False
            if not table_existed and not deleted and rows:
                statements.append(f"SELECT * INTO [{table}] FROM [{template_table}] WHERE 1=0")
            seen_ids = set()
            new_rows = []
            for d in rows:
                mid = int(d.get("MUSIC") or 0)
                if mid > 0 and mid in orig:
                    seen_ids.add(mid)
                    changes = self._diff(orig[mid], d, proj)
                    if changes:
                        statements.append(build_update_row(table, mid, changes))
                        if is_mirror:
                            statements.append(build_update_row("MUSIC_T", mid, changes))
                        schedule_changed = True
                else:
                    new_rows.append(d)
            for mid in orig:
                if mid not in seen_ids:
                    statements.append(build_delete_row(table, mid))
                    if is_mirror:
                        statements.append(build_delete_row("MUSIC_T", mid))
                    schedule_changed = True
            next_id = None
            for d in new_rows:
                row_data = dict(d)
                if next_id is None:
                    next_id = proj.next_music_id(table)
                row_data["MUSIC"] = next_id
                next_id += 1
                row_data["Music_no"] = int(row_data.get("Music_no") or 0)
                row_data.update(row_duration_changes(row_data, proj.audio_durations))
                statements.append(build_insert_row(table, row_data))
                if is_mirror:
                    statements.append(build_insert_row("MUSIC_T", row_data))
                schedule_changed = True
            if schedule_changed and idx <= 8:
                stamp = f"'{datetime.now():%Y/%m/%d %H:%M:%S}'"
                statements.append(f"UPDATE [SETTING] SET [Up_music{idx}] = {stamp}")
            old_name = (proj.tab_names.get(idx) or "").strip()
            new_name = "" if deleted else (self.tab_names.get(idx) or "").strip()
            if old_name != new_name:
                if idx in proj.tab_names:
                    statements.append(
                        f"UPDATE [MUSIC] SET [Music_name] = {sql_quote(new_name)} WHERE [MUSIC] = {idx}")
                elif new_name:
                    statements.append(
                        f"INSERT INTO [MUSIC] ([MUSIC], [Music_name]) VALUES ({idx}, {sql_quote(new_name)})")
                name_changed = True
        if name_changed:
            statements.append(
                f"UPDATE [SETTING] SET [Up_music] = '{datetime.now():%Y/%m/%d %H:%M:%S}'")
        return statements

    def _music_t_mirror(self):
        if "MUSIC_T" not in self.proj._catalog:
            return None
        mt_rows = self.proj.load_rows("MUSIC_T")
        if not mt_rows:
            return None
        mirror_rows = {r.music_id: r.data for r in mt_rows}
        matches = []
        for idx in range(1, 9):
            table = f"MUSIC{idx}"
            rows = {r.music_id: r.data for r in self.proj.load_rows(table)}
            if rows == mirror_rows:
                matches.append(table)
        return matches[0] if len(matches) == 1 else None

    def _diff(self, orig, cur, proj):
        changes = {}

        def norm(v):
            if v is None:
                return None
            if isinstance(v, bool):
                return v
            if isinstance(v, (int, float)):
                return v
            return str(v)

        def s(v):
            return "" if v is None else str(v)

        if s(orig.get("Music_name")) != s(cur.get("Music_name")):
            changes["Music_name"] = str(cur.get("Music_name") or "")
        if int(orig.get("Music_no") or 0) != int(cur.get("Music_no") or 0):
            changes["Music_no"] = int(cur.get("Music_no") or 0)
        for w in range(1, 8):
            if bool(orig.get(f"Week{w}")) != bool(cur.get(f"Week{w}")):
                changes[f"Week{w}"] = bool(cur.get(f"Week{w}"))
        for col in ("Clock_year", "Clock_month", "Clock_day"):
            if s(orig.get(col)) != s(cur.get(col)):
                changes[col] = s(cur.get(col))
        for col in ("Clock_start_hour", "Clock_start_min"):
            if s(orig.get(col)) != s(cur.get(col)):
                changes[col] = s(cur.get(col))
        if int(orig.get("Mode") or 1) != int(cur.get("Mode") or 1):
            changes["Mode"] = int(cur.get("Mode") or 1)
        o_on = bool(orig.get("Music_on")) if orig.get("Music_on") is not None else True
        c_on = bool(cur.get("Music_on")) if cur.get("Music_on") is not None else True
        if o_on != c_on:
            changes["Music_on"] = 1 if c_on else 0
        for col in ("Dir0",):
            if s(orig.get(col)) != s(cur.get(col)):
                changes[col] = s(cur.get(col))
        for col in ("Time", "Total_time", "End_hour", "End_min", "End_sec"):
            if s(orig.get(col)) != s(cur.get(col)):
                changes[col] = s(cur.get(col)) if col.startswith("End_") else int(cur.get(col) or 0)
        for i in range(1, 4):
            if int(orig.get(f"Group{i}") or 0) != int(cur.get(f"Group{i}") or 0):
                changes[f"Group{i}"] = int(cur.get(f"Group{i}") or 0)
        if bool(orig.get("Allcall")) != bool(cur.get("Allcall")):
            changes["Allcall"] = bool(cur.get("Allcall"))
        if (orig.get("Vol") or 0) != (cur.get("Vol") or 0):
            changes["Vol"] = int(cur.get("Vol") or 0)
        return changes


class EditorApp:
    """主視窗：選單、工作區與狀態列。"""

    def __init__(self, root, initial_path=None):
        self.root = root
        root.title("廣播時程編輯器")
        root.geometry("1100x700")
        root.minsize(880, 520)

        menubar = tk.Menu(root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="新鐘聲組別", command=self.new_bell_group)
        file_menu.add_separator()
        file_menu.add_command(label="開啟 .WEX ...", command=self.open_dialog, accelerator="Ctrl+O")
        file_menu.add_command(label="匯出 .WEX ...", command=self.export_dialog, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="結束", command=self.on_close)
        menubar.add_cascade(label="檔案", menu=file_menu)
        self.file_menu = file_menu
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="設定課表作息 ...",
                                  command=lambda: self.active_workspace.open_period_settings())
        settings_menu.add_command(label="喜好設定 ...",
                                  command=lambda: self.active_workspace.open_prefs_dialog())
        menubar.add_cascade(label="設定", menu=settings_menu)
        root.config(menu=menubar)

        self.status = ttk.Label(root, anchor="w", relief="sunken",
                                text="請按「匯入舊WEX檔」選擇完整的 WEX 檔案")
        self.status.pack(fill="x", side="bottom")

        self.active_workspace = Workspace(root, self)
        self.active_workspace.pack(fill="both", expand=True, padx=6, pady=(0, 2))

        root.bind("<Control-o>", lambda e: self.open_dialog())
        root.bind("<Control-s>", lambda e: self.export_dialog())
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        if initial_path:
            root.after(100, lambda: self.open_path(initial_path))

    def _confirm_discard_changes(self, action_label):
        """有未儲存修改時詢問；回傳 True = 可繼續動作。"""
        ws = self.active_workspace
        if not ws._has_unsaved_changes():
            return True
        ans = messagebox.askyesnocancel(
            "未儲存的修改",
            f"目前有尚未匯出的修改。\n\n"
            f"要先匯出 WEX 嗎？\n\n"
            f"是：匯出後{action_label}\n"
            f"否：放棄修改並{action_label}\n"
            f"取消：返回程式",
            parent=self.root)
        if ans is None:
            return False
        if ans and not ws.export_dialog():
            return False
        return True

    def on_close(self):
        if not self._confirm_discard_changes("關閉"):
            return
        self.root.destroy()

    def update_title(self):
        path = self.active_workspace.path
        base = "廣播時程編輯器"
        self.root.title(f"{base} - {os.path.basename(path)}" if path else base)

    def new_bell_group(self):
        if not self._confirm_discard_changes("清空"):
            return
        self.active_workspace.clear_workspace()
        self.update_title()

    def open_dialog(self):
        path = filedialog.askopenfilename(
            title="選擇廣播時程檔",
            filetypes=[("廣播時程檔 (*.WEX)", "*.WEX"), ("所有檔案", "*.*")],
        )
        if path:
            self.open_path(path)

    def open_path(self, path):
        if self._confirm_discard_changes("開啟其他檔案"):
            return self.active_workspace.load_path(path)
        return False

    def export_dialog(self):
        return self.active_workspace.export_dialog()


def main():
    import argparse
    parser = argparse.ArgumentParser(description="廣播時程編輯器")
    parser.add_argument("path", nargs="?")
    args = parser.parse_args()
    root = tk.Tk()
    EditorApp(root, initial_path=args.path)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        try:
            input("發生錯誤，按 Enter 結束... ")
        except Exception:
            pass
