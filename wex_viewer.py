import argparse
import io
import os
import shutil
import struct
import sys
import tempfile
from datetime import datetime, timedelta

MDB_BOOL = 0x01
MDB_BYTE = 0x02
MDB_INT = 0x03
MDB_LONGINT = 0x04
MDB_MONEY = 0x05
MDB_FLOAT = 0x06
MDB_DOUBLE = 0x07
MDB_DATETIME = 0x08
MDB_BINARY = 0x09
MDB_TEXT = 0x0A
MDB_OLE = 0x0B
MDB_MEMO = 0x0C
MDB_REPID = 0x0F
MDB_NUMERIC = 0x10
MDB_COMPLEX = 0x12

JET4 = {
    "pg_size": 4096,
    "row_count_offset": 0x0C,
    "tab_num_rows_offset": 16,
    "tab_num_cols_offset": 45,
    "tab_num_idxs_offset": 47,
    "tab_num_ridxs_offset": 51,
    "tab_usage_map_offset": 55,
    "tab_first_dpg_offset": 56,
    "tab_cols_start_offset": 63,
    "tab_ridx_entry_size": 12,
    "col_scale_offset": 11,
    "col_prec_offset": 12,
    "col_flags_offset": 15,
    "col_size_offset": 23,
    "col_num_offset": 5,
    "tab_col_entry_size": 25,
    "tab_free_map_offset": 59,
    "tab_col_offset_var": 7,
    "tab_col_offset_fixed": 21,
    "tab_row_col_num_offset": 9,
}
OFFSET_MASK = 0x1FFF
HEADER_RC4_KEY = bytes([0xC7, 0xDA, 0x39, 0x6B])


def rc4(key, data):
    s = list(range(256))
    j = 0
    for i in range(256):
        j = (j + s[i] + key[i % len(key)]) % 256
        s[i], s[j] = s[j], s[i]
    out = bytearray()
    i = j = 0
    for b in data:
        i = (i + 1) % 256
        j = (j + s[i]) % 256
        s[i], s[j] = s[j], s[i]
        out.append(b ^ s[(s[i] + s[j]) % 256])
    return bytes(out)


def i16(buf, off):
    return struct.unpack_from("<H", buf, off)[0]


def i32(buf, off):
    return struct.unpack_from("<I", buf, off)[0]


def i64(buf, off):
    return struct.unpack_from("<Q", buf, off)[0]


def f32(buf, off):
    return struct.unpack_from("<f", buf, off)[0]


def f64(buf, off):
    return struct.unpack_from("<d", buf, off)[0]


class Column:
    __slots__ = (
        "col_type", "col_num", "var_col_num", "is_fixed",
        "fixed_offset", "col_size", "scale", "prec", "name",
    )

    def __init__(self):
        self.col_type = 0
        self.col_num = 0
        self.var_col_num = 0
        self.is_fixed = False
        self.fixed_offset = 0
        self.col_size = 0
        self.scale = 0
        self.prec = 0
        self.name = ""


class CatalogEntry:
    def __init__(self, name, obj_type, table_pg, flags):
        self.name = name
        self.obj_type = obj_type
        self.table_pg = table_pg
        self.flags = flags


class Table:
    def __init__(self, entry):
        self.entry = entry
        self.num_rows = 0
        self.num_cols = 0
        self.num_var_cols = 0
        self.num_real_idxs = 0
        self.columns = []


class PageCursor:
    def __init__(self, db, page_no, pos):
        self.db = db
        self.page = db.read_page(page_no)
        self.pos = pos

    def read(self, n):
        out = bytearray()
        pg_size = JET4["pg_size"]
        while n > 0:
            if self.pos >= pg_size:
                next_pg = i32(self.page, 4)
                self.page = self.db.read_page(next_pg)
                self.pos = 8
            take = min(n, pg_size - self.pos)
            out += self.page[self.pos:self.pos + take]
            self.pos += take
            n -= take
        return bytes(out)


class JetDatabase:
    def __init__(self, source):
        if isinstance(source, (str, os.PathLike)):
            with open(source, "rb") as f:
                self.data = f.read()
        else:
            self.data = source.read()
        self.num_pages = len(self.data) // JET4["pg_size"]
        self.db_key = 0
        self._decrypt_header()

    def _decrypt_header(self):
        head = bytearray(self.data[:0x18 + 128])
        head[0x18:0x18 + 128] = rc4(HEADER_RC4_KEY, bytes(head[0x18:0x18 + 128]))
        self.code_page = i16(head, 0x3C)
        self.db_key = i32(head, 0x3E)

    def read_page(self, page_no):
        pg_size = JET4["pg_size"]
        off = page_no * pg_size
        page = self.data[off:off + pg_size]
        if len(page) < pg_size:
            page = page + b"\x00" * (pg_size - len(page))
        if page_no != 0 and self.db_key != 0:
            key = (self.db_key ^ page_no) & 0xFFFFFFFF
            page = rc4(struct.pack("<I", key), page)
        return page

    def find_row_in_page(self, page, row):
        rco = JET4["row_count_offset"]
        start = i16(page, rco + 2 + row * 2)
        if row == 0:
            next_start = JET4["pg_size"]
        else:
            next_start = i16(page, rco + row * 2) & OFFSET_MASK
        length = next_start - (start & OFFSET_MASK)
        return start, length

    def find_pg_row(self, pg_row):
        page_no = pg_row >> 8
        row = pg_row & 0xFF
        page = self.read_page(page_no)
        start, length = self.find_row_in_page(page, row)
        start &= OFFSET_MASK
        return page, start, length

    def read_catalog(self):
        entry = CatalogEntry("MSysObjects", 1, 2, 0)
        table = self.read_table(entry)
        self.read_columns(table)
        name_idx = self._col_index(table, "Name")
        type_idx = self._col_index(table, "Type")
        flags_idx = self._col_index(table, "Flags")
        id_idx = self._col_index(table, "Id")
        if None in (name_idx, type_idx, flags_idx, id_idx):
            raise RuntimeError("MSysObjects 欄位讀取失敗")
        rows = self.fetch_rows(table)
        self.resolve_text_columns(table, rows)
        catalog = []
        for row in rows:
            obj_type = row[type_idx]
            if obj_type is None:
                continue
            obj_type = int(obj_type) & 0x7F
            obj_id = int(row[id_idx] or 0)
            flags = int(row[flags_idx] or 0)
            catalog.append(CatalogEntry(str(row[name_idx] or ""), obj_type, obj_id & 0x00FFFFFF, flags))
        return catalog

    @staticmethod
    def _col_index(table, col_name):
        for i, col in enumerate(table.columns):
            if col.name.lower() == col_name.lower():
                return i
        return None

    def read_table(self, entry):
        page = self.read_page(entry.table_pg)
        if page[0] != 0x02:
            raise RuntimeError(f"資料表 {entry.name} 的定義頁無效 (page {entry.table_pg})")
        fmt = JET4
        table = Table(entry)
        table.num_rows = i32(page, fmt["tab_num_rows_offset"])
        table.num_var_cols = i16(page, fmt["tab_num_cols_offset"] - 2)
        table.num_cols = i16(page, fmt["tab_num_cols_offset"])
        table.num_real_idxs = i32(page, fmt["tab_num_ridxs_offset"])
        return table

    def read_columns(self, table):
        fmt = JET4
        pos = fmt["tab_cols_start_offset"] + table.num_real_idxs * fmt["tab_ridx_entry_size"]
        cursor = PageCursor(self, table.entry.table_pg, pos)
        columns = []
        for _ in range(table.num_cols):
            raw = cursor.read(fmt["tab_col_entry_size"])
            col = Column()
            col.col_type = raw[0]
            col.col_num = raw[fmt["col_num_offset"]]
            col.var_col_num = i16(raw, fmt["tab_col_offset_var"])
            col.scale = raw[fmt["col_scale_offset"]]
            col.prec = raw[fmt["col_prec_offset"]]
            col.is_fixed = bool(raw[fmt["col_flags_offset"]] & 0x01)
            col.fixed_offset = i16(raw, fmt["tab_col_offset_fixed"])
            if col.col_type != MDB_BOOL:
                col.col_size = i16(raw, fmt["col_size_offset"])
            columns.append(col)
        for col in columns:
            name_len = i16(cursor.read(2), 0)
            raw_name = cursor.read(name_len)
            col.name = raw_name.decode("utf-16-le", errors="replace").rstrip("\x00")
        columns.sort(key=lambda c: c.col_num)
        table.columns = columns

    def iter_data_pages(self, table):
        pages = self.table_data_pages(table)
        if pages is None:
            for page_no in range(1, self.num_pages):
                page = self.read_page(page_no)
                if page[0] == 0x01 and i32(page, 4) == table.entry.table_pg:
                    yield page
            return
        for page_no in pages:
            page = self.read_page(page_no)
            if page[0] == 0x01 and i32(page, 4) == table.entry.table_pg:
                yield page

    def read_usage_map_row(self, table):
        page = self.read_page(table.entry.table_pg)
        pg_row = i32(page, JET4["tab_usage_map_offset"])
        pg_no = pg_row >> 8
        row_no = pg_row & 0xFF
        if pg_no == 0 or pg_no >= self.num_pages:
            return None
        map_page = self.read_page(pg_no)
        if map_page[0] != 0x01:
            return None
        rco = JET4["row_count_offset"]
        nrows = i16(map_page, rco)
        if row_no >= nrows:
            return None
        start = i16(map_page, rco + 2 + row_no * 2) & OFFSET_MASK
        if row_no == 0:
            next_start = JET4["pg_size"]
        else:
            next_start = i16(map_page, rco + row_no * 2) & OFFSET_MASK
        length = next_start - start
        if length <= 0:
            return None
        return map_page[start:start + length]

    def table_data_pages(self, table):
        """回傳資料頁號清單；usage map 無法解析時回傳 None（改用全掃描）。"""
        row = self.read_usage_map_row(table)
        if not row or len(row) < 5:
            return None
        pages = []
        if row[0] == 0:
            pgnum = i32(row, 1)
            bitmap = row[5:]
            for i, b in enumerate(bitmap):
                if b:
                    for bit in range(8):
                        if b & (1 << bit):
                            pages.append(pgnum + i * 8 + bit)
        elif row[0] == 1:
            usage_bitlen = (JET4["pg_size"] - 4) * 8
            max_maps = (len(row) - 1) // 4
            for map_ind in range(max_maps):
                map_pg = i32(row, map_ind * 4 + 1)
                if map_pg == 0 or map_pg >= self.num_pages:
                    continue
                bmp_page = self.read_page(map_pg)
                if bmp_page[0] != 0x05:
                    continue
                bitmap = bmp_page[4:]
                for i, b in enumerate(bitmap):
                    if b:
                        for bit in range(8):
                            if b & (1 << bit):
                                pages.append(map_ind * usage_bitlen + i * 8 + bit)
        else:
            return None
        return sorted(set(pages))

    def fetch_rows(self, table):
        rows = []
        rco = JET4["row_count_offset"]
        for page in self.iter_data_pages(table):
            nrows = i16(page, rco)
            for row_no in range(nrows):
                start, length = self.find_row_in_page(page, row_no)
                if start & 0x4000:
                    continue
                start &= OFFSET_MASK
                if length <= 0:
                    continue
                try:
                    values = self._crack_row(table, page, start, length)
                except Exception:
                    continue
                if values is not None:
                    rows.append(values)
        return rows

    def _crack_row(self, table, page, row_start, row_size):
        fmt = JET4
        row_end = row_start + row_size - 1
        row_cols = i16(page, row_start)
        col_count_size = 2
        bitmask_sz = (row_cols + 7) // 8
        if bitmask_sz + 1 >= row_size:
            return None
        nullmask = page[row_end - bitmask_sz + 1:row_end + 1]
        var_offsets = []
        row_var_cols = 0
        if table.num_var_cols > 0:
            row_var_cols = i16(page, row_end - bitmask_sz - 1)
            var_offsets = [i16(page, row_end - bitmask_sz - 3 - i * 2) for i in range(row_var_cols + 1)]
        result = [None] * len(table.columns)
        fixed_found = 0
        row_fixed_cols = row_cols - row_var_cols
        for idx, col in enumerate(table.columns):
            byte_num = col.col_num // 8
            bit_num = col.col_num % 8
            is_null = not (nullmask[byte_num] & (1 << bit_num))
            if col.is_fixed and fixed_found < row_fixed_cols:
                col_start = row_start + col.fixed_offset + col_count_size
                result[idx] = self._decode_value(col, page, col_start, col.col_size, is_null)
                fixed_found += 1
            elif (not col.is_fixed) and col.var_col_num < row_var_cols:
                vstart = var_offsets[col.var_col_num]
                vsize = var_offsets[col.var_col_num + 1] - vstart
                result[idx] = self._decode_value(col, page, row_start + vstart, vsize, is_null)
            else:
                result[idx] = None
        return result

    def _decode_value(self, col, page, start, size, is_null):
        if col.col_type == MDB_BOOL:
            return not is_null
        if is_null or size < 0:
            return None
        if start + size > len(page):
            return None
        t = col.col_type
        if t == MDB_BYTE:
            return page[start]
        if t == MDB_INT:
            return struct.unpack_from("<h", page, start)[0]
        if t in (MDB_LONGINT, MDB_COMPLEX):
            return struct.unpack_from("<i", page, start)[0]
        if t == MDB_FLOAT:
            return f32(page, start)
        if t == MDB_DOUBLE:
            return f64(page, start)
        if t == MDB_MONEY:
            return i64(page, start) / 10000.0
        if t == MDB_DATETIME:
            return self._decode_datetime(f64(page, start))
        if t == MDB_TEXT:
            return ("__text__", page[start:start + size])
        if t == MDB_MEMO:
            return self._decode_memo(page, start, size)
        if t == MDB_OLE:
            return "[OLE]"
        if t == MDB_BINARY:
            raw = page[start:start + size]
            return raw.hex().upper()
        if t == MDB_REPID:
            return self._format_guid(page[start:start + 16])
        if t == MDB_NUMERIC:
            return f"[NUMERIC:{size}B:{page[start:start + size].hex()}]"
        return f"[type {t}:{size}B]"

    @staticmethod
    def _decode_datetime(td):
        try:
            if td is None or td < 0 or td > 1e6:
                return None
            base = datetime(1899, 12, 30)
            return base + timedelta(days=td)
        except Exception:
            return None

    @staticmethod
    def _format_guid(b):
        if len(b) < 16:
            return b.hex().upper()
        d1, d2, d3 = struct.unpack_from("<IHH", b, 0)
        rest = b[8:16].hex().upper()
        return f"{{{d1:08X}-{d2:04X}-{d3:04X}-{rest[:4]}-{rest[4:]}}}"

    def _decode_memo(self, page, start, size):
        if size < 12:
            return None
        memo_len = i32(page, start)
        if memo_len & 0x80000000:
            raw = page[start + 12:start + size]
            return ("__text__", raw)
        pg_row = i32(page, start + 4)
        if memo_len & 0x40000000:
            lval_page, row_start, row_len = self.find_pg_row(pg_row)
            raw = lval_page[row_start:row_start + row_len]
            return ("__text__", raw)
        total = memo_len & 0x00FFFFFF
        chunks = bytearray()
        seen = set()
        while pg_row and pg_row not in seen and len(chunks) < total + 8:
            seen.add(pg_row)
            lval_page, row_start, row_len = self.find_pg_row(pg_row)
            if row_len < 4:
                break
            chunks += lval_page[row_start + 4:row_start + row_len]
            pg_row = i32(lval_page, row_start)
        return ("__text__", bytes(chunks[:total]))

    @staticmethod
    def _decode_mixed_segments(body, be=False):
        parts = []
        for seg in body.split(b"\x00"):
            if not seg:
                continue
            if all(b < 0x80 for b in seg):
                parts.append(seg.decode("ascii", errors="replace"))
            elif len(seg) % 2 == 0:
                parts.append(seg.decode("utf-16-be" if be else "utf-16-le", errors="replace"))
            else:
                parts.append(seg.decode("cp950", errors="replace"))
        return "".join(parts)

    @classmethod
    def decode_text_value(cls, raw):
        if not raw:
            return None
        if raw.startswith(b"\xff\xfe"):
            s = cls._decode_mixed_segments(raw[2:])
        elif raw.startswith(b"\xfe\xff"):
            s = cls._decode_mixed_segments(raw[2:], be=True)
        elif len(raw) % 2 == 1:
            s = raw.decode("cp950", errors="replace")
        else:
            s = raw.decode("utf-16-le", errors="replace")
            if "�" in s:
                s2 = cls._decode_mixed_segments(raw)
                s = s2 if "�" not in s2 else raw.decode("cp950", errors="replace")
        return s.replace("﻿", "").rstrip("\x00")

    def resolve_text_columns(self, table, rows):
        for idx, col in enumerate(table.columns):
            if col.col_type not in (MDB_TEXT, MDB_MEMO):
                continue
            for row in rows:
                v = row[idx]
                if isinstance(v, tuple) and len(v) == 2 and v[0] == "__text__":
                    row[idx] = self.decode_text_value(v[1])
        return rows

    def read_table_data(self, entry):
        table = self.read_table(entry)
        self.read_columns(table)
        rows = self.fetch_rows(table)
        self.resolve_text_columns(table, rows)
        return table, rows


def load_wex_wmp(path):
    import py7zr
    with py7zr.SevenZipFile(path, "r") as z:
        names = z.getnames()
        target = next((n for n in names if n.upper().endswith("EDIT.WMP")), None)
        if target is None:
            target = next((n for n in names if n.upper().endswith(".WMP")), None)
        if target is None:
            raise RuntimeError("壓縮檔內找不到 EDIT.WMP")
        tmpdir = tempfile.mkdtemp(prefix="wex_")
        try:
            z.extract(tmpdir, targets=[target])
            extracted = os.path.join(tmpdir, target.replace("/", os.sep))
            with open(extracted, "rb") as f:
                return f.read()
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


def open_wex_database(path):
    if hasattr(path, "read"):
        return JetDatabase(path)
    upper = path.upper()
    if upper.endswith(".WEX") or upper.endswith(".7Z"):
        return JetDatabase(io.BytesIO(load_wex_wmp(path)))
    return JetDatabase(path)


def format_value(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, datetime):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return f"{v:.6g}"
    s = str(v).replace("﻿", "")
    return "".join(ch for ch in s if ch == "\t" or ch == "\n" or ch == "\r" or ord(ch) >= 32)


def dump_database(path, max_rows=8):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    db = open_wex_database(path)
    print(f"檔案: {path}")
    print(f"頁面數: {db.num_pages}, code_page={db.code_page}, db_key=0x{db.db_key:08X}")
    catalog = db.read_catalog()
    user_tables = [e for e in catalog if e.obj_type == 1 and not (e.flags & 0x80000002)]
    system_tables = [e for e in catalog if e.obj_type == 1 and (e.flags & 0x80000002)]
    user_tables.sort(key=lambda e: e.name.lower())
    system_tables.sort(key=lambda e: e.name.lower())
    print(f"資料表: 使用者 {len(user_tables)} 個, 系統 {len(system_tables)} 個")
    for entry in user_tables + system_tables:
        try:
            table, rows = db.read_table_data(entry)
        except Exception as e:
            print(f"\n=== {entry.name} (讀取失敗: {e}) ===")
            continue
        kind = "系統" if (entry.flags & 0x80000002) else "使用者"
        print(f"\n=== {entry.name} [{kind}] {len(rows)} 列 x {len(table.columns)} 欄 ===")
        print("欄位: " + " | ".join(c.name for c in table.columns))
        for row in rows[:max_rows]:
            print("  " + " | ".join(format_value(v) for v in row))
        if len(rows) > max_rows:
            print(f"  ... (共 {len(rows)} 列)")


class SheetApp:
    def __init__(self, root, initial_path=None):
        import tkinter as tk
        from tkinter import ttk, messagebox, filedialog
        from tksheet import Sheet

        self.root = root
        self.tk = tk
        self.ttk = ttk
        self.messagebox = messagebox
        self.filedialog = filedialog
        self.Sheet = Sheet
        self.db = None
        self.current_path = None

        root.title("廣播時程檢視器")
        root.geometry("1280x760")
        root.minsize(800, 500)

        menubar = tk.Menu(root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="開啟檔案 (.WEX / .WMP) ...", command=self.open_dialog, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="結束", command=root.destroy)
        menubar.add_cascade(label="檔案", menu=file_menu)
        root.config(menu=menubar)
        root.bind("<Control-o>", lambda e: self.open_dialog())

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=4, pady=4)

        self.status = ttk.Label(root, anchor="w", relief="sunken", text="請開啟 .WEX 或 .WMP 檔案")
        self.status.pack(fill="x", side="bottom")

        self.empty_label = ttk.Label(self.notebook, text="尚未載入資料\n\n檔案 → 開啟檔案，選擇廣播系統的 .WEX 時程檔", justify="center")
        self.notebook.add(self.empty_label, text="歡迎")

        if initial_path:
            root.after(100, lambda: self.load_path(initial_path))

    def open_dialog(self):
        path = self.filedialog.askopenfilename(
            title="選擇廣播時程檔",
            filetypes=[
                ("廣播時程檔 (*.WEX)", "*.WEX"),
                ("資料庫 (*.WMP;*.MDB)", "*.WMP *.mdb *.MDB"),
                ("所有檔案", "*.*"),
            ],
        )
        if path:
            self.load_path(path)

    def load_path(self, path):
        try:
            self.status.config(text=f"載入中: {path}")
            self.root.update_idletasks()
            db = open_wex_database(path)
            catalog = db.read_catalog()
            tables = [e for e in catalog if e.obj_type == 1 and not (e.flags & 0x80000002)]
            if not tables:
                tables = [e for e in catalog if e.obj_type == 1]
            tables.sort(key=lambda e: e.name.lower())
            self.db = db
            self.current_path = path
            for tab in self.notebook.winfo_children():
                self.notebook.forget(tab)
                tab.destroy()
            loaded = 0
            for entry in tables:
                try:
                    table, rows = db.read_table_data(entry)
                except Exception:
                    continue
                headers = [c.name for c in table.columns]
                data = [[format_value(v) for v in row] for row in rows]
                frame = self.ttk.Frame(self.notebook)
                sheet = self.Sheet(
                    frame,
                    data=data if data else [["" for _ in headers]],
                    headers=headers if headers else ["(無欄位)"],
                    width=1180,
                    height=680,
                    header_font=("Microsoft JhengHei UI", 10, "bold"),
                    table_font=("Microsoft JhengHei UI", 10, "normal"),
                )
                sheet.enable_bindings()
                sheet.readonly(True)
                sheet.set_all_column_widths(110)
                sheet.pack(fill="both", expand=True)
                self.notebook.add(frame, text=f"{entry.name} ({len(rows)})")
                loaded += 1
            self.root.title(f"廣播時程檢視器 - {os.path.basename(path)}")
            self.status.config(
                text=f"{path}  |  已載入 {loaded} 個資料表  |  code_page={db.code_page}"
            )
        except Exception:
            import traceback
            tb = traceback.format_exc()
            try:
                log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "wex_viewer_error.log")
                with open(log_path, "w", encoding="utf-8") as f:
                    f.write(tb)
                tb += "\n\n(已寫入 " + log_path + ")"
            except Exception:
                pass
            self.messagebox.showerror("載入失敗", tb[-1500:])
            self.status.config(text="載入失敗")


def selftest(path):
    import tkinter as tk
    root = tk.Tk()
    app = SheetApp(root, initial_path=path)
    root.after(3000, root.destroy)
    root.mainloop()
    print("selftest ok")


def find_default_wex():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    skip_dirs = {"WEX", "__pycache__"}
    search_dirs = [script_dir]
    try:
        search_dirs += sorted(
            os.path.join(script_dir, e) for e in os.listdir(script_dir)
            if os.path.isdir(os.path.join(script_dir, e)) and not e.startswith("__")
            and e not in skip_dirs
        )
    except OSError:
        pass
    cwd = os.path.abspath(".")
    if cwd not in search_dirs:
        search_dirs.append(cwd)
    for d in search_dirs:
        try:
            cands = sorted(p for p in os.listdir(d) if p.upper().endswith(".WEX"))
        except OSError:
            continue
        if cands:
            return os.path.join(d, cands[0])
    return None


def main():
    parser = argparse.ArgumentParser(description="廣播時程 .WEX 檔案檢視器")
    parser.add_argument("path", nargs="?", help=".WEX 或 .WMP 檔案路徑")
    parser.add_argument("--dump", action="store_true", help="以文字列出資料表內容 (不開視窗)")
    parser.add_argument("--selftest", action="store_true", help="開啟視窗 3 秒後自動關閉")
    args = parser.parse_args()
    if args.dump and not args.path:
        print("請提供檔案路徑，例如: python wex_viewer.py --dump \"xxx.WEX\"")
        return
    if args.dump:
        dump_database(args.path)
        return
    if args.selftest:
        selftest(args.path)
        return
    path = args.path or find_default_wex()
    import tkinter as tk
    root = tk.Tk()
    SheetApp(root, initial_path=path)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        traceback.print_exc()
        print()
        try:
            input("發生錯誤，按 Enter 結束... ")
        except Exception:
            pass
