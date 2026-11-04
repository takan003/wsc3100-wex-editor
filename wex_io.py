"""WEX 檔案讀寫核心：7z 容器、ACE 密碼寫回、時程資料模型。

用法：
  from wex_io import WexProject
  proj = WexProject('xxx.WEX')
  proj.tabs            # [(id, name, [row_dict, ...]), ...]
  proj.export('out.WEX', statements)
"""
import io
import math
import os
import re
import shutil
import struct
import tempfile
import wave
from datetime import datetime

import py7zr
import pyodbc

from wex_viewer import open_wex_database, format_value

DB_PASSWORD = "wmpedit"
NO_DATE = (2015, 3, 14)
ACE_CONN = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ={path};PWD=" + DB_PASSWORD
WEEK_LABELS = ["一", "二", "三", "四", "五", "六", "日"]
MODE_LABELS = {1: "單曲"}


def load_wex_members(path):
    members = {}
    dir_entries = []
    with py7zr.SevenZipFile(path, "r") as z:
        for name in z.getnames():
            norm = name.replace("\\", "/").rstrip("/")
            if not norm:
                continue
            if norm.endswith("/"):
                dir_entries.append(norm.rstrip("/"))
    with py7zr.SevenZipFile(path, "r") as z:
        tmp = tempfile.mkdtemp(prefix="wex_load_")
        try:
            z.extract(tmp)
            for root, dirs, files in os.walk(tmp):
                for d in dirs:
                    rel = os.path.relpath(os.path.join(root, d), tmp).replace("\\", "/")
                    members[rel + "/"] = None
                for f in files:
                    rel = os.path.relpath(os.path.join(root, f), tmp).replace("\\", "/")
                    with open(os.path.join(root, f), "rb") as fp:
                        members[rel] = fp.read()
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    return members


def save_wex_members(path, members):
    tmp = tempfile.mkdtemp(prefix="wex_save_")
    try:
        dirs = set()
        for name, data in members.items():
            if data is None:
                dirs.add(name.rstrip("/"))
                continue
            full = os.path.join(tmp, name.replace("/", os.sep))
            os.makedirs(os.path.dirname(full), exist_ok=True)
            with open(full, "wb") as f:
                f.write(data)
            parts = name.split("/")
            for i in range(1, len(parts)):
                dirs.add("/".join(parts[:i]))
        for d in sorted(dirs):
            os.makedirs(os.path.join(tmp, d.replace("/", os.sep)), exist_ok=True)
        with py7zr.SevenZipFile(path, "w", filters=[{"id": py7zr.FILTER_COPY}]) as z:
            for entry in sorted(os.listdir(tmp)):
                full = os.path.join(tmp, entry)
                if os.path.isdir(full):
                    z.writeall(full, arcname=entry)
                else:
                    z.write(full, arcname=entry)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


class AceDatabase:
    def __init__(self, wmp_path):
        self.path = wmp_path
        self.conn = None

    def __enter__(self):
        self.conn = pyodbc.connect(ACE_CONN.replace("{path}", self.path), autocommit=False)
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.conn is not None:
            if exc_type is None:
                self.conn.commit()
            else:
                self.conn.rollback()
            self.conn.close()
            self.conn = None
        return False

    def execute(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        return cur

    def query(self, sql, params=()):
        cur = self.execute(sql, params)
        return cur.fetchall()


def apply_sql_to_wmp(wmp_bytes, statements):
    tmp = tempfile.mkdtemp(prefix="wex_sql_")
    try:
        wmp_path = os.path.join(tmp, "EDIT.WMP")
        with open(wmp_path, "wb") as f:
            f.write(wmp_bytes)
        with AceDatabase(wmp_path) as db:
            for stmt in statements:
                if isinstance(stmt, tuple):
                    sql, params = stmt
                    db.execute(sql, params)
                else:
                    db.execute(stmt)
        with open(wmp_path, "rb") as f:
            return f.read()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def wav_duration(data):
    with wave.open(io.BytesIO(data)) as w:
        return w.getnframes() / float(w.getframerate())


def mp3_duration(data):
    if data[:3] == b"ID3":
        sz = ((data[6] & 0x7F) << 21) | ((data[7] & 0x7F) << 14) | ((data[8] & 0x7F) << 7) | (data[9] & 0x7F)
        data = data[10 + sz:]
    br_table = [0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 0]
    for i in range(len(data) - 4):
        if data[i] == 0xFF and (data[i + 1] & 0xE0) == 0xE0:
            ver = (data[i + 1] >> 3) & 3
            layer = (data[i + 1] >> 1) & 3
            bri = (data[i + 2] >> 4) & 0xF
            sri = (data[i + 2] >> 2) & 3
            if ver == 3 and layer == 1 and bri not in (0, 15) and sri != 3:
                sr = [44100, 48000, 32000][sri]
                kbps = br_table[bri]
                side = 32 if ((data[i + 3] >> 6) & 1) == 0 else 17
                xing_at = i + 4 + side
                if data[xing_at:xing_at + 4] in (b"Xing", b"Info"):
                    flags = struct.unpack_from(">I", data, xing_at + 4)[0]
                    if flags & 1:
                        frames = struct.unpack_from(">I", data, xing_at + 8)[0]
                        return frames * 1152 / float(sr)
                return len(data) * 8.0 / (kbps * 1000)
    return None


def audio_duration(data):
    if data[:4] == b"RIFF" and data[8:12] == b"WAVE":
        return wav_duration(data)
    return mp3_duration(data)


def duration_to_clock_text(seconds):
    seconds = int(math.ceil(seconds)) if seconds is not None else 0
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


class ScheduleRow:
    __slots__ = ("table", "data")

    def __init__(self, table, data):
        self.table = table
        self.data = data

    def get(self, col, default=None):
        v = self.data.get(col)
        return default if v is None else v

    @property
    def music_id(self):
        return int(self.data.get("MUSIC") or 0)

    @property
    def music_no(self):
        return int(self.data.get("Music_no") or 0)

    @property
    def name(self):
        return str(self.data.get("Music_name") or "")

    @property
    def weeks(self):
        return [bool(self.data.get(f"Week{i}")) for i in range(1, 8)]

    @property
    def start_time(self):
        h = int(self.data.get("Clock_start_hour") or 0)
        m = int(self.data.get("Clock_start_min") or 0)
        return h, m

    @property
    def date(self):
        y = int(self.data.get("Clock_year") or 0)
        mo = int(self.data.get("Clock_month") or 0)
        d = int(self.data.get("Clock_day") or 0)
        if (y, mo, d) == NO_DATE or y == 0:
            return None
        return y, mo, d

    @property
    def duration(self):
        return int(self.data.get("Total_time") or 0)

    @property
    def dir0(self):
        return str(self.data.get("Dir0") or "")

    @property
    def groups(self):
        return [int(self.data.get(f"Group{i}") or 0) for i in range(1, 4)]

    @property
    def allcall(self):
        return bool(self.data.get("Allcall"))

    @property
    def volume(self):
        v = self.data.get("Vol")
        return int(v) if v is not None else None

    @property
    def mode(self):
        return int(self.data.get("Mode") or 1)

    @property
    def music_on(self):
        v = self.data.get("Music_on")
        return bool(v) if v is not None else True


class WexProject:
    def __init__(self, path):
        self.path = path
        self.members = load_wex_members(path)
        wmp_key = next(k for k in self.members if k.upper().endswith("EDIT.WMP"))
        self.wmp_key = wmp_key
        self.wmp_bytes = self.members[wmp_key]
        self.db = open_wex_database(io.BytesIO(self.wmp_bytes))
        catalog = self.db.read_catalog()
        self._catalog = {e.name: e for e in catalog if e.obj_type == 1}
        self.tab_names = self._load_tab_names()
        self.group_names = self._load_group_names()
        self.area_names = self._load_area_names()
        self.audio_files = sorted(
            name.split("/", 1)[1] for name in self.members
            if name.startswith("Normal/") and not name.endswith("/")
        )
        self.audio_durations = {}
        for name in self.audio_files:
            raw = self.members.get("Normal/" + name)
            if raw:
                try:
                    self.audio_durations[name] = audio_duration(raw)
                except Exception:
                    pass
        self.schedule_ids = sorted(
            int(m.group(1)) for name in self._catalog
            if (m := re.fullmatch(r"MUSIC(\d+)", name))
        )

    def _load_tab_names(self):
        names = {}
        if "MUSIC" in self._catalog:
            t, rows = self.db.read_table_data(self._catalog["MUSIC"])
            cols = [c.name for c in t.columns]
            for row in rows:
                d = dict(zip(cols, row))
                names[int(d.get("MUSIC") or 0)] = str(d.get("Music_name") or "")
        return names

    def _load_group_names(self):
        names = {}
        if "TGROUP" in self._catalog:
            t, rows = self.db.read_table_data(self._catalog["TGROUP"])
            cols = [c.name for c in t.columns]
            for row in rows:
                d = dict(zip(cols, row))
                gid = int(d.get("TGROUP") or 0)
                gname = str(d.get("Group_name") or "")
                names[gid] = "" if gname.startswith("Group") and gname[5:].isdigit() else gname
        return names

    def _load_area_names(self):
        names = {}
        if "AREA" in self._catalog:
            t, rows = self.db.read_table_data(self._catalog["AREA"])
            cols = [c.name for c in t.columns]
            for row in rows:
                d = dict(zip(cols, row))
                dec = int(d.get("Decoder") or d.get("AREA") or 0)
                names[dec] = str(d.get("Area_name") or "")
        return names

    def load_rows(self, table_name):
        entry = self._catalog.get(table_name)
        if entry is None:
            return []
        t, rows = self.db.read_table_data(entry)
        cols = [c.name for c in t.columns]
        out = []
        for row in rows:
            d = dict(zip(cols, row))
            out.append(ScheduleRow(table_name, d))
        out.sort(key=lambda r: (r.music_no, r.music_id))
        return out

    @property
    def tabs(self):
        ids = sorted(set(self.schedule_ids) | set(self.tab_names))
        result = []
        for i in ids:
            name = self.tab_names.get(i, "")
            rows = self.load_rows(f"MUSIC{i}") if f"MUSIC{i}" in self._catalog else []
            if not name and not rows:
                continue
            result.append((i, name or f"時程{i}", rows))
        return result

    def group_label(self, gid):
        if not gid:
            return ""
        name = self.group_names.get(gid, "")
        return name or f"群組{gid}"

    def area_members(self, gid):
        entry = self._catalog.get("TGROUP")
        if entry is None:
            return []
        t, rows = self.db.read_table_data(entry)
        cols = [c.name for c in t.columns]
        for row in rows:
            d = dict(zip(cols, row))
            if int(d.get("TGROUP") or 0) == gid:
                inc = str(d.get("IncludeArea") or "")
                return [self.area_names[i] for i, ch in enumerate(inc) if ch == "1" and i in self.area_names]
        return []

    def clone_row(self, table_name, music_id):
        for row in self.load_rows(table_name):
            if row.music_id == music_id:
                return dict(row.data)
        raise KeyError(f"{table_name} 中找不到 MUSIC={music_id}")

    def next_music_id(self, table_name):
        rows = self.load_rows(table_name)
        return max((r.music_id for r in rows), default=0) + 1

    def next_music_no(self, table_name):
        rows = self.load_rows(table_name)
        return max((r.music_no for r in rows), default=0) + 1


_TEXT_ESCAPES = {"\x00": "", "\r": "", "\n": ""}


def sql_quote(value):
    if value is None:
        return "NULL"
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, (int, float)):
        return str(value)
    s = str(value)
    for k, v in _TEXT_ESCAPES.items():
        s = s.replace(k, v)
    return "'" + s.replace("'", "''") + "'"


def build_update_row(table_name, music_id, changes):
    sets = ", ".join(f"[{col}] = {sql_quote(val)}" for col, val in changes.items())
    return f"UPDATE [{table_name}] SET {sets} WHERE [MUSIC] = {int(music_id)}"


def build_delete_row(table_name, music_id):
    return f"DELETE FROM [{table_name}] WHERE [MUSIC] = {int(music_id)}"


def build_insert_row(table_name, row_data):
    cols = list(row_data.keys())
    col_sql = ", ".join(f"[{c}]" for c in cols)
    val_sql = ", ".join(sql_quote(row_data[c]) for c in cols)
    return f"INSERT INTO [{table_name}] ({col_sql}) VALUES ({val_sql})"


def row_duration_changes(row_data, audio_durations):
    changes = {}
    dir0 = str(row_data.get("Dir0") or "")
    dur = audio_durations.get(dir0)
    if dur is None:
        return changes
    secs = int(math.ceil(dur))
    h, rem = divmod(secs, 3600)
    m, s = divmod(rem, 60)
    changes["Time"] = secs
    changes["Total_time"] = secs
    changes["End_hour"] = f"{h:02d}"
    changes["End_min"] = f"{m:02d}"
    changes["End_sec"] = f"{s:02d}"
    return changes


def export_wex(src_path, dst_path, statements, wmp_transform=None):
    members = load_wex_members(src_path)
    wmp_key = next(k for k in members if k.upper().endswith("EDIT.WMP"))
    wmp = members[wmp_key]
    if statements:
        wmp = apply_sql_to_wmp(wmp, statements)
    if wmp_transform is not None:
        wmp = wmp_transform(wmp)
    members[wmp_key] = wmp
    save_wex_members(dst_path, members)
    return dst_path
