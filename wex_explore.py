"""欄位語意探勘工具：用來驗證 MUSIC*/SETTING/AREA/TGROUP 等欄位與畫面的映射關係。

用法：
  python wex_explore.py [檔案.WEX|WMP]
  python wex_explore.py 檔案.WEX --table MUSIC1
"""
import argparse
import io
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from wex_viewer import open_wex_database, format_value, JetDatabase


def load(path):
    db = open_wex_database(path)
    catalog = db.read_catalog()
    tables = {}
    for e in catalog:
        if e.obj_type != 1:
            continue
        try:
            tables[e.name] = db.read_table_data(e)
        except Exception as ex:
            print(f"[warn] {e.name} 讀取失敗: {ex}")
    return db, tables


def rows_as_dicts(table, rows):
    cols = [c.name for c in table.columns]
    return [dict(zip(cols, r)) for r in rows]


def show_setting(tables):
    if "SETTING" not in tables:
        return
    table, rows = tables["SETTING"]
    print("=== SETTING ===")
    for r in rows_as_dicts(table, rows):
        for k, v in r.items():
            if v in (None, "", False):
                continue
            print(f"  {k} = {format_value(v)!r}")


def show_pages(tables):
    for name in ("MUSIC", "PAGE"):
        if name not in tables:
            continue
        table, rows = tables[name]
        print(f"=== {name} ===")
        for r in rows_as_dicts(table, rows):
            print("  " + " | ".join(f"{k}={format_value(v)}" for k, v in r.items() if v not in (None, "", False)))


def show_speed(tables):
    for name in sorted(tables):
        if not name.startswith("SPEED"):
            continue
        table, rows = tables[name]
        print(f"=== {name} ===")
        for r in rows_as_dicts(table, rows):
            print("  " + " | ".join(f"{k}={format_value(v)}" for k, v in r.items() if v not in (None, "", False)))


def show_tgroup(tables):
    if "TGROUP" not in tables:
        return
    table, rows = tables["TGROUP"]
    print("=== TGROUP ===")
    cols = [c.name for c in table.columns]
    print("欄位: " + " | ".join(cols))
    for r in rows_as_dicts(table, rows):
        print("  " + " | ".join(f"{format_value(v)}" for v in r.values()))


def show_music_tables(tables):
    for i in list(range(1, 9)) + ["_T", ""]:
        name = "MUSIC" + str(i) if i != "" else "MUSIC"
        if name not in tables:
            continue
        table, rows = tables[name]
        if not rows:
            print(f"=== {name} (0 列) ===")
            continue
        drows = rows_as_dicts(table, rows)
        print(f"=== {name}: {len(drows)} 列 ===")
        interesting = [
            "Mode", "Music_on", "Allcall", "Channel", "Multi_mode", "Vol",
            "Group1", "Group2", "Group3", "Week_day", "Dir_OK", "Rhythm",
            "Speed", "Pitch", "Modulation", "Speaker",
        ]
        for key in interesting:
            if key not in drows[0]:
                continue
            cnt = Counter(format_value(r.get(key)) for r in drows)
            if len(cnt) <= 12:
                print(f"  {key}: {dict(cnt)}")
            else:
                print(f"  {key}: {len(cnt)} 種值, top5={cnt.most_common(5)}")
        for key in ("Clock_year", "Clock_month", "Clock_day"):
            if key in drows[0]:
                cnt = Counter(format_value(r.get(key)) for r in drows)
                print(f"  {key}: {len(cnt)} 種, top3={cnt.most_common(3)}")
        # 時間欄位比較：Clock_* vs Start_*/End_*
        if "Start_hour" in drows[0]:
            diff = 0
            same = 0
            for r in drows:
                cs = (format_value(r.get("Clock_start_hour")), format_value(r.get("Clock_start_min")))
                ss = (format_value(r.get("Start_hour")), format_value(r.get("Start_min")))
                if cs == ss:
                    same += 1
                else:
                    diff += 1
            print(f"  Clock_start vs Start_hour/min: 相同 {same} / 不同 {diff}")
        # 播放日期是否有意義
        if "Clock_year" in drows[0]:
            with_date = sum(1 for r in drows if format_value(r.get("Clock_year")) not in ("", "2015"))
            print(f"  Clock_year ≠ 2015 的列數: {with_date}")
        # Dir 陣列
        dir_cols = [c for c in drows[0] if c.startswith("Dir") and c[3:].isdigit()]
        dir_nonzero = Counter()
        for r in drows:
            used = [c for c in dir_cols if r.get(c) not in (None, "")]
            dir_nonzero[len(used)] += 1
        print(f"  Dir 陣列非空個數分佈: {dict(sorted(dir_nonzero.items()))}")
        # Group 值域
        gcols = [c for c in drows[0] if c.startswith("Group") and c[5:].isdigit()]
        gvals = set()
        for r in drows:
            for c in gcols:
                v = r.get(c)
                if v not in (None, "", False, 0):
                    gvals.add(v)
        print(f"  Group 使用值: {sorted(gvals, key=lambda x: (isinstance(x, str), x))[:40]}")
        # 範例列
        show_cols = [c for c in (
            "Music_no", "Music_name", "Mode", "Music_on", "Allcall",
            "Week1", "Week2", "Week3", "Week4", "Week5", "Week6", "Week7",
            "Clock_start_hour", "Clock_start_min", "Clock_start_sec",
            "Clock_end_hour", "Clock_end_min", "Clock_end_sec",
            "Start_hour", "Start_min", "Start_sec", "End_hour", "End_min", "End_sec",
            "Time", "Total_time", "Vol", "Group1", "Group2", "Group3",
            "Dir0", "Dir1", "Text0", "Multi_mode",
        ) if c in drows[0]]
        print("  範例列: " + " | ".join(show_cols))
        for r in drows[:3]:
            print("    " + " | ".join(format_value(r.get(c)) for c in show_cols))
        print()


def show_area_summary(tables):
    if "AREA" not in tables:
        return
    table, rows = tables["AREA"]
    drows = rows_as_dicts(table, rows)
    print(f"=== AREA: {len(drows)} 列 ===")
    gcols = [c for c in drows[0] if c.startswith("Group") and c[5:].isdigit()]
    # 每個 Group 欄位代表的群組成員
    usage = Counter()
    for r in drows:
        for c in gcols:
            if r.get(c) is True:
                usage[c] += 1
    print(f"  Group 欄位使用情況: {dict(sorted(usage.items(), key=lambda x: int(x[0][5:])))}")
    # 前幾列的 Area_name / Number
    print("  範例: " + " | ".join(
        f"{format_value(r.get('Area_name'))}(Num={format_value(r.get('Number'))},Decoder={format_value(r.get('Decoder'))},"
        f"Groups={[c for c in gcols if r.get(c) is True]})"
        for r in drows[:5]
    ))
    # Function/Channel
    fcnt = Counter(format_value(r.get("Function_on")) for r in drows)
    print(f"  Function_on: {dict(fcnt)}")
    print()


def show_raw_dir(path):
    """檢視 Dir 欄位的原始 bytes 編碼方式。"""
    db = open_wex_database(path)
    catalog = db.read_catalog()
    entry = next((e for e in catalog if e.name.upper() == "MUSIC1"), None)
    if entry is None:
        return
    table = db.read_table(entry)
    db.read_columns(table)
    rows = db.fetch_rows(table)  # 不 resolve，保留原始 __text__
    dir_cols = [i for i, c in enumerate(table.columns) if c.name.startswith("Dir") and c.name[3:].isdigit()]
    print("=== Dir 欄位原始編碼抽樣 (MUSIC1) ===")
    shown = 0
    for row in rows:
        for i in dir_cols:
            v = row[i]
            if isinstance(v, tuple) and v[1]:
                raw = v[1]
                print(f"  {table.columns[i].name}: len={len(raw)} hex={raw[:24].hex()} repr={raw[:40]!r}")
                shown += 1
        if shown > 15:
            break


def find_default_wex():
    d = os.path.dirname(os.path.abspath(__file__))
    cwd = os.path.abspath(".")
    for search in (d, cwd):
        try:
            cands = sorted(p for p in os.listdir(search) if p.upper().endswith(".WEX"))
        except OSError:
            continue
        if cands:
            return os.path.join(search, cands[0])
    return None


def main():
    ap = argparse.ArgumentParser(description="WEX 欄位語意探勘")
    ap.add_argument("path", nargs="?", help=".WEX 或 .WMP 檔案路徑（省略則自動找同資料夾第一個 .WEX）")
    ap.add_argument("--table", help="只印某資料表全部列")
    args = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if not args.path:
        args.path = find_default_wex()
        if not args.path:
            print("找不到 .WEX 檔案，請指定路徑：python wex_explore.py 檔案.WEX")
            return

    if args.table:
        db, tables = load(args.path)
        if args.table not in tables:
            print(f"資料表 {args.table} 不存在，可使用：{sorted(tables)}")
            return
        table, rows = tables[args.table]
        drows = rows_as_dicts(table, rows)
        cols = [c.name for c in table.columns]
        print(" | ".join(cols))
        for r in drows:
            print(" | ".join(format_value(r[c]) for c in cols))
        return

    db, tables = load(args.path)
    print(f"檔案: {args.path}")
    print(f"資料表: {sorted(tables)}")
    print()
    show_setting(tables)
    print()
    show_pages(tables)
    print()
    show_speed(tables)
    print()
    show_tgroup(tables)
    print()
    show_music_tables(tables)
    print()
    show_area_summary(tables)
    show_raw_dir(args.path)


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
    else:
        if len(sys.argv) == 1:
            try:
                input("\n完成，按 Enter 結束... ")
            except Exception:
                pass
