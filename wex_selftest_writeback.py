"""前置驗證：py7zr Copy 重打包往返 + ACE ODBC 密碼讀寫。"""
import io
import os
import shutil
import struct
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import py7zr
import pyodbc

SRC = "1151014-16-全校第一次定期考.WEX"
PWD = "wmpedit"


def test_7z_roundtrip():
    print("=== py7zr Copy 重打包往返 ===")
    with py7zr.SevenZipFile(SRC, "r") as z:
        names = z.getnames()
        print("原始成員:", names)
        tmp1 = tempfile.mkdtemp(prefix="wex_rt1_")
        z.extract(tmp1)
    # 重打包（copy filter, 不壓縮）
    out_buf = io.BytesIO()
    with py7zr.SevenZipFile(out_buf, "w", filters=[{"id": py7zr.FILTER_COPY}]) as z:
        for entry in sorted(os.listdir(tmp1)):
            full = os.path.join(tmp1, entry)
            if os.path.isdir(full):
                z.writeall(full, arcname=entry)
            else:
                z.write(full, arcname=entry)
    data = out_buf.getvalue()
    print(f"原始 WEX 大小: {os.path.getsize(SRC)}, 重打包大小: {len(data)}")
    # 重新開啟比較成員內容
    with py7zr.SevenZipFile(io.BytesIO(data), "r") as z2:
        names2 = z2.getnames()
        print("重打包成員:", names2)
        tmp2 = tempfile.mkdtemp(prefix="wex_rt2_")
        z2.extract(tmp2)
    # 逐檔比較
    ok = True
    for root, dirs, files in os.walk(tmp1):
        for f in files:
            p1 = os.path.join(root, f)
            rel = os.path.relpath(p1, tmp1)
            p2 = os.path.join(tmp2, rel)
            if not os.path.exists(p2):
                print(f"  缺少: {rel}")
                ok = False
                continue
            b1 = open(p1, "rb").read()
            b2 = open(p2, "rb").read()
            same = b1 == b2
            if not same:
                ok = False
            print(f"  {rel}: {'相同' if same else '不同!'} ({len(b1)} vs {len(b2)})")
    # 空目錄保留?
    d1 = {os.path.relpath(r, tmp1) for r, _, _ in os.walk(tmp1)}
    d2 = {os.path.relpath(r, tmp2) for r, _, _ in os.walk(tmp2)}
    if d1 != d2:
        print(f"  目錄差異: 只在原={d1 - d2}, 只在新={d2 - d1}")
    else:
        print("  目錄結構相同")
    shutil.rmtree(tmp1, ignore_errors=True)
    shutil.rmtree(tmp2, ignore_errors=True)
    return ok


def test_ace():
    print("\n=== ACE ODBC 密碼讀寫 ===")
    tmpdir = tempfile.mkdtemp(prefix="wex_ace_")
    with py7zr.SevenZipFile(SRC, "r") as z:
        target = next(n for n in z.getnames() if n.upper().endswith("EDIT.WMP"))
        z.extract(tmpdir, targets=[target])
    wmp = os.path.join(tmpdir, target.replace("/", os.sep))
    print("WMP:", wmp, os.path.getsize(wmp), "bytes")
    # 嘗試不帶密碼 → 應失敗
    try:
        conn = pyodbc.connect(
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + wmp + ";",
            autocommit=True,
        )
        conn.close()
        print("  無密碼開啟: 意外成功!")
    except Exception as e:
        print(f"  無密碼開啟: 預期失敗 → {str(e)[:80]}")
    # 帶密碼
    conn = pyodbc.connect(
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + wmp + f";PWD={PWD};",
        autocommit=False,
    )
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM SETTING")
    print("  SETTING 列數:", cur.fetchone()[0])
    cur.execute("SELECT Vol FROM SETTING")
    print("  原 Vol:", cur.fetchone()[0])
    cur.execute("UPDATE SETTING SET Vol=91")
    conn.commit()
    conn.close()
    # 重開驗證 + 頁面差異
    orig = load_wmp_bytes(SRC)
    new = open(wmp, "rb").read()
    pg = 4096
    ndiff = sum(1 for i in range(len(orig) // pg) if orig[i * pg:(i + 1) * pg] != new[i * pg:(i + 1) * pg])
    print(f"  UPDATE 後檔案大小: {len(new)} (原 {len(orig)}), 差異頁數: {ndiff}")
    conn = pyodbc.connect(
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + wmp + f";PWD={PWD};",
        autocommit=True,
    )
    cur = conn.cursor()
    cur.execute("SELECT Vol FROM SETTING")
    print("  讀回 Vol:", cur.fetchone()[0])
    # 測試文字欄位 (MUSIC3 名稱)
    cur.execute("SELECT Music_no, Music_name FROM MUSIC3 WHERE Music_no=1")
    row = cur.fetchone()
    print(f"  MUSIC3 Music_no=1: {row[0]} / {row[1]}")
    cur.execute("UPDATE MUSIC3 SET Clock_start_hour='08', Clock_start_min='01' WHERE Music_no=1 AND Music_name='打掃'")
    print("  UPDATE 影響列數:", cur.rowcount)
    conn.commit()
    cur.execute("SELECT Clock_start_hour, Clock_start_min FROM MUSIC3 WHERE Music_no=1 AND Music_name='打掃'")
    print("  讀回時間:", cur.fetchone()[:])
    conn.close()
    shutil.rmtree(tmpdir, ignore_errors=True)


def load_wmp_bytes(path):
    with py7zr.SevenZipFile(path, "r") as z:
        target = next(n for n in z.getnames() if n.upper().endswith("EDIT.WMP"))
        tmpdir = tempfile.mkdtemp(prefix="wex_rd_")
        try:
            z.extract(tmpdir, targets=[target])
            with open(os.path.join(tmpdir, target.replace("/", os.sep)), "rb") as f:
                return f.read()
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ok = test_7z_roundtrip()
    print()
    test_ace()
