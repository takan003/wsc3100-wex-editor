"""產生 Excel 時程範本檔。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HEADERS = ["序號", "名稱註記", "一", "二", "三", "四", "五", "六", "日",
           "啟動日期", "播放模式", "啟動時間", "終止時間", "狀態", "音檔"]

EXAMPLES = [
    [1, "打掃", "✓", "", "✓", "", "", "", "", "", "單曲", "08:00:00", "08:00:53", "已排入時程", "上課.wav"],
    [2, "打掃結束", "✓", "", "✓", "", "", "", "", "", "單曲", "08:20:00", "08:20:55", "已排入時程", "下課.wav"],
    [3, "1預備", "", "", "✓", "✓", "", "", "", "", "單曲", "08:50:00", "08:50:20", "已排入時程", "01-預備鈴.mp3"],
    [4, "1考試", "", "", "✓", "✓", "", "", "", "", "單曲", "09:00:00", "09:00:19", "已排入時程", "02-考試鈴.mp3"],
]

NOTES = [
    "【使用方式】",
    "1. 在下方表格填入時程（可整批複製貼上）",
    "2. 開啟「廣播時程編輯器」→ 選一組鐘聲 → 按「匯入Excel」選本檔",
    "3. 確認表格無誤 → 按「匯出 .WEX」",
    "",
    "【欄位說明】",
    "序號：數字，顯示用排序",
    "名稱註記：必填，如「打掃」「1預備」「2考試」",
    "一~日：打 ✓ 或 v 或 1 = 當週執行，留空 = 不執行",
    "啟動日期：YYYY-MM-DD（留空 = 依星期循環，不指定日期）",
    "播放模式：目前僅「單曲」",
    "啟動時間：HH:MM:SS 或 HH:MM（秒固定 00）",
    "終止時間：自動計算（啟動時間 + 音檔長度），匯入時忽略此欄",
    "狀態：已排入時程 / 未排入",
    "音檔：Normal 資料夾內的檔名，如 上課.wav、下課.wav、01-預備鈴.mp3、02-考試鈴.mp3、03-結束鈴.mp3",
    "",
    "【可支援的音檔】",
    "上課.wav（53秒）、下課.wav（55秒）",
    "01-預備鈴.mp3（20秒）、02-考試鈴.mp3（19秒）、03-結束鈴.mp3（17秒）",
    "費玉青-晚安曲.mp3（215秒）、Alarm1/2/4.mp3",
]


def make_template(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "時程"

    header_font = Font(name="Microsoft JhengHei UI", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="2062AF", end_color="2062AF", fill_type="solid")
    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    example_font = Font(name="Microsoft JhengHei UI", size=10, color="888888")
    for r, row in enumerate(EXAMPLES, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = example_font
            cell.border = border
            if c >= 3 and c <= 9:
                cell.alignment = center

    widths = [6, 18, 4, 4, 4, 4, 4, 4, 4, 12, 8, 11, 11, 11, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    ws2 = wb.create_sheet("說明")
    ws2.column_dimensions["A"].width = 90
    note_font = Font(name="Microsoft JhengHei UI", size=10)
    title_font = Font(name="Microsoft JhengHei UI", size=11, bold=True)
    for i, line in enumerate(NOTES, 1):
        cell = ws2.cell(row=i, column=1, value=line)
        cell.font = title_font if line.startswith("【") else note_font

    wb.save(path)
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "時程Excel範本.xlsx")
    make_template(out)
    print(f"已產生: {out}")
