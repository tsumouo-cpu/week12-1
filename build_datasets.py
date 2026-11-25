# -*- coding: utf-8 -*-
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

output_dir = "C:/proj/week12-1"
os.makedirs(output_dir, exist_ok=True)

# Styles
font_header = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
font_sub_header = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
font_data = Font(name="微軟正黑體", size=10)
font_data_bold = Font(name="微軟正黑體", size=10, bold=True)
font_title = Font(name="微軟正黑體", size=14, bold=True, color="107C41")
font_note = Font(name="微軟正黑體", size=9, italic=True, color="605E5C")

fill_header_green = PatternFill(start_color="107C41", end_color="107C41", fill_type="solid")
fill_header_blue = PatternFill(start_color="0078D4", end_color="0078D4", fill_type="solid")
fill_header_dark = PatternFill(start_color="323130", end_color="323130", fill_type="solid")
fill_zebra = PatternFill(start_color="F9FBF9", end_color="F9FBF9", fill_type="solid")
fill_dirty_highlight = PatternFill(start_color="FFF0F0", end_color="FFF0F0", fill_type="solid")
fill_clean_highlight = PatternFill(start_color="F0FFF0", end_color="F0FFF0", fill_type="solid")

thin_border_side = Side(border_style="thin", color="D2D0CE")
border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
border_header = Border(left=thin_border_side, right=thin_border_side, top=Side(border_style="medium", color="107C41"), bottom=Side(border_style="medium", color="107C41"))

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")

# -------------------------------------------------------------
# 1. 建立 30 筆示範檔 (航運事務資料_30筆示範檔.xlsx)
# -------------------------------------------------------------
def build_30_demo():
    wb = openpyxl.Workbook()
    # Sheet 1: 原始資料 (未清洗)
    ws_raw = wb.active
    ws_raw.title = "01_原始資料_未清洗"

    ws_raw.views.sheetView[0].showGridLines = True
    ws_raw.append(["航運行政業務日常報表 - 原始匯出未整理資料檔 (含多種常見格式瑕疵)"])
    ws_raw.cell(row=1, column=1).font = font_title
    ws_raw.append(["※ 說明：本工作表為直接從舊版系統或外部報關行收集之原始紀錄，包含前後空格、未分欄、文字數值、格式不一等瑕疵。"])
    ws_raw.cell(row=2, column=1).font = font_note
    ws_raw.append([]) # Row 3 blank

    raw_headers = [
        "項次", "提單號碼(B/L No.)", "船名與航次(Vessel/Voyage)", "裝貨港(POL)", "卸貨港(POD)",
        "櫃型尺寸", "件數與單位(Packages)", "毛重(Gross Wt)", "海運費(Freight USD)", "付款方式", "申報日期", "報關狀態"
    ]
    ws_raw.append(raw_headers)
    for col_idx in range(1, len(raw_headers) + 1):
        cell = ws_raw.cell(row=4, column=col_idx)
        cell.font = font_header
        cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid") # Red for dirty
        cell.alignment = align_center
        cell.border = border_all

    raw_rows = [
        (1, "  EVER10928301  ", "EVER GIVEN / 0124E", "TWKHH", "USLAX", "40HC", "1,200 CTNS", "18,450.00 KGS", "$ 2,450.00", "PP", "2026/03/01", "Released"),
        (2, "whlu2839102", "WAN HAI 315 / W082", "twkhh", "JPTYO", "20GP", "450 PLTS", "8,200 KGS", "USD 850", "pp", "2026-03-01", "released"),
        (3, "MAEU99201928", "MAERSK MC-KINNEY / 2401W", "TWTPE", "NLRTM", "40HC", "850 PKGS", "14,200.50 KGS", "3,100.00", "CC", "02/03/2026", "Released"),
        (4, "YMLU88172039  ", "YM WELLHEAD / 221S", "TWKHH", "USOAK", "40GP", "2,100 CTNS", "21,000 KGS", "$ 2,250", "PP", "2026/03/02", "Pending"),
        (5, "ONEY77281920", "ONE TRIUMPH / 032E", "TWKHH", "USLGB", "40HC", "1,500 CTNS", "19,800.00 KGS", "2,600.00", "PP", "20260303", "Released"), # Duplicate row anchor
        (6, "  cosu66271920", "COSCO SHIPPING TAURUS / 019W", "twtxg", "DEHAM", "40HC", "950 PKGS", "16,400 KGS", "USD 2,950.00", "cc", "2026/03/03", "Released"),
        (7, "EVER10928302", "EVER GOLDEN / 0088E", "TWKHH", "USLAX", "40HC", "1,100 CTNS", "17,500.00 KGS", "$ 2,450.00", "PP", "2026-03-04", "Released"),
        (8, "MSCU55192830", "MSC OSCAR / 2405W", "TWKHH", "NLRTM", "40HC", "720 PLTS", "15,800 KGS", "3,050.00", "CC", "2026/03/04", "Released"),
        (9, "WHLU99281726", "WAN HAI 271 / S192", "TWKHH", "SGSIN", "20GP", "380 DRUMS", "9,500.00 KGS", "920", "PP", "05/03/2026", "Released"),
        (10, "OOLU28192810  ", "OOCL POLAND / 054E", "TWKHH", "USLGB", "40GP", "1,800 CTNS", "18,900 KGS", "$ 2,300.00", "pp", "20260305", "Pending"),
        (11, "HLCU88192831", "AL DAHNA EXPRESS / 2402W", "TWTPE", "DEHAM", "40HC", "640 CRATES", "13,600.00 KGS", "USD 3,150", "CC", "2026/03/06", "Released"),
        (12, "  cma99182736", "CMA CGM ANTOINE / 092E", "TWKHH", "USOAK", "40HC", "1,350 CTNS", "20,100 KGS", "2,500.00", "PP", "2026-03-06", "Released"),
        (13, "EVER10928303", "EVER BRAVE / 0310S", "TWKHH", "THBKK", "20GP", "520 CTNS", "7,800.00 KGS", "$ 780.00", "PP", "2026/03/07", "Released"),
        (14, "YMLU33819280", "YM TARGET / 042E", "twkhh", "USLAX", "40HC", "1,400 CTNS", "19,200 KGS", "USD 2,400", "PP", "07/03/2026", "Inspection"),
        (15, "ONEY77281920", "ONE TRIUMPH / 032E", "TWKHH", "USLGB", "40HC", "1,500 CTNS", "19,800.00 KGS", "2,600.00", "PP", "20260303", "Released"), # DUPLICATE of Row 5!
        (16, "SMLU77182910  ", "SM YANTIAN / 2401E", "TWKHH", "USLGB", "40GP", "1,250 CTNS", "17,200.00 KGS", "$ 2,350.00", "pp", "2026/03/08", "Released"),
        (17, "MAEU11827364", "MUNICH MAERSK / 2403W", "TWTPE", "NLRTM", "40HC", "890 PKGS", "16,100 KGS", "3,200.00", "CC", "2026-03-08", "Released"),
        (18, "WHLU66182910", "WAN HAI 516 / W043", "twtxg", "VNCLI", "20GP", "600 BAGS", "11,500.00 KGS", "650", "PP", "20260309", "Released"),
        (19, "  EVER10928304", "EVER FORTUNE / 0205E", "TWKHH", "USLAX", "40HC", "1,600 CTNS", "21,500.00 KGS", "$ 2,550.00", "PP", "2026/03/09", "Released"),
        (20, "MSCU99281720", "MSC GULSUN / 2406W", "TWKHH", "DEHAM", "40HC", "780 PLTS", "15,200 KGS", "USD 3,000", "cc", "10/03/2026", "Released"),
        (21, "cosu11827391", "COSCO GLORY / 088W", "TWKHH", "NLRTM", "40HC", "1,050 PKGS", "18,200.00 KGS", "2,980.00", "CC", "2026-03-10", "Pending"),
        (22, "ONEY55182910  ", "ONE HARBOUR / 018E", "TWKHH", "USOAK", "40HC", "1,300 CTNS", "19,400 KGS", "$ 2,480.00", "PP", "2026/03/11", "Released"),
        (23, "YMLU77182912", "YM TRAVEL / 035E", "twkhh", "USLAX", "40GP", "1,150 CTNS", "16,800.00 KGS", "2,380.00", "pp", "20260311", "Released"),
        (24, "HLCU33819282", "HAMBURG EXPRESS / 2404W", "TWTPE", "DEHAM", "40HC", "920 CTNS", "14,900 KGS", "USD 3,100.00", "CC", "12/03/2026", "Released"),
        (25, "  whlu55192830", "WAN HAI 162 / S221", "TWKHH", "MYPKG", "20GP", "480 PLTS", "8,900.00 KGS", "820", "PP", "2026/03/12", "Released"),
        (26, "EVER10928305", "EVER APEX / 0012E", "TWKHH", "USLAX", "40HC", "1,750 CTNS", "22,800.00 KGS", "$ 2,600.00", "PP", "2026-03-13", "Released"),
        (27, "EVER10928305", "EVER APEX / 0012E", "TWKHH", "USLAX", "40HC", "1,750 CTNS", "22,800.00 KGS", "$ 2,600.00", "PP", "2026-03-13", "Released"), # DUPLICATE of Row 26!
        (28, "OOLU99182730  ", "OOCL SCANDINAVIA / 028W", "TWKHH", "NLRTM", "40HC", "810 PKGS", "15,600 KGS", "3,020.00", "cc", "20260314", "Released"),
        (29, "cma44182910", "CMA CGM BOUGAINVILLE / 041E", "twtxg", "USLGB", "40HC", "1,450 CTNS", "20,500.00 KGS", "USD 2,520", "PP", "14/03/2026", "Released"),
        (30, "  YMLU99182733  ", "YM UTILITY / 067E", "TWKHH", "USOAK", "40GP", "1,220 CTNS", "17,600 KGS", "$ 2,420.00", "PP", "2026/03/15", "Released")
    ]

    for r_idx, r_data in enumerate(raw_rows):
        ws_raw.append(list(r_data))
        row_num = 5 + r_idx
        for c_idx in range(1, len(r_data) + 1):
            cell = ws_raw.cell(row=row_num, column=c_idx)
            cell.font = font_data
            cell.border = border_all
            # Highlight duplicate rows in pink
            if r_idx in [4, 14, 25, 26]:
                cell.fill = fill_dirty_highlight
            elif r_idx % 2 == 1:
                cell.fill = fill_zebra
            
            # Alignments
            if c_idx in [1, 4, 5, 6, 10, 11, 12]:
                cell.alignment = align_center
            elif c_idx in [7, 8, 9]:
                cell.alignment = align_right
            else:
                cell.alignment = align_left

    # Sheet 2: 清洗後標準資料
    ws_clean = wb.create_sheet(title="02_清洗後標準資料")
    ws_clean.views.sheetView[0].showGridLines = True
    ws_clean.append(["航運行政業務日常報表 - 標準化前處理完成檔 (分析就緒 Ready for Analysis)"])
    ws_clean.cell(row=1, column=1).font = font_title
    ws_clean.append(["※ 說明：已完成提單號去空格大寫、剔除2筆重複項(保留28筆)、船名與航次拆分、件數與單位分離、運費與毛重轉為純數值、日期標準化。"])
    ws_clean.cell(row=2, column=1).font = font_note
    ws_clean.append([])

    clean_headers = [
        "項次", "提單號碼", "船名 (Vessel)", "航次 (Voyage)", "裝貨港 (POL)", "卸貨港 (POD)",
        "櫃型尺寸", "件數 (數量)", "包裝單位", "毛重 (KG)", "海運費 (USD)", "付款方式", "申報日期", "報關狀態"
    ]
    ws_clean.append(clean_headers)
    for col_idx in range(1, len(clean_headers) + 1):
        cell = ws_clean.cell(row=4, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header_green
        cell.alignment = align_center
        cell.border = border_all

    # 28 Cleaned records (Deduplicated, standardized)
    clean_rows = [
        (1, "EVER10928301", "EVER GIVEN", "0124E", "TWKHH", "USLAX", "40HC", 1200, "CTN", 18450.00, 2450.00, "PP", "2026-03-01", "RELEASED"),
        (2, "WHLU2839102", "WAN HAI 315", "W082", "TWKHH", "JPTYO", "20GP", 450, "PLT", 8200.00, 850.00, "PP", "2026-03-01", "RELEASED"),
        (3, "MAEU99201928", "MAERSK MC-KINNEY", "2401W", "TWTPE", "NLRTM", "40HC", 850, "PKG", 14200.50, 3100.00, "CC", "2026-03-02", "RELEASED"),
        (4, "YMLU88172039", "YM WELLHEAD", "221S", "TWKHH", "USOAK", "40GP", 2100, "CTN", 21000.00, 2250.00, "PP", "2026-03-02", "PENDING"),
        (5, "ONEY77281920", "ONE TRIUMPH", "032E", "TWKHH", "USLGB", "40HC", 1500, "CTN", 19800.00, 2600.00, "PP", "2026-03-03", "RELEASED"),
        (6, "COSU66271920", "COSCO SHIPPING TAURUS", "019W", "TWTXG", "DEHAM", "40HC", 950, "PKG", 16400.00, 2950.00, "CC", "2026-03-03", "RELEASED"),
        (7, "EVER10928302", "EVER GOLDEN", "0088E", "TWKHH", "USLAX", "40HC", 1100, "CTN", 17500.00, 2450.00, "PP", "2026-03-04", "RELEASED"),
        (8, "MSCU55192830", "MSC OSCAR", "2405W", "TWKHH", "NLRTM", "40HC", 720, "PLT", 15800.00, 3050.00, "CC", "2026-03-04", "RELEASED"),
        (9, "WHLU99281726", "WAN HAI 271", "S192", "TWKHH", "SGSIN", "20GP", 380, "DRUM", 9500.00, 920.00, "PP", "2026-03-05", "RELEASED"),
        (10, "OOLU28192810", "OOCL POLAND", "054E", "TWKHH", "USLGB", "40GP", 1800, "CTN", 18900.00, 2300.00, "PP", "2026-03-05", "PENDING"),
        (11, "HLCU88192831", "AL DAHNA EXPRESS", "2402W", "TWTPE", "DEHAM", "40HC", 640, "CRATE", 13600.00, 3150.00, "CC", "2026-03-06", "RELEASED"),
        (12, "CMA99182736", "CMA CGM ANTOINE", "092E", "TWKHH", "USOAK", "40HC", 1350, "CTN", 20100.00, 2500.00, "PP", "2026-03-06", "RELEASED"),
        (13, "EVER10928303", "EVER BRAVE", "0310S", "TWKHH", "THBKK", "20GP", 520, "CTN", 7800.00, 780.00, "PP", "2026-03-07", "RELEASED"),
        (14, "YMLU33819280", "YM TARGET", "042E", "TWKHH", "USLAX", "40HC", 1400, "CTN", 19200.00, 2400.00, "PP", "2026-03-07", "INSPECTION"),
        (15, "SMLU77182910", "SM YANTIAN", "2401E", "TWKHH", "USLGB", "40GP", 1250, "CTN", 17200.00, 2350.00, "PP", "2026-03-08", "RELEASED"),
        (16, "MAEU11827364", "MUNICH MAERSK", "2403W", "TWTPE", "NLRTM", "40HC", 890, "PKG", 16100.00, 3200.00, "CC", "2026-03-08", "RELEASED"),
        (17, "WHLU66182910", "WAN HAI 516", "W043", "TWTXG", "VNCLI", "20GP", 600, "BAG", 11500.00, 650.00, "PP", "2026-03-09", "RELEASED"),
        (18, "EVER10928304", "EVER FORTUNE", "0205E", "TWKHH", "USLAX", "40HC", 1600, "CTN", 21500.00, 2550.00, "PP", "2026-03-09", "RELEASED"),
        (19, "MSCU99281720", "MSC GULSUN", "2406W", "TWKHH", "DEHAM", "40HC", 780, "PLT", 15200.00, 3000.00, "CC", "2026-03-10", "RELEASED"),
        (20, "COSU11827391", "COSCO GLORY", "088W", "TWKHH", "NLRTM", "40HC", 1050, "PKG", 18200.00, 2980.00, "CC", "2026-03-10", "PENDING"),
        (21, "ONEY55182910", "ONE HARBOUR", "018E", "TWKHH", "USOAK", "40HC", 1300, "CTN", 19400.00, 2480.00, "PP", "2026-03-11", "RELEASED"),
        (22, "YMLU77182912", "YM TRAVEL", "035E", "TWKHH", "USLAX", "40GP", 1150, "CTN", 16800.00, 2380.00, "PP", "2026-03-11", "RELEASED"),
        (23, "HLCU33819282", "HAMBURG EXPRESS", "2404W", "TWTPE", "DEHAM", "40HC", 920, "CTN", 14900.00, 3100.00, "CC", "2026-03-12", "RELEASED"),
        (24, "WHLU55192830", "WAN HAI 162", "S221", "TWKHH", "MYPKG", "20GP", 480, "PLT", 8900.00, 820.00, "PP", "2026-03-12", "RELEASED"),
        (25, "EVER10928305", "EVER APEX", "0012E", "TWKHH", "USLAX", "40HC", 1750, "CTN", 22800.00, 2600.00, "PP", "2026-03-13", "RELEASED"),
        (26, "OOLU99182730", "OOCL SCANDINAVIA", "028W", "TWKHH", "NLRTM", "40HC", 810, "PKG", 15600.00, 3020.00, "CC", "2026-03-14", "RELEASED"),
        (27, "CMA44182910", "CMA CGM BOUGAINVILLE", "041E", "TWTXG", "USLGB", "40HC", 1450, "CTN", 20500.00, 2520.00, "PP", "2026-03-14", "RELEASED"),
        (28, "YMLU99182733", "YM UTILITY", "067E", "TWKHH", "USOAK", "40GP", 1220, "CTN", 17600.00, 2420.00, "PP", "2026-03-15", "RELEASED")
    ]

    for r_idx, r_data in enumerate(clean_rows):
        ws_clean.append(list(r_data))
        row_num = 5 + r_idx
        for c_idx in range(1, len(r_data) + 1):
            cell = ws_clean.cell(row=row_num, column=c_idx)
            cell.font = font_data
            cell.border = border_all
            if r_idx % 2 == 1:
                cell.fill = fill_zebra
            
            # Alignments & formats
            if c_idx in [1, 2, 4, 5, 6, 7, 9, 12, 13, 14]:
                cell.alignment = align_center
            elif c_idx in [8, 10, 11]:
                cell.alignment = align_right
                if c_idx == 8:
                    cell.number_format = "#,##0"
                elif c_idx == 10:
                    cell.number_format = "#,##0.00"
                elif c_idx == 11:
                    cell.number_format = "$#,##0.00"
            else:
                cell.alignment = align_left

    # Summary Row for Clean Data
    sum_row_idx = 5 + len(clean_rows)
    ws_clean.cell(row=sum_row_idx, column=1, value="合計 / 摘要統計").font = font_data_bold
    ws_clean.cell(row=sum_row_idx, column=1).alignment = align_center
    ws_clean.cell(row=sum_row_idx, column=8, value=f"=SUM(H5:H{sum_row_idx-1})").font = font_data_bold
    ws_clean.cell(row=sum_row_idx, column=8).number_format = "#,##0"
    ws_clean.cell(row=sum_row_idx, column=10, value=f"=SUM(J5:J{sum_row_idx-1})").font = font_data_bold
    ws_clean.cell(row=sum_row_idx, column=10).number_format = "#,##0.00"
    ws_clean.cell(row=sum_row_idx, column=11, value=f"=SUM(K5:K{sum_row_idx-1})").font = font_data_bold
    ws_clean.cell(row=sum_row_idx, column=11).number_format = "$#,##0.00"
    for c_i in range(1, len(clean_headers)+1):
        c = ws_clean.cell(row=sum_row_idx, column=c_i)
        c.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        c.border = Border(top=Side(border_style="thin", color="107C41"), bottom=Side(border_style="double", color="107C41"))

    # Auto-fit column widths
    for ws in [ws_raw, ws_clean]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                if cell.row in [1, 2]: continue
                # Count length roughly
                cur_len = sum(2 if ord(ch) > 127 else 1 for ch in val)
                if cur_len > max_len: max_len = cur_len
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    demo_path = os.path.join(output_dir, "航運事務資料_30筆示範檔.xlsx")
    wb.save(demo_path)
    print(f"Demo file created at: {demo_path}")

# -------------------------------------------------------------
# 2. 建立 300 筆多情境航運實戰練習檔 (航運案例情境_300筆實戰練習檔.xlsx)
# -------------------------------------------------------------
def build_300_practice():
    import random
    wb = openpyxl.Workbook()
    ws_raw = wb.active
    ws_raw.title = "01_練習題_300筆待處理資料"
    ws_raw.views.sheetView[0].showGridLines = True

    ws_raw.append(["航運行政高階實務 - 300 筆多情境案例前處理實戰練習檔"])
    ws_raw.cell(row=1, column=1).font = font_title
    ws_raw.append(["※ 任務要求：本表包含美加遠洋、歐地、東南亞近洋、特種冷鏈等300筆業務資料，請依教學手冊指引完成去重、分欄、去空格、數值化與日期格式標準化。"])
    ws_raw.cell(row=2, column=1).font = font_note
    ws_raw.append([])

    headers_raw = [
        "序號", "提單號碼 (B/L No.)", "船名航次 (Vessel/Voyage)", "裝貨港代碼 (POL)", "卸貨港代碼 (POD)",
        "貨櫃規格 (Type)", "件數與單位 (Packages)", "毛重 (Gross Weight)", "運費金額 (Ocean Freight)",
        "付款條款 (Terms)", "申報結關日 (Filing Date)", "業務情境分類", "通關狀態 (Customs Status)"
    ]
    ws_raw.append(headers_raw)
    for col_idx in range(1, len(headers_raw) + 1):
        cell = ws_raw.cell(row=4, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header_dark
        cell.alignment = align_center
        cell.border = border_all

    # Seed for deterministic data generation
    random.seed(42)

    carriers = [
        ("EVER", "EVER GIVEN", "EVER GOLDEN", "EVER BRAVE", "EVER APEX", "EVER FORTUNE"),
        ("WHLU", "WAN HAI 315", "WAN HAI 271", "WAN HAI 516", "WAN HAI 162", "WAN HAI 291"),
        ("MAEU", "MAERSK MC-KINNEY", "MUNICH MAERSK", "MADRID MAERSK", "MANCHESTER MAERSK"),
        ("YMLU", "YM WELLHEAD", "YM TARGET", "YM TRAVEL", "YM UTILITY", "YM TRIUMPH"),
        ("ONEY", "ONE TRIUMPH", "ONE HARBOUR", "ONE APUS", "ONE STORK"),
        ("COSU", "COSCO TAURUS", "COSCO GLORY", "COSCO SHIPPING PLANET", "COSCO NEBULA"),
        ("MSCU", "MSC OSCAR", "MSC GULSUN", "MSC ISABELLA", "MSC MIA"),
        ("HLCU", "HAMBURG EXPRESS", "AL DAHNA EXPRESS", "BERLIN EXPRESS")
    ]

    pols = [("TWKHH", "twkhh", " TWKHH ", "TW-KHH"), ("TWTPE", "twtpe", " TWTPE "), ("TWTXG", "twtxg", " TWTXG ")]
    
    scenarios_pool = [
        ("A-美加遠洋", ["USLAX", "USLGB", "USOAK", "USSEA", "USNYC", "USSAV"], ["40HC", "40GP", "45HC"], 2200, 3600),
        ("B-歐地航線", ["NLRTM", "DEHAM", "BEANR", "GBFXT", "ITGOA"], ["40HC", "20GP", "40GP"], 2600, 4200),
        ("C-東南亞近洋", ["SGSIN", "MYPKG", "THBKK", "VNSGN", "IDJKT", "PHMNL"], ["20GP", "40GP"], 650, 1400),
        ("D-特種冷鏈", ["USLAX", "NLRTM", "JPTYO", "KRPUS", "SGSIN"], ["40HR", "20RF"], 3500, 6800)
    ]

    units_pool = ["CTNS", "PLTS", "PKGS", "BAGS", "DRUMS", "CRATES", "BOXES", "ROLLS"]
    status_pool = ["Released", "released", "Pending", "pending", "Inspection", "RELEASED"]

    # Generate 285 unique base records + 15 exact duplicates = 300 records
    unique_records = []
    
    for i in range(1, 286):
        # Scenario
        if i <= 80:
            scen_name, pod_list, types_list, min_fr, max_fr = scenarios_pool[0]
        elif i <= 155:
            scen_name, pod_list, types_list, min_fr, max_fr = scenarios_pool[1]
        elif i <= 230:
            scen_name, pod_list, types_list, min_fr, max_fr = scenarios_pool[2]
        else:
            scen_name, pod_list, types_list, min_fr, max_fr = scenarios_pool[3]

        c_prefix, *v_names = random.choice(carriers)
        vessel = random.choice(v_names)
        voy = f"{random.randint(10, 99):02d}{random.randint(1, 9):01d}{random.choice(['E','W','S','N'])}"
        
        # B/L No
        clean_bl = f"{c_prefix}{random.randint(1000000, 9999999)}"
        # Add dirty variants to B/L
        dice = random.random()
        if dice < 0.2:
            dirty_bl = f"  {clean_bl}  "
        elif dice < 0.35:
            dirty_bl = clean_bl.lower()
        elif dice < 0.45:
            dirty_bl = f"{clean_bl} "
        else:
            dirty_bl = clean_bl

        # Vessel / Voyage dirty
        dice_v = random.random()
        if dice_v < 0.6:
            dirty_ves_voy = f"{vessel} / {voy}"
        elif dice_v < 0.8:
            dirty_ves_voy = f"{vessel}  {voy}"
        else:
            dirty_ves_voy = f"{vessel} V.{voy}"

        # POL
        pol_clean = random.choice(["TWKHH", "TWTPE", "TWTXG"])
        dice_pol = random.random()
        if dice_pol < 0.25:
            dirty_pol = pol_clean.lower()
        elif dice_pol < 0.4:
            dirty_pol = f" {pol_clean} "
        elif dice_pol < 0.5 and pol_clean == "TWKHH":
            dirty_pol = "KHH"
        else:
            dirty_pol = pol_clean

        # POD
        pod_clean = random.choice(pod_list)
        dirty_pod = pod_clean.lower() if random.random() < 0.2 else (f" {pod_clean} " if random.random() < 0.3 else pod_clean)

        # Container type
        cntr_type = random.choice(types_list)

        # Qty & Unit
        qty = random.randint(120, 2800)
        unit = random.choice(units_pool)
        dice_q = random.random()
        if dice_q < 0.5:
            dirty_pkg = f"{qty:,} {unit}"
        elif dice_q < 0.8:
            dirty_pkg = f"{qty} {unit}"
        else:
            dirty_pkg = f"{qty:,}{unit.lower()}"

        # Gross Weight
        gw_kg = round(qty * random.uniform(8.5, 22.0), 2)
        dice_w = random.random()
        if dice_w < 0.4:
            dirty_gw = f"{gw_kg:,.2f} KGS"
        elif dice_w < 0.7:
            dirty_gw = f"{gw_kg:,.2f} kg"
        elif dice_w < 0.85:
            dirty_gw = f"{gw_kg:,.1f}"
        else:
            dirty_gw = str(gw_kg)

        # Ocean Freight
        freight = round(random.uniform(min_fr, max_fr), 2)
        dice_f = random.random()
        if dice_f < 0.35:
            dirty_fr = f"$ {freight:,.2f}"
        elif dice_f < 0.6:
            dirty_fr = f"USD {freight:,.2f}"
        elif dice_f < 0.8:
            dirty_fr = f"{freight:,.2f}"
        else:
            dirty_fr = str(freight)

        # Payment terms
        terms_clean = random.choice(["PP", "CC"])
        dice_t = random.random()
        if dice_t < 0.3:
            dirty_terms = terms_clean.lower()
        elif dice_t < 0.5:
            dirty_terms = "Prepaid" if terms_clean == "PP" else "Collect"
        elif dice_t < 0.55:
            dirty_terms = "" # missing!
        else:
            dirty_terms = terms_clean

        # Date
        m = random.randint(1, 4)
        d = random.randint(1, 28)
        clean_date_str = f"2026-{m:02d}-{d:02d}"
        dice_d = random.random()
        if dice_d < 0.3:
            dirty_date = f"2026/{m:02d}/{d:02d}"
        elif dice_d < 0.55:
            dirty_date = f"{d:02d}/{m:02d}/2026"
        elif dice_d < 0.75:
            dirty_date = f"2026{m:02d}{d:02d}"
        else:
            dirty_date = clean_date_str

        # Status
        st = random.choice(status_pool)

        clean_rec = {
            "bl": clean_bl,
            "vessel": vessel,
            "voy": voy,
            "pol": pol_clean,
            "pod": pod_clean,
            "type": cntr_type,
            "qty": qty,
            "unit": unit.rstrip('S').upper(), # e.g. CTN, PLT
            "gw": gw_kg,
            "freight": freight,
            "terms": "PP" if terms_clean.upper() in ["PP", "PREPAID"] or dirty_terms=="" else "CC",
            "date": clean_date_str,
            "scen": scen_name,
            "status": st.upper()
        }

        dirty_rec = [
            i, dirty_bl, dirty_ves_voy, dirty_pol, dirty_pod,
            cntr_type, dirty_pkg, dirty_gw, dirty_fr,
            dirty_terms, dirty_date, scen_name, st
        ]

        unique_records.append((dirty_rec, clean_rec))

    # Pick 15 records to duplicate
    dup_indices = [12, 35, 68, 92, 115, 142, 168, 189, 210, 228, 245, 260, 271, 279, 284]
    all_300_dirty = []
    
    seq = 1
    for idx, (d_rec, c_rec) in enumerate(unique_records):
        d_rec[0] = seq
        all_300_dirty.append((d_rec, c_rec, False))
        seq += 1
        if idx in dup_indices:
            # Clone dirty record
            cloned = list(d_rec)
            cloned[0] = seq
            all_300_dirty.append((cloned, c_rec, True)) # is duplicate
            seq += 1

    # Write Raw sheet
    for r_idx, (d_rec, _, is_dup) in enumerate(all_300_dirty):
        ws_raw.append(d_rec)
        row_n = 5 + r_idx
        for col_n in range(1, len(d_rec) + 1):
            cell = ws_raw.cell(row=row_n, column=col_n)
            cell.font = font_data
            cell.border = border_all
            if is_dup:
                cell.fill = fill_dirty_highlight
            elif r_idx % 2 == 1:
                cell.fill = fill_zebra
            if col_n in [1, 4, 5, 6, 10, 11, 12, 13]:
                cell.alignment = align_center
            elif col_n in [7, 8, 9]:
                cell.alignment = align_right
            else:
                cell.alignment = align_left

    # Sheet 2: Clean Solution (285 records)
    ws_clean = wb.create_sheet(title="02_參考解答_標準化結果")
    ws_clean.views.sheetView[0].showGridLines = True
    ws_clean.append(["300筆實戰練習檔 - 參考解答 (已去除15筆重複，共保留285筆唯一有效標準記錄)"])
    ws_clean.cell(row=1, column=1).font = font_title
    ws_clean.append(["※ 檢核標準：唯一提單數 = 285 列；所有文字前後空白已清除；數值欄位可精確加總；日期格式統一為 YYYY-MM-DD。"])
    ws_clean.cell(row=2, column=1).font = font_note
    ws_clean.append([])

    headers_clean = [
        "項次", "提單號碼", "船名 (Vessel)", "航次 (Voyage)", "裝貨港 (POL)", "卸貨港 (POD)",
        "櫃型規格", "件數 (數量)", "包裝單位", "毛重 (KG)", "運費金額 (USD)", "付款方式", "申報日期", "業務情境", "通關狀態"
    ]
    ws_clean.append(headers_clean)
    for col_idx in range(1, len(headers_clean) + 1):
        cell = ws_clean.cell(row=4, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header_blue
        cell.alignment = align_center
        cell.border = border_all

    total_qty = 0
    total_gw = 0.0
    total_fr = 0.0

    for idx, (_, c_rec) in enumerate(unique_records):
        clean_row = [
            idx + 1, c_rec["bl"], c_rec["vessel"], c_rec["voy"], c_rec["pol"], c_rec["pod"],
            c_rec["type"], c_rec["qty"], c_rec["unit"], c_rec["gw"], c_rec["freight"],
            c_rec["terms"], c_rec["date"], c_rec["scen"], c_rec["status"]
        ]
        total_qty += c_rec["qty"]
        total_gw += c_rec["gw"]
        total_fr += c_rec["freight"]

        ws_clean.append(clean_row)
        row_n = 5 + idx
        for col_n in range(1, len(clean_row) + 1):
            cell = ws_clean.cell(row=row_n, column=col_n)
            cell.font = font_data
            cell.border = border_all
            if idx % 2 == 1:
                cell.fill = fill_zebra
            if col_n in [1, 2, 4, 5, 6, 7, 9, 12, 13, 14, 15]:
                cell.alignment = align_center
            elif col_n in [8, 10, 11]:
                cell.alignment = align_right
                if col_n == 8:
                    cell.number_format = "#,##0"
                elif col_n == 10:
                    cell.number_format = "#,##0.00"
                elif col_n == 11:
                    cell.number_format = "$#,##0.00"
            else:
                cell.alignment = align_left

    # Summary Row in Sheet 2
    sum_row_idx = 5 + len(unique_records)
    ws_clean.cell(row=sum_row_idx, column=1, value="總計 (Total)").font = font_data_bold
    ws_clean.cell(row=sum_row_idx, column=1).alignment = align_center
    ws_clean.cell(row=sum_row_idx, column=8, value=f"=SUM(H5:H{sum_row_idx-1})").font = font_data_bold
    ws_clean.cell(row=sum_row_idx, column=8).number_format = "#,##0"
    ws_clean.cell(row=sum_row_idx, column=10, value=f"=SUM(J5:J{sum_row_idx-1})").font = font_data_bold
    ws_clean.cell(row=sum_row_idx, column=10).number_format = "#,##0.00"
    ws_clean.cell(row=sum_row_idx, column=11, value=f"=SUM(K5:K{sum_row_idx-1})").font = font_data_bold
    ws_clean.cell(row=sum_row_idx, column=11).number_format = "$#,##0.00"

    for c_i in range(1, len(headers_clean) + 1):
        c = ws_clean.cell(row=sum_row_idx, column=c_i)
        c.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        c.border = Border(top=Side(border_style="thin", color="0078D4"), bottom=Side(border_style="double", color="0078D4"))

    # Sheet 3: 解題指引與樞紐統計摘要
    ws_guide = wb.create_sheet(title="03_解題檢核與統計驗證")
    ws_guide.views.sheetView[0].showGridLines = True
    ws_guide.append(["300筆實戰練習檔 - 官方檢核數字與統計基準"])
    ws_guide.cell(row=1, column=1).font = font_title
    ws_guide.append(["※ 學員完成清洗後，請比對下列各項關鍵加總值，若完全一致即代表資料前處理 100% 正確！"])
    ws_guide.cell(row=2, column=1).font = font_note
    ws_guide.append([])

    benchmarks = [
        ("檢核指標", "官方基準數值", "計算公式 / 驗證方法", "常見出錯原因分析"),
        ("原始總筆數", "300 筆", "=COUNTA(A5:A304)", "未完整載入原始工作表"),
        ("剔除重複後筆數", "285 筆 (移除15筆重複)", "資料 > 移除重複項 (僅勾選提單號)", "誤勾選其他欄位導致未完全去重"),
        ("總件數加總 (Packages)", f"{total_qty:,} 件", "=SUM(H5:H289)", "未將包裝文字單位拆出，SUM結果為0"),
        ("總毛重加總 (Gross Wt)", f"{total_gw:,.2f} KG", "=SUM(J5:J289)", "字尾KGS未替換，千分位被識別為文字"),
        ("海運費總額 (Freight USD)", f"${total_fr:,.2f}", "=SUM(K5:K289)", "貨幣符號未去除或未轉換為純數值"),
        ("唯一裝貨港數", "3 個 (TWKHH, TWTPE, TWTXG)", "=COUNTA(UNIQUE(E5:E289))", "大小寫未統一或'KHH'未校正為TWKHH"),
        ("有效日期範圍", "2026-01-01 至 2026-04-28", "=MIN(M5:M289) & ' ~ ' & MAX(M5:M289)", "文字型日期無法參與最小值/最大值運算")
    ]

    for r_i, r_vals in enumerate(benchmarks):
        ws_guide.append(list(r_vals))
        row_n = 4 + r_i
        for c_i, v in enumerate(r_vals):
            cell = ws_guide.cell(row=row_n, column=c_i + 1)
            cell.border = border_all
            if r_i == 0:
                cell.font = font_header
                cell.fill = fill_header_green
                cell.alignment = align_center
            else:
                cell.font = font_data_bold if c_i in [0, 1] else font_data
                if c_i == 1:
                    cell.font = Font(name="微軟正黑體", size=10, bold=True, color="C00000")
                if r_i % 2 == 1:
                    cell.fill = fill_zebra
                cell.alignment = align_center if c_i in [0, 1] else align_left

    # Auto-fit column widths for all sheets
    for ws in [ws_raw, ws_clean, ws_guide]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                if cell.row in [1, 2]: continue
                cur_len = sum(2 if ord(ch) > 127 else 1 for ch in val)
                if cur_len > max_len: max_len = cur_len
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    practice_path = os.path.join(output_dir, "航運案例情境_300筆實戰練習檔.xlsx")
    wb.save(practice_path)
    print(f"Practice file created at: {practice_path}")
    print(f"Stats: Total Unique = {len(unique_records)}, Total Qty = {total_qty}, Total GW = {total_gw:.2f}, Total Freight = {total_fr:.2f}")

if __name__ == "__main__":
    build_30_demo()
    build_300_practice()
