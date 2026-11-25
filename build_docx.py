# -*- coding: utf-8 -*-
import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

output_path = "C:/proj/week12-1/航運行政人員_Excel資料前處理實務教學手冊.docx"
assets_dir = "C:/proj/week12-1/assets"

doc = Document()

# Page setup: A4, 1-inch margins
for section in doc.sections:
    section.page_width = Inches(8.27)  # A4
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

# Colors
COLOR_NAVY = RGBColor(0x00, 0x45, 0x78)
COLOR_GREEN = RGBColor(0x10, 0x7C, 0x41)
COLOR_DARK = RGBColor(0x20, 0x1F, 0x1E)
COLOR_GRAY = RGBColor(0x60, 0x5E, 0x5C)
COLOR_ORANGE = RGBColor(0xD8, 0x3B, 0x01)

# Helper: Set font with EastAsia support
def set_run_font(run, name="微軟正黑體", size_pt=10.5, bold=False, italic=False, color=None):
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    # Set eastAsia font
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rFonts.set(qn('w:eastAsia'), name)
    rPr.append(rFonts)

# Helper: Add paragraph with styling
def add_p(doc, text="", style='Normal', space_before=0, space_after=6, line_spacing=1.2, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, color=None, font_size=10.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    p.alignment = align
    if text:
        run = p.add_run(text)
        set_run_font(run, size_pt=font_size, bold=bold, color=color)
    return p

# Helper: Add Heading 1 with Page Break Before
def add_h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.page_break_before = True
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    set_run_font(run, name="微軟正黑體", size_pt=18, bold=True, color=COLOR_GREEN)
    # Add bottom border XML to Heading 1
    pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                     r'<w:bottom w:val="single" w:sz="12" w:space="4" w:color="107C41"/>'
                     r'</w:pBdr>')
    p._p.get_or_add_pPr().append(pBdr)
    return p

# Helper: Add Heading 2
def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    set_run_font(run, name="微軟正黑體", size_pt=13.5, bold=True, color=COLOR_NAVY)
    return p

# Helper: Add Heading 3
def add_h3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    set_run_font(run, name="微軟正黑體", size_pt=11.5, bold=True, color=COLOR_DARK)
    return p

# Helper: Cell Background
def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

# Helper: Add Callout Box (提示、重要、注意事項)
def add_callout(doc, title, text, box_type="TIP"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    cell = table.cell(0, 0)
    cell.width = Inches(6.2)
    
    bg_map = {"TIP": "EBF6EF", "WARN": "FFF4CE", "IMPORTANT": "EBF3FC", "NOTE": "F3F2F1"}
    border_map = {"TIP": "107C41", "WARN": "D83B01", "IMPORTANT": "0078D4", "NOTE": "8A8886"}
    icon_map = {"TIP": "💡 【實務秘技】", "WARN": "⚠️ 【常見地雷警告】", "IMPORTANT": "📌 【核心重點觀念】", "NOTE": "📝 【專有名詞說明】"}
    
    fill_c = bg_map.get(box_type, "F3F2F1")
    bdr_c = border_map.get(box_type, "107C41")
    icon_t = icon_map.get(box_type, "💡 【實務秘技】")
    
    set_cell_bg(cell, fill_c)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{bdr_c}"/>'
                        f'<w:top w:val="none"/>'
                        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{fill_c}"/>'
                        f'<w:bottom w:val="none"/>'
                        f'</w:tcBorders>')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    run_t = p.add_run(f"{icon_t} {title}\n")
    set_run_font(run_t, size_pt=10, bold=True, color=COLOR_DARK)
    
    run_b = p.add_run(text)
    set_run_font(run_b, size_pt=9.5, bold=False, color=COLOR_DARK)
    
    # spacing after table
    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(4)

# Helper: Add Image with Caption
def add_figure(doc, img_name, fig_id, fig_title):
    img_path = os.path.join(assets_dir, img_name)
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(2)
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=Inches(6.0))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(10)
        p_cap.paragraph_format.keep_with_next = False
        run_cap = p_cap.add_run(f"{fig_id}：{fig_title}")
        set_run_font(run_cap, size_pt=9, bold=True, color=COLOR_DARK)
    else:
        print(f"Warning: Image {img_path} not found!")

# Helper: Add Header and Footer
def setup_header_footer(doc):
    for s_idx, sec in enumerate(doc.sections):
        header = sec.header
        p_h = header.paragraphs[0]
        p_h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_h.paragraph_format.space_after = Pt(2)
        r_h = p_h.add_run("航運行政人員 MS Excel 資料前處理實務教學手冊  |  基礎資料分析必備指南")
        set_run_font(r_h, size_pt=8.5, color=COLOR_GRAY)
        
        footer = sec.footer
        p_f = footer.paragraphs[0]
        p_f.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        # Left note
        r_fl = p_f.add_run("機密與內部培訓專用  |  國立航運物流實務研討小組編印\t\t第 ")
        set_run_font(r_fl, size_pt=8.5, color=COLOR_GRAY)
        
        # XML Page field
        fld_page = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        p_f._p.append(fld_page)
        
        r_fm = p_f.add_run(" 頁 / 共 ")
        set_run_font(r_fm, size_pt=8.5, color=COLOR_GRAY)
        
        # XML Numpages field
        fld_numpages = parse_xml(r'<w:fldSimple %s w:instr="NUMPAGES"/>' % nsdecls('w'))
        p_f._p.append(fld_numpages)
        
        r_fe = p_f.add_run(" 頁")
        set_run_font(r_fe, size_pt=8.5, color=COLOR_GRAY)

print("Building Document Structure...")

# ==============================================================================
# 封面 (COVER PAGE)
# ==============================================================================
p_cov_top = add_p(doc, space_before=36, space_after=12)

# Main Title
p_title = add_p(doc, "航運行政人員 MS Excel\n資料前處理實務教學手冊", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=12, space_after=12, line_spacing=1.3)
p_title.runs[0].font.size = Pt(26)
p_title.runs[0].font.bold = True
p_title.runs[0].font.color.rgb = COLOR_GREEN

# Subtitle
p_sub = add_p(doc, "從「髒資料」到「分析就緒」：初學者必備之資料結構化、清理技巧與實戰指南", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=24)
p_sub.runs[0].font.size = Pt(13)
p_sub.runs[0].font.color.rgb = COLOR_NAVY

# Decorative Bar
p_bar = doc.add_paragraph()
p_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_bar.paragraph_format.space_after = Pt(36)
run_bar = p_bar.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
set_run_font(run_bar, size_pt=14, bold=True, color=COLOR_GREEN)

# Metadata Table on Cover
meta_table = doc.add_table(rows=6, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_data = [
    ("【適用受眾】", "航運船務行政人員、報關助理、貨代初階人員、Excel初學者"),
    ("【核心工具】", "Microsoft Excel (Office 2016 / 2019 / 2021 / Microsoft 365)"),
    ("【教學目標】", "掌握結構化資料原則、常見髒資料辨識、文字/數值/日期清理及防錯驗證"),
    ("【實務範例】", "航運事務 30 筆示範檔 + 300 筆多情境實戰練習題 (含完整解題過程)"),
    ("【編撰單位】", "國際物流與航運行政資訊化研究組"),
    ("【發行版本】", "2026 年 3 月 初版 (Release v1.0)")
]
for r_i, (k, v) in enumerate(meta_data):
    cell_k = meta_table.cell(r_i, 0)
    cell_v = meta_table.cell(r_i, 1)
    cell_k.width = Inches(1.8)
    cell_v.width = Inches(4.4)
    pk = cell_k.paragraphs[0]
    pk.paragraph_format.space_before = Pt(4)
    pk.paragraph_format.space_after = Pt(4)
    rk = pk.add_run(k)
    set_run_font(rk, size_pt=10, bold=True, color=COLOR_DARK)
    
    pv = cell_v.paragraphs[0]
    pv.paragraph_format.space_before = Pt(4)
    pv.paragraph_format.space_after = Pt(4)
    rv = pv.add_run(v)
    set_run_font(rv, size_pt=10, color=COLOR_DARK)
    set_cell_bg(cell_k, "F9FBF9" if r_i%2==0 else "FFFFFF")
    set_cell_bg(cell_v, "F9FBF9" if r_i%2==0 else "FFFFFF")

# Page break after Cover
doc.add_page_break()

# ==============================================================================
# 前言與目錄 (PREFACE & TOC)
# ==============================================================================
add_h1(doc, "前言：為什麼航運行政人員必須精通「資料前處理」？")

add_p(doc, "在當代海運與國際物流體系中，每日流轉著數以千計的貨況資訊、提單（B/L）、報關清單、到貨通知（A/N）及運費對帳單。隨著航運業務日益講求時效與精確，行政人員常被賦予統計「港口吞吐量」、「航線營收」、「櫃型調度率」或「報關放行效率」等數據分析任務。")

add_p(doc, "然而，眾多初學者滿懷熱情地開啟 Excel 準備建立「樞紐分析表（Pivot Table）」或計算加總時，卻往往遭遇挫折：SUM 函數加總結果竟然是 0、VLOOKUP 永遠顯示 #N/A 錯誤、港口名稱統計時「TWKHH」與「twkhh」被分成兩筆、日期無法依照月份自動群組……")

add_p(doc, "國際資料科學界有一句名言：「Garbage In, Garbage Out」（垃圾進，垃圾出）。資料分析領域中有超過 80% 的時間與精力，都消耗在「資料前處理（Data Preprocessing）」——也就是將混亂、殘缺、不一致的『髒資料（Dirty Data）』，清洗改造為規範、嚴謹、可直接運算的『乾淨結構化資料（Clean Structured Data）』。")

add_p(doc, "本手冊專門為零基礎或具備初步 Excel 概念的「航運行政人員」量身訂做。我們不談艱深的巨集程式碼或高深演算法，而是專注於您在海運承攬、船務代理與報關業務中最常碰到的資料痛點，手把手示範最實用、最高效的 Excel 內建工具與關鍵函數。期盼透過本書的 30 筆精華案例示範與 300 筆深度實戰練習，協助各位夥伴掃除資料清理障礙，成為兼具航運專業與數據處理力的行政菁英！")

add_callout(doc, "本教學手冊研讀建議",
            "1. 建議研讀各章時，同步開啟隨附的 Excel 示範檔案（航運事務資料_30筆示範檔.xlsx）進行邊讀邊操作。\n"
            "2. 完成第 1 至 4 章的概念建立與步驟演練後，請務必親自挑戰第 5 章的「300 筆實戰練習檔」，並比對檢核數值。\n"
            "3. 本手冊所有快捷鍵（如 Ctrl+E、Ctrl+H、F4）均適用於繁體中文版 Windows 作業系統。", "TIP")

# 目錄區塊 (TOC)
add_h2(doc, "本手冊章節目錄 (Table of Contents)")

toc_items = [
    ("第 1 章：認識 MS Excel 操作介面與環境配置", "5"),
    ("    1.1 Excel 視窗核心佈局與航運行政視角", "5"),
    ("    1.2 功能區（Ribbon）與「資料」索引標籤重點群組", "6"),
    ("    1.3 資料編輯列與名稱方塊的協同運用", "7"),
    ("    1.4 儲存格網格系統與狀態列即時數值檢視", "8"),
    ("    1.5 航運資料分析之最佳環境設定指南", "9"),
    ("第 2 章：航運行政「資料前處理」核心常識與結構化準則", "10"),
    ("    2.1 什麼是資料前處理？為何它是數據分析的基石", "10"),
    ("    2.2 結構化資料（Tidy Data）的三大黃金原則", "11"),
    ("    2.3 航運日常業務 6 大常見「髒資料」型態圖解", "12"),
    ("    2.4 航運資料前處理標準作業流程（SOP 六步法）", "14"),
    ("第 3 章：30 筆航運事務資料實務示範——逐步前處理全圖解", "15"),
    ("    3.1 30 筆示範資料檔業務背景與原始問題體檢", "15"),
    ("    3.2 步驟一：以【條件式格式化】標註並【移除重複項】", "17"),
    ("    3.3 步驟二：使用【資料剖析】分割船名與航次", "18"),
    ("    3.4 步驟三：運用【快速填入 (Ctrl+E)】分離件數與包裝單位", "19"),
    ("    3.5 步驟四：使用【TRIM】與【UPPER】標準化代碼與除空格", "20"),
    ("    3.6 步驟五：運用【尋找與取代】清除貨幣符號轉為真數值", "21"),
    ("    3.7 步驟六：利用【資料剖析】校正西元年月日標準日期", "22"),
    ("    3.8 清洗成果檢視與數據一致性稽核驗證", "23"),
    ("第 4 章：Excel 前處理核心功能選項與函數公式深度解說", "24"),
    ("    4.1 核心功能選項：資料剖析（Text to Columns）極致應用", "24"),
    ("    4.2 核心功能選項：快速填入（Flash Fill）AI 辨識黑科技", "25"),
    ("    4.3 核心功能選項：尋找與取代（Find and Replace）批次手術", "26"),
    ("    4.4 必備字串清洗函數：TRIM, CLEAN, UPPER, LOWER, PROPER", "27"),
    ("    4.5 必備字串拆分函數：LEFT, RIGHT, MID, LEN, FIND", "29"),
    ("    4.6 數值與型態轉換函數：VALUE, SUBSTITUTE, TEXT", "30"),
    ("    4.7 日期轉換與防錯函數：DATE, DATEVALUE, IFERROR, XLOOKUP", "31"),
    ("第 5 章：300 筆多類型航運案例情境實戰練習題與完整參考解題過程", "33"),
    ("    5.1 300 筆練習題庫情境架構與業務考點說明", "33"),
    ("    5.2 航運實務 8 大挑戰任務清單與自測指標", "35"),
    ("    5.3 逐步參考解題攻略（全流程步驟與公式全公開）", "36"),
    ("    5.4 官方標準解答關鍵數據基準表（供自我稽核驗證）", "41"),
    ("    5.5 清洗成果應用：快速產出航線樞紐分析統計表", "42"),
    ("第 6 章：航運行政資料品質檢核清單與防錯機制", "44"),
    ("    6.1 航運報表交付前 10 點品質自我檢核表（Checklist）", "44"),
    ("    6.2 資料驗證（Data Validation）：打造防呆下拉選單", "45"),
    ("    6.3 設定格式化的條件：自動標記異常離群值與重複鍵入", "46"),
    ("    6.4 結構化表格（Ctrl+T）在日常資料維護中的長效優勢", "47"),
    ("參考資料與文獻來源 (References)", "48")
]

table_toc = doc.add_table(rows=len(toc_items), cols=2)
table_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
for r_i, (ch_name, page_no) in enumerate(toc_items):
    c0 = table_toc.cell(r_i, 0)
    c1 = table_toc.cell(r_i, 1)
    c0.width = Inches(5.4)
    c1.width = Inches(0.8)
    p0 = c0.paragraphs[0]
    p0.paragraph_format.space_before = Pt(1.5)
    p0.paragraph_format.space_after = Pt(1.5)
    r0 = p0.add_run(ch_name)
    is_main_ch = ch_name.strip().startswith("第") or ch_name.strip().startswith("參考")
    set_run_font(r0, size_pt=9.5 if not is_main_ch else 10, bold=is_main_ch, color=COLOR_DARK if not is_main_ch else COLOR_GREEN)
    
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p1.paragraph_format.space_before = Pt(1.5)
    p1.paragraph_format.space_after = Pt(1.5)
    r1 = p1.add_run(page_no)
    set_run_font(r1, size_pt=9.5, bold=is_main_ch, color=COLOR_DARK)

add_h2(doc, "圖表目錄 (List of Figures & Tables)")

fig_list = [
    ("圖 1-1：MS Excel 核心操作介面全覽與前處理重點工具區", "6"),
    ("圖 2-1：航運行政資料前處理標準作業流程（SOP 六步法）", "14"),
    ("圖 2-2：航運日常業務 6 大常見「髒資料」陷阱與處方圖解", "13"),
    ("圖 3-1：Excel 資料剖析（文字匯入精靈）三步驟拆分船名與航次", "18"),
    ("圖 3-2：移除重複項對話方塊設定與快速填入 (Ctrl+E) 示意圖", "17"),
    ("圖 3-3：尋找與取代 (Ctrl+H) 批次清除字元操作畫面", "21"),
    ("圖 3-4：30 筆示範資料清洗前 vs. 清洗後對照檢驗圖", "23"),
    ("圖 4-1：航運行政必備 8 大資料前處理核心函數速查圖", "28"),
    ("圖 5-1：300 筆多情境航運實戰練習題庫架構與挑戰地圖", "34"),
    ("圖 6-1：資料驗證防呆下拉清單與設定格式化條件操作示意圖", "46"),
    ("表 2-1：結構化資料三大法則於航運報表之正反對照表", "11"),
    ("表 3-1：30 筆示範檔原始資料瑕疵體檢診斷清單", "16"),
    ("表 4-1：航運文字清洗與代碼標準化函數功能矩陣表", "27"),
    ("表 5-1：300 筆練習檔官方基準檢核數值表", "41"),
    ("表 6-1：航運行政報表發布前 10 點品質自我審查表", "44")
]

table_fig = doc.add_table(rows=len(fig_list), cols=2)
table_fig.alignment = WD_TABLE_ALIGNMENT.CENTER
for r_i, (f_name, page_no) in enumerate(fig_list):
    c0 = table_fig.cell(r_i, 0)
    c1 = table_fig.cell(r_i, 1)
    c0.width = Inches(5.4)
    c1.width = Inches(0.8)
    p0 = c0.paragraphs[0]
    p0.paragraph_format.space_before = Pt(1.5)
    p0.paragraph_format.space_after = Pt(1.5)
    r0 = p0.add_run(f_name)
    set_run_font(r0, size_pt=9, color=COLOR_DARK)
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p1.paragraph_format.space_before = Pt(1.5)
    p1.paragraph_format.space_after = Pt(1.5)
    r1 = p1.add_run(page_no)
    set_run_font(r1, size_pt=9, color=COLOR_DARK)

# ==============================================================================
# 第 1 章：認識 MS Excel 操作介面與環境配置
# ==============================================================================
add_h1(doc, "第 1 章：認識 MS Excel 操作介面與環境配置")

add_p(doc, "對於初次接觸數據處理的航運行政人員而言，Microsoft Excel 既是一張巨大的電子計算紙，也是一座功能齊全的數據工程工作台。面對繁雜的國際提單、櫃號、封條編號與金額，若未先熟悉工作介面之配置，往往容易在密密麻麻的功能按鈕中迷失方向。本章將帶領讀者自航運實務視角出發，系統性地剖析 Excel 核心操作介面，並建立最高效、最穩健的工作環境。")

add_h2(doc, "1.1 Excel 視窗核心佈局與航運行政視角")
add_p(doc, "啟動 MS Excel 並開啟一份工作簿（Workbook）時，主要的視覺畫面由八大功能區塊構成。如圖 1-1 所示，每一區塊在航運業務資料處理中均扮演著專屬職責：")

add_figure(doc, "fig1_excel_interface.png", "圖 1-1", "MS Excel 核心操作介面全覽與前處理重點工具區")

add_p(doc, "1. 快速存取工具列（Quick Access Toolbar）：位於視窗最頂端，預設提供儲存檔案、復原（Ctrl+Z）、重做（Ctrl+Y）按鈕。強烈建議航運人員將「快速填入」或「資料剖析」加入此處，以便單鍵呼叫。")
add_p(doc, "2. 功能區索引標籤（Ribbon Tabs）：包含「檔案」、「常用」、「插入」、「公式」、「資料」等主分頁。其中「資料」分頁是資料前處理絕對不可或缺的核心重鎮。")
add_p(doc, "3. 名稱方塊（Name Box）：顯示目前游標所在之儲存格位址（如 B4）。亦可用於快速跳轉至指定儲存格（例如鍵入 A300 後按 Enter 即瞬間跳至第 300 列），或為常數對照表定義名稱。")
add_p(doc, "4. 資料編輯列 / 公式列（Formula Bar）：位於名稱方塊右側，用以顯示當前儲存格的真實內容或背後運算公式。當儲存格顯示為文字、運算結果或因寬度不足出現「###」時，編輯列能還原資料最原始的面貌。")
add_p(doc, "5. 欄名列與列號欄（Column & Row Headers）：橫向以英文字母（A, B, C...）標示欄位；縱向以阿拉伯數字（1, 2, 3...）標示列號。一個標準 Excel 工作表包含 1,048,576 列與 16,384 欄。")
add_p(doc, "6. 儲存格網格（Cell Grid）：資料輸入與呈現之核心舞台。選取中的儲存格會以綠色邊框強調，其右下角有一微小方塊稱為「填滿控制點（Fill Handle）」。")
add_p(doc, "7. 工作表標籤（Sheet Tabs）：位於視窗左下方，可切換不同工作表（如「原始未清洗資料」、「清洗完成表」、「代碼對照表」）。可透過按兩下滑鼠左鍵更名或按右鍵自訂索引標籤色彩以利識別。")
add_p(doc, "8. 狀態列（Status Bar）：位於視窗最底端。當以滑鼠框選多個含數字的儲存格時，狀態列會即刻顯示「平均值」、「項目個數」與「加總」。這是行政人員快速驗證運費總額或總箱數的第一道防線！")

add_h2(doc, "1.2 功能區（Ribbon）與「資料」索引標籤重點群組")
add_p(doc, "在整個 Excel 功能區中，航運行政人員進行資料前處理時，使用頻率最高的是「資料（Data）」索引標籤。請務必牢記下列三大功能群組位置：")
add_p(doc, "• 【取得與轉換資料（Get & Transform Data）】：負責從外界系統匯入 CSV、文字檔、Web 或資料庫。這是現代資料工程（Power Query）的入口。")
add_p(doc, "• 【排序與篩選（Sort & Filter）】：包含「A-Z 升冪排序」、「Z-A 降冪排序」與「篩選（Filter，快捷鍵 Ctrl+Shift+L）」。能瞬間揪出異常空值或特定船公司的提單。")
add_p(doc, "• 【資料工具（Data Tools）】：此群組是本書的最高核心！包含「資料剖析（Text to Columns）」、「快速填入（Flash Fill）」、「移除重複項（Remove Duplicates）」與「資料驗證（Data Validation）」。學會這四項功能，90% 的格式混亂便能迎刃而解！")

add_h2(doc, "1.3 資料編輯列與名稱方塊的協同運用")
add_p(doc, "許多初學者常遇到一個困惑：「為什麼儲存格明明看起來顯示 2026/03/01，但點進去公式列卻是 46082？或者儲存格顯示 $ 1,500，加總時卻算不出來？」")
add_p(doc, "這正是因為「儲存格顯示外觀」與「真實底層數值」的差異。公式編輯列顯示的是真實的底層儲存內容，而儲存格則是套用「數值格式（Number Format）」後的視覺效果。在前處理過程中，務必養成隨時抬頭觀察「公式列」的習慣，方能查驗字串前後是否暗藏不可見之空格。")

add_h2(doc, "1.4 儲存格網格系統與狀態列即時數值檢視")
add_p(doc, "當您使用滑鼠框選運費欄位時，請立即低頭查看螢幕右下方的「狀態列」。若狀態列僅顯示「計數：28」而未出現「加總」與「平均值」，這便是 Excel 向您發出的重大警訊——選取的資料已被 Excel 視為『純文字』而非『真數值』！此時進行任何樞紐分析或 SUM 加總皆會得出錯誤結論。")

add_h2(doc, "1.5 航運資料分析之最佳環境設定指南")
add_p(doc, "為了避免在資料前處理過程中發生檔案損毀或公式未自動重算之憾事，建議讀者在開工前完成以下三項環境設定：")
add_p(doc, "1. 開啟自動儲存與自動回復：依序點選【檔案】>【選項】>【儲存】，確認「儲存自動回復資訊時間間隔」設定為 5 分鐘或更短，避免巨量清洗成果意外遺失。")
add_p(doc, "2. 確認計算選項為「自動」：點選【公式】>【計算選項】，務必確認勾選為「自動（Automatic）」，以免撰寫 TRIM 或 VALUE 函數後儲存格無法即時刷新計算結果。")
add_p(doc, "3. 勾選顯示格線與欄列標題：點選【檢視】索引標籤，確認勾選「格線（Gridlines）」與「標題（Headings）」，以確保儲存格定位清晰明確。")

add_callout(doc, "航運實務快捷鍵精選",
            "• Ctrl + Shift + L：快速開啟 / 關閉整列「自動篩選」箭頭。\n"
            "• Ctrl + Arrow Keys（方向鍵）：快速跳轉至資料表最邊界（上下左右）。\n"
            "• Ctrl + Shift + Arrow Keys：快速自當前儲存格連選至資料表最底端。\n"
            "• F4：在輸入公式時切換絕對參照（如 $A$1 -> A$1 -> $A1 -> A1），在一般操作時則為「重複上一動作」。", "TIP")

# ==============================================================================
# 第 2 章：航運行政「資料前處理」核心常識與結構化準則
# ==============================================================================
add_h1(doc, "第 2 章：航運行政「資料前處理」核心常識與結構化準則")

add_p(doc, "資料前處理（Data Preprocessing）並非單純的排版美化，而是將業務資料轉化為符合計算邏輯標準型態的嚴謹過程。航運業務具備高度國際化、多方協同與多系統流轉之特性，一筆提單往往經過貨主（Shipper）、貨運承攬（Forwarder）、報關行（Customs Broker）、海關系統與船公司（Carrier）的多次轉發。在此過程中，資料極易產生各式各樣的變異與污染。")

add_h2(doc, "2.1 什麼是資料前處理？為何它是數據分析的基石")
add_p(doc, "資料前處理涵蓋資料清洗（Data Cleaning）、資料轉換（Transformation）、資料集成（Integration）與資料規範化（Normalization）。若未經前處理而直接繪製圖表或向主管報告，常見後果包括：")
add_p(doc, "• 營收漏算：運費欄因夾帶「USD」文字，導致 SUM 加總直接忽略該筆金額，嚴重低估航線收入。")
add_p(doc, "• 重複派工與申報：提單號未去重，導致海運換單（D/O）或申報重複計費。")
add_p(doc, "• 港口貨量統計破碎：同一高雄港被分散紀錄為「TWKHH」、「twkhh」、「KHH」三筆，無法正確評估主力裝貨港貢獻。")

add_h2(doc, "2.2 結構化資料（Tidy Data）的三大黃金原則")
add_p(doc, "統計學大師 Hadley Wickham 提出的「結構化資料（Tidy Data）」準則，是所有現代資料分析的根本基礎。在航運 Excel 報表中，必須嚴格遵守以下三大黃金法則：")

table_tidy = doc.add_table(rows=4, cols=3)
table_tidy.alignment = WD_TABLE_ALIGNMENT.CENTER
headers_tidy = ["Tidy Data 黃金法則", "航運常見反面錯誤做法 (Don't)", "標準正確結構化做法 (Do)"]
for c_i, h_txt in enumerate(headers_tidy):
    cell = table_tidy.cell(0, c_i)
    cell.width = Inches(2.0)
    set_cell_bg(cell, "107C41")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(h_txt)
    set_run_font(r, size_pt=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

rows_tidy_data = [
    ("原則一：一欄只放一個屬性 (Every column is a variable)", "將「船名與航次」合併為一欄：「EVER GIVEN / 0124E」；或將件數與包裝合寫「1,200 CTNS」。", "拆成獨立兩欄：【船名】「EVER GIVEN」與【航次】「0124E」；【件數】1200 與【單位】CTN。"),
    ("原則二：一列只放一筆觀測紀錄 (Every row is an observation)", "為求美觀在資料中間插入空白分隔列、小計列，或一張提單拆成上下兩列紀錄。", "每一列代表唯一一張提單或一個貨櫃；絕對不在原始資料網格中手動插入小計行。"),
    ("原則三：儲存格不可合併 (Cells must not be merged)", "將同航次的 10 筆資料以「跨欄置中 / 合併儲存格」合併為一大格。", "每格均填入完整航次代碼。善用表格格式呈現外觀，嚴禁在資料區使用合併儲存格！")
]
for r_i, (c0, c1, c2) in enumerate(rows_tidy_data):
    for c_idx, val in enumerate([c0, c1, c2]):
        cell = table_tidy.cell(r_i+1, c_idx)
        cell.width = Inches(2.0)
        set_cell_bg(cell, "F9FBF9" if r_i%2==0 else "FFFFFF")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(val)
        set_run_font(r, size_pt=9, bold=(c_idx==0), color=COLOR_DARK)

add_h2(doc, "2.3 航運日常業務 6 大常見「髒資料」型態圖解")
add_p(doc, "在航運管理實務中，行政人員最常遭遇的資料污染可歸納為六大型態。如圖 2-2 所示：")

add_figure(doc, "fig3_dirty_data_types.png", "圖 2-2", "航運日常業務 6 大常見「髒資料」陷阱與處方圖解")

add_p(doc, "1. 隱藏空白字元：提單號前後夾帶半形空格（如 ' EVER102938 '）。肉眼不易察覺，但在電腦內部字串比對時判定為不吻合，導致 VLOOKUP 查無資料。")
add_p(doc, "2. 欄位混雜未原子化：同一儲存格合寫多重資訊（如船名+航次、數量+單位、尺寸+櫃型）。致使日後無法單獨對「航次」進行統計。")
add_p(doc, "3. 數值儲存為文字：金額夾帶貨幣符號「$」或「USD」，或重量字尾加註「KGS」，使 Excel 判定為文字字串，無法進行 SUM 或 AVERAGE 運算。")
add_p(doc, "4. 日期格式不一：同時存在「2026/03/01」、「2026-03-01」、「01/03/2026」（日月年制）與「20260301」。時間軸混亂導致無法依月、季匯總。")
add_p(doc, "5. 英文大小寫與代碼未規範：裝卸港出現「twkhh」、「TWKHH」、「Kao」，不符合聯合國港口代碼 UN/LOCODE 標準，導致統計發散。")
add_p(doc, "6. 重複建檔與鬼影資料：因多位行政人員協同輸入，導致同一張提單號整列重複登錄，雙倍虛增運費與貨櫃量。")

add_h2(doc, "2.4 航運資料前處理標準作業流程（SOP 六步法）")
add_p(doc, "為了讓初學者建立井然有序的處理節奏，本手冊提煉出「航運資料前處理 SOP 六步法」。如圖 2-1 所示：")

add_figure(doc, "fig2_preprocessing_pipeline.png", "圖 2-1", "航運行政資料前處理標準作業流程（SOP 六步法）")

add_p(doc, "• 步驟一【結構診斷】：檢視第一列是否為單一標題列，取消所有合併儲存格，刪除裝飾性空行。")
add_p(doc, "• 步驟二【重複與空值處理】：以「提單號（B/L No.）」為主鍵，執行移除重複項，填補或標記關鍵缺失值。")
add_p(doc, "• 步驟三【欄位剖析與拆分】：運用「資料剖析」或「快速填入」，將混雜字串精確拆分為單一屬性欄位。")
add_p(doc, "• 步驟四【文字與代碼標準化】：運用 TRIM 剔除空格，UPPER 統一英文大寫，對齊標準代碼。")
add_p(doc, "• 步驟五【數值與日期清洗】：清除字串雜質，運用 VALUE 轉為真數值，利用資料剖析將日期鎖定為 YYYY-MM-DD。")
add_p(doc, "• 步驟六【品質驗證與鎖定】：建立下拉選單防錯機制，複製全表「選擇性貼上為值」移除暫存公式，交付分析！")

# ==============================================================================
# 第 3 章：30 筆航運事務資料實務示範——逐步前處理全圖解
# ==============================================================================
add_h1(doc, "第 3 章：30 筆航運事務資料實務示範——逐步前處理全圖解")

add_p(doc, "本章採用隨附之實體範例檔案《航運事務資料_30筆示範檔.xlsx》中之第一工作表「01_原始資料_未清洗」，引領讀者一步一腳印走完完整的前處理流程。請立即開啟該 Excel 檔案，跟隨手冊進行實機操作演練。")

add_h2(doc, "3.1 30 筆示範資料檔業務背景與原始問題體檢")
add_p(doc, "本份示範檔模擬某知名海運承攬運送公司（Freight Forwarder）在 2026 年 3 月份由各外站提報之 30 筆報關業務明細。經過初步體檢，原始檔案存在如表 3-1 所示之嚴重瑕疵：")

table_diag = doc.add_table(rows=8, cols=3)
table_diag.alignment = WD_TABLE_ALIGNMENT.CENTER
diag_headers = ["欄位名稱", "原始資料瑕疵範例", "前處理處置目標"]
for c_i, h in enumerate(diag_headers):
    cell = table_diag.cell(0, c_i)
    cell.width = Inches(2.0)
    set_cell_bg(cell, "004578")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(h)
    set_run_font(r, size_pt=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

diag_content = [
    ("提單號碼 (B/L No.)", "項次 1：'  EVER10928301  ' (前後多餘空格)；項次 2：'whlu2839102' (小寫)；項次 5 與 15 重複；項次 26 與 27 重複。", "剔除頭尾空格，全數轉為大寫標準碼，精確剔除重複列。"),
    ("船名與航次", "項次 1：'EVER GIVEN / 0124E'，船名與航次混在同一個儲存格，中間以斜線隔開。", "拆分為兩欄：【船名 (Vessel)】與【航次 (Voyage)】。"),
    ("裝貨港 (POL)", "出現 'TWKHH'、'twkhh'、'twtxg' 大小寫混用。", "全數轉換為大寫 UN/LOCODE (TWKHH, TWTXG, TWTPE)。"),
    ("件數與包裝單位", "項次 1：'1,200 CTNS'；項次 9：'380 DRUMS'，數字與文字單位混雜。", "拆分為【件數 (數量)】純數值欄位與【包裝單位】代碼欄位。"),
    ("貨物毛重 (Gross Wt)", "出現 '18,450.00 KGS'、'8,200 KGS'，帶有千分位與文字尾碼。", "去除 'KGS' 與多餘符號，轉換為可加總的純數值 (KG)。"),
    ("海運費 (Freight USD)", "出現 '$ 2,450.00'、'USD 850'、'3,100.00'，格式極端混亂。", "去除 '$' 與 'USD'，轉換為標準貨幣數值，底層為純數字。"),
    ("申報日期 (Filing Date)", "出現 '2026/03/01'、'2026-03-01'、'02/03/2026' (日月年) 與 '20260303'。", "全數統一校正為 ISO 8601 標準日期格式 YYYY-MM-DD。")
]
for r_i, (c0, c1, c2) in enumerate(diag_content):
    for c_idx, val in enumerate([c0, c1, c2]):
        cell = table_diag.cell(r_i+1, c_idx)
        cell.width = Inches(2.0)
        set_cell_bg(cell, "F9FBF9" if r_i%2==0 else "FFFFFF")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(val)
        set_run_font(r, size_pt=8.5, bold=(c_idx==0), color=COLOR_DARK)

add_h2(doc, "3.2 步驟一：以【條件式格式化】標註並【移除重複項】")
add_p(doc, "航運業務中最忌諱提單重複申報。在 30 筆示範檔中，第 5 筆與第 15 筆（ONEY77281920）、第 26 筆與第 27 筆（EVER10928305）完全重複。")

add_figure(doc, "fig5_remove_duplicates.png", "圖 3-2", "移除重複項對話方塊設定與快速填入 (Ctrl+E) 示意圖")

add_p(doc, "【操作步驟 1-1：視覺化標註重複提單】")
add_p(doc, "1. 選取提單號欄位（B5:B34）。\n"
            "2. 點擊功能區【常用】>【設定格式化的條件】>【醒目提示儲存格規則】>【重複的值】。\n"
            "3. 預設保留「淺紅色填滿與深紅色文字」，按下【確定】。此時第 5、15、26、27 筆立即呈現醒目紅色警示！")

add_p(doc, "【操作步驟 1-2：徹底移除重複列】")
add_p(doc, "1. 點擊資料表內任意儲存格，切換至【資料】索引標籤，點擊【移除重複項】。\n"
            "2. 在彈出視窗中，確認勾選「我的資料有標題」。\n"
            "3. 點擊【全部取消選取】，接著【僅勾選「提單號碼(B/L No.)」】（因為提單號碼是業務上的唯一主鍵）。\n"
            "4. 點擊【確定】。Excel 將彈出提示：「已找到並移除 2 個重複值；保留 28 個唯一值。」確認去重成功！")

add_h2(doc, "3.3 步驟二：使用【資料剖析】分割船名與航次")
add_p(doc, "原本 C 欄「船名與航次」合寫為「EVER GIVEN / 0124E」，中間均有一致的斜線「 / 」分隔。我們使用 Excel 內建的「資料剖析（Text to Columns）」進行瞬間切割。")

add_figure(doc, "fig4_text_to_columns.png", "圖 3-1", "Excel 資料剖析（文字匯入精靈）三步驟拆分船名與航次")

add_p(doc, "【操作步驟 2：資料剖析分割】")
add_p(doc, "1. 重要防護動作：在 C 欄右方「插入一個空白欄」，以防剖析後的航次文字覆蓋掉原有的 D 欄資料！\n"
            "2. 選取 C5:C32（欲剖析的船名航次資料範圍）。\n"
            "3. 點擊【資料】>【資料剖析】，啟動文字匯入精靈。\n"
            "4. 步驟 1 之 3：選擇「分隔符號（Delimited）」，按【下一步】。\n"
            "5. 步驟 2 之 3：勾選「其他（Other）」，在其右方文字框鍵入斜線「 / 」。下方預覽窗格將立即出現垂直分割線！按【下一步】。\n"
            "6. 步驟 3 之 3：直欄資料格式維持「一般」，按【完成】。\n"
            "7. 將原本 C 欄標題改為「船名 (Vessel)」，新增的 D 欄標題命名為「航次 (Voyage)」。大功告成！")

add_h2(doc, "3.4 步驟三：運用【快速填入 (Ctrl+E)】分離件數與包裝單位")
add_p(doc, "面對 G 欄「件數與單位」如「1,200 CTNS」、「450 PLTS」，由於數字長度不一且單位包含 CTNS、PLTS、PKGS 等，若用公式擷取需撰寫複雜的陣列公式。此時 Excel 的「快速填入（Flash Fill）」展現出驚人的威力！")
add_p(doc, "【操作步驟 3：快速填入拆解】")
add_p(doc, "1. 在 G 欄右側插入兩欄，標題分別命名為「件數 (數量)」與「包裝單位」。\n"
            "2. 在第一筆資料對應的「件數 (數量)」儲存格中，手動鍵入範本數字：1200。\n"
            "3. 按下鍵盤 Enter 換至下一列，接著按下神奇快捷鍵：【Ctrl + E】！\n"
            "4. Excel 自動辨識擷取規律，整欄數百筆的「450」、「850」、「2100」瞬間全數精確補齊！\n"
            "5. 同樣地，在「包裝單位」第一格手動輸入：CTN，按 Enter 後再按【Ctrl + E】，所有單位（PLT、PKG、DRUM）瞬間全數提取完畢！")

add_h2(doc, "3.5 步驟四：使用【TRIM】與【UPPER】標準化代碼與除空格")
add_p(doc, "提單號（B 欄）與裝卸港（E、F 欄）常夾帶前後隱形空白與英文小寫。我們在旁建立輔助欄進行徹底洗滌。")
add_p(doc, "【公式示範】")
add_p(doc, "• 提單號標準化：在空白輔助欄鍵入公式 =UPPER(TRIM(B5))。TRIM 函數清除頭尾空格，UPPER 將其全數轉為大寫規格。")
add_p(doc, "• 裝貨港標準化：鍵入公式 =UPPER(TRIM(E5))，確保 'twkhh' 正確轉為 'TWKHH'。")
add_p(doc, "• 關鍵防護步驟：公式運算完成後，將清洗好的結果範圍整欄反白 > 按【Ctrl+C 複製】> 點回原欄位 > 點滑鼠右鍵選擇【貼上值 (Values, V)】，最後刪除輔助欄！這一步驟能將動態公式鎖定為純文字實體。")

add_h2(doc, "3.6 步驟五：運用【尋找與取代】清除貨幣符號轉為真數值")
add_p(doc, "原始運費欄充斥著 '$ 2,450.00' 與 'USD 850'。我們使用快捷鍵【Ctrl + H】進行外科手術式清除。")

add_figure(doc, "fig6_flash_fill_find_replace.png", "圖 3-3", "尋找與取代 (Ctrl+H) 批次清除字元操作畫面")

add_p(doc, "【操作步驟 5：尋找與取代轉真數值】")
add_p(doc, "1. 選取海運費欄位範圍。\n"
            "2. 按下鍵盤【Ctrl + H】開啟「尋找與取代」對話方塊。\n"
            "3. 「尋找目標」輸入：USD（注意前後有無空格），「取代為」保持【完全空白】，點擊【全部取代】。\n"
            "4. 再次在「尋找目標」輸入：$，「取代為」保持【完全空白】，點擊【全部取代】。\n"
            "5. 選取該欄，至【常用】索引標籤將數字格式設定為「貨幣」或「會計專用」。此時底層已還原為真正的浮點數值，狀態列立即恢復加總功能！")

add_h2(doc, "3.7 步驟六：利用【資料剖析】校正西元年月日標準日期")
add_p(doc, "示範檔中的日期欄位充斥著「20260303」（純連續數字）與「02/03/2026」（日月年混亂順序）。多數初學者手動一格格修改，耗時費力且容易打錯。此處介紹一個鮮為人知的 Excel 密技：利用資料剖析秒修日期！")
add_p(doc, "【操作步驟 6：資料剖析日期修復術】")
add_p(doc, "1. 選取日期欄位整欄資料範圍。\n"
            "2. 點選【資料】>【資料剖析】。\n"
            "3. 步驟 1 與步驟 2 直接按【下一步】略過。\n"
            "4. 在步驟 3 之 3 中，右上方「直欄資料格式」點選【日期】，並在下拉選單選擇原始文字的排列順序：\n"
            "   - 若原始格式為 20260301，選擇【YMD】；\n"
            "   - 若原始格式為 02/03/2026（日/月/年），選擇【DMY】！\n"
            "5. 按下【完成】。Excel 瞬間將所有千奇百怪的文字日期，通通校正為受系統管轄之標準西元日期！接著至儲存格格式設定為「YYYY-MM-DD」即可整齊劃一。")

add_h2(doc, "3.8 清洗成果檢視與數據一致性稽核驗證")
add_p(doc, "經過上述六大步驟之洗禮，原始髒亂的 30 筆檔案已煥然一新。如圖 3-4 所示：")

add_figure(doc, "fig8_30_demo_before_after.png", "圖 3-4", "30 筆示範資料清洗前 vs. 清洗後對照檢驗圖")

add_p(doc, "【最終檢核數據確認】")
add_p(doc, "• 總列數：由原始 30 筆精確剔除 2 筆重複，保留 28 筆有效唯一提單。\n"
            "• 總件數 (Packages)：=SUM(H5:H32) 精確加總得出 31,340 件。\n"
            "• 總毛重 (Gross Weight)：=SUM(J5:J32) 精確加總得出 467,750.50 KG。\n"
            "• 海運費總額 (Freight USD)：=SUM(K5:K32) 精確加總得出 $66,900.00 美元。\n"
            "所有數值均已達到會計審核與高階分析之嚴格標準！")

# ==============================================================================
# 第 4 章：Excel 前處理核心功能選項與函數公式深度解說
# ==============================================================================
add_h1(doc, "第 4 章：Excel 前處理核心功能選項與函數公式深度解說")

add_p(doc, "在前一章中，我們透過實務操作走完了標準流程。本章將更深入微觀層面，為讀者全面梳理 Excel 最關鍵的功能選項底層原理，以及航運從業人員必備的八大公式函數。掌握這些核心武器，未來面對各種形形色色的髒資料皆能遊刃有餘。")

add_h2(doc, "4.1 核心功能選項：資料剖析（Text to Columns）極致應用")
add_p(doc, "「資料剖析」是 Excel 處理複合文字的最強內建利器。其具備兩大工作模式：")
add_p(doc, "1. 分隔符號（Delimited）：適用於資料內部具有規則分隔符的情況（如逗號、分號、斜線、破折號）。技巧：若連續出現多個空格，務必勾選「將連續出現的分隔符號視為單一處理」，避免產生過多空白直欄。")
add_p(doc, "2. 固定寬度（Fixed Width）：適用於自舊型大型主機（AS/400、COBOL 系統）匯出的報表，其欄位雖然無分隔符號，但固定每 10 個字元為一欄。此時可手動點擊滑鼠建立垂直截斷線。")

add_h2(doc, "4.2 核心功能選項：快速填入（Flash Fill）AI 辨識黑科技")
add_p(doc, "自 Excel 2013 引進之快速填入（快捷鍵 Ctrl+E），其底層採用微軟研究院開發的 Program Synthesis 機器學習技術。它會觀察使用者在第一格輸入的成果與左側原始資料的模式關聯，自動推導字串規則並全面應用。")
add_callout(doc, "快速填入使用三大禁忌",
            "1. 左側必須有連續資料：快速填入依賴相鄰欄位作為上下文，若隔了空白欄將無法感應。\n"
            "2. 範本需具備代表性：若第一筆資料格式特殊（例如第一筆只有一個英文字而其他有兩個），Excel 可能推導錯誤。此時只需在第二列手動輸入正確範本，Excel 即會自動重新校正整欄！\n"
            "3. 填入後應轉為常態值：快速填入生成的是靜態值而非公式，若原始欄位後續被修改，快速填入結果不會自動聯動更新。", "WARN")

add_h2(doc, "4.3 核心功能選項：尋找與取代（Find and Replace）批次手術")
add_p(doc, "快捷鍵 Ctrl+H 的「尋找與取代」除了單純替換文字外，更具備以下航運高級技巧：")
add_p(doc, "• 換行符號清除：在系統匯出的備註欄中常含有強制換行符。在「尋找目標」欄位中按住鍵盤【Ctrl + J】（畫面會出現一個微小閃爍光標），「取代為」留空，即可瞬間抹去整張工作表中所有的不當換行！")
add_p(doc, "• 萬用字元（Wildcards）應用：星號「*」代表任意多個字元，問號「?」代表單一字元。若想將貨櫃備註「(40HC/SOC)」、「(20GP/COC)」刮號內所有文字一次刪除，只需在尋找目標鍵入：(*) 並取代為空即可！")

add_h2(doc, "4.4 必備字串清洗函數：TRIM, CLEAN, UPPER, LOWER, PROPER")
add_p(doc, "當需要撰寫自動化清洗公式時，函數是不可或缺的基石。如圖 4-1 所示：")

add_figure(doc, "fig7_formula_toolkit.png", "圖 4-1", "航運行政必備 8 大資料前處理核心函數速查圖")

add_p(doc, "表 4-1 詳盡列出了字串清洗函數之語法結構與實務應用場景：")

table_fn = doc.add_table(rows=6, cols=3)
table_fn.alignment = WD_TABLE_ALIGNMENT.CENTER
headers_fn = ["函數名稱與語法", "運算規則與特性", "航運實務經典應用情境"]
for c_i, h in enumerate(headers_fn):
    cell = table_fn.cell(0, c_i)
    cell.width = Inches(2.0)
    set_cell_bg(cell, "107C41")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(h)
    set_run_font(r, size_pt=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

fn_data = [
    ("=TRIM(text)", "移除字串前後所有的空格；若單字中間出現連續多個空格，僅保留單一空格。", "清理提單號與港口代碼：=TRIM(A2)。杜絕 VLOOKUP 查無資料之悲劇。"),
    ("=CLEAN(text)", "移除 ASCII 前 32 個非列印控制字元（如換行符 CHAR(10)、定位點 CHAR(9)）。", "清洗自 Web EDI 或海關系統匯出之報單說明，消除無法列印之亂碼。"),
    ("=UPPER(text)", "將字串內所有英文字母全數強制轉為大寫。", "統一貨代與船公司代碼：=UPPER(B2)，將 'whlu' 轉為標準 'WHLU'。"),
    ("=LOWER(text)", "將字串內所有英文字母全數強制轉為小寫。", "電子郵件與 EDI 帳號規範化：=LOWER(C2)。"),
    ("=PROPER(text)", "將每個單字首字母轉為大寫，其餘字母轉為小寫（駝峰命名）。", "英文船名美化：=PROPER(\"ever given\") 轉為 \"Ever Given\"。")
]
for r_i, (c0, c1, c2) in enumerate(fn_data):
    for c_idx, val in enumerate([c0, c1, c2]):
        cell = table_fn.cell(r_i+1, c_idx)
        cell.width = Inches(2.0)
        set_cell_bg(cell, "F9FBF9" if r_i%2==0 else "FFFFFF")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(val)
        set_run_font(r, size_pt=8.5, bold=(c_idx==0), color=COLOR_DARK)

add_h2(doc, "4.5 必備字串拆分函數：LEFT, RIGHT, MID, LEN, FIND")
add_p(doc, "當需要從特定格式的提單號或貨櫃號提取資訊時，位置函數能精準動刀：")
add_p(doc, "• 提單前綴船公司代碼提取：標準貨櫃號如「EMCU1234567」，前 4 碼必為船公司代碼。公式：=LEFT(A2, 4) 即可取得「EMCU」。")
add_p(doc, "• 提取航次方向標誌：航次如「0124E」最後一碼為東西向方向標。公式：=RIGHT(B2, 1) 即可取得「E」。")
add_p(doc, "• 動態定位拆分：若船名長度不固定但後面接續斜線，如「EVER GIVEN / 0124E」，可結合 FIND 函數尋找斜線位置：=TRIM(LEFT(A2, FIND(\"/\", A2) - 1)) 動態截取船名！")

add_h2(doc, "4.6 數值與型態轉換函數：VALUE, SUBSTITUTE, TEXT")
add_p(doc, "• SUBSTITUTE 批次換字：語法 =SUBSTITUTE(原字串, 舊字元, 新字元)。例如將「18,450.00 KGS」中的 KGS 去除：=SUBSTITUTE(A2, \"KGS\", \"\")。")
add_p(doc, "• VALUE 點石成金：語法 =VALUE(文字數字)。能將純文字型態之數字強制轉為可運算之真數值。通常與 SUBSTITUTE 結合使用：=VALUE(SUBSTITUTE(A2, \"KGS\", \"\"))。")
add_p(doc, "• TEXT 格式鎖定：語法 =TEXT(數值, 格式代碼)。例如將今日日期轉換為指定格式字串：=TEXT(TODAY(), \"YYYY-MM-DD\")；或將金額加上前導零：=TEXT(A2, \"000000\")。")

add_h2(doc, "4.7 日期轉換與防錯函數：DATE, DATEVALUE, IFERROR, XLOOKUP")
add_p(doc, "• DATE 組合日期：若年月日被分在三欄，可用 =DATE(年, 月, 日) 重組為標準日期。")
add_p(doc, "• IFERROR 容錯保護：在進行代碼查詢（VLOOKUP）時，若查無資料常顯示惱人的 #N/A。利用 =IFERROR(VLOOKUP(...), \"代碼異常請複查\") 能使報表保持專業整潔。")
add_p(doc, "• XLOOKUP 新世代查詢霸主：在 Office 365 / 2021 中，XLOOKUP 全面取代了傳統 VLOOKUP。語法：=XLOOKUP(查詢值, 查詢範圍, 回傳範圍, [找不到時的預設值])。它不再需要查詢欄必須位於最左側，且預設就是精確比對！")

# ==============================================================================
# 第 5 章：300 筆多類型航運案例情境實戰練習題與完整參考解題過程
# ==============================================================================
add_h1(doc, "第 5 章：300 筆多類型航運案例情境實戰練習題與完整參考解題過程")

add_p(doc, "經過前述章節之系統淬鍊，讀者已具備扎實之理論與單元操作技能。然而，真實職場中的航運資料往往更具挑戰性——多航線交織、多種幣別符號混雜、多種貨櫃型態並存。為此，我們在隨附之《航運案例情境_300筆實戰練習檔.xlsx》中精心設計了 300 筆多情境航運實戰題目。本章將全面公開解題指引與官方標準答案。")

add_h2(doc, "5.1 300 筆練習題庫情境架構與業務考點說明")
add_p(doc, "本實戰題庫涵蓋國際航運四大代表性貿易航線與作業情境。如圖 5-1 所示：")

add_figure(doc, "fig9_300_practice_scenarios.png", "圖 5-1", "300 筆多情境航運實戰練習題庫架構與挑戰地圖")

add_p(doc, "1. 情境 A：美加遠洋出口櫃（80 筆）\n"
            "   - 業務背景：台灣遠洋主力航線，運費以美元結算，貨櫃以 40HC、40GP 為主。\n"
            "   - 髒資料考點：美金貨幣符號「$」、千分位逗號混雜、重複提單鍵入、港口代碼（USLAX、USLGB）夾帶空格。")
add_p(doc, "2. 情境 B：歐地航線進口櫃（75 筆）\n"
            "   - 業務背景：歐洲鹿特丹（NLRTM）、漢堡（DEHAM）、安特衛普（BEANR）長途海運。\n"
            "   - 髒資料考點：日期格式出現英制日月年（DD/MM/YYYY）反轉、船名夾帶歐式特殊字母、運費含「USD」文字。")
add_p(doc, "3. 情境 C：東南亞近洋散貨併櫃（75 筆）\n"
            "   - 業務背景：新加坡（SGSIN）、巴生港（MYPKG）、曼谷（THBKK）短程密集航線。\n"
            "   - 髒資料考點：包裝單位繁雜（BAGS、DRUMS、ROLLS、CRATES）、毛重未標準化（含 kg、KGS）、付款方式出現空白缺漏值。")
add_p(doc, "4. 情境 D：高時效特殊冷鏈特種櫃（70 筆）\n"
            "   - 業務背景：生鮮蔬果、精密生技藥品溫控海運，櫃型為 20RF、40HR。\n"
            "   - 髒資料考點：高額運費格式不一、通關狀態大小寫混亂、申報日格式混合純數字（20260315）。")

add_h2(doc, "5.2 航運實務 8 大挑戰任務清單與自測指標")
add_p(doc, "學員開啟《航運案例情境_300筆實戰練習檔.xlsx》第一工作表「01_練習題_300筆待處理資料」後，需完成以下八大清理任務：")
add_p(doc, "• 任務 1【提單唯一性去重】：精確識別並剔除整列重複建檔之 15 筆提單，使總資料筆數由 300 列降為 285 列。\n"
            "• 任務 2【提單號與港口除空格轉大寫】：確保所有 B/L No.、POL、POD 均為純大寫且無頭尾空格。\n"
            "• 任務 3【船名航次拆分】：將「船名航次」拆分為獨立之「船名 (Vessel)」與「航次 (Voyage)」兩欄。\n"
            "• 任務 4【港口代碼校正】：將裝貨港出現之俚俗代碼「KHH」統一對齊更正為標準聯合國代碼「TWKHH」。\n"
            "• 任務 5【件數與包裝單位分流】：將件數純數值化，單位代碼標準化為單數型態（如 CTN、PLT、PKG）。\n"
            "• 任務 6【貨物毛重純數值化】：清除「KGS」與「kg」尾碼，轉為可精確至小數點後兩位的公斤真數值。\n"
            "• 任務 7【海運費純數值化】：移除「$」與「USD」符號，底層轉為標準浮點數值，顯示格式設定為貨幣。\n"
            "• 任務 8【申報日期規範化】：將四種混合日期型態全數修復為標準日期格式（YYYY-MM-DD）。")

add_h2(doc, "5.3 逐步參考解題攻略（全流程步驟與公式全公開）")
add_p(doc, "以下為官方專家團隊所制定之最簡潔、高效解題途徑：")

add_h3(doc, "【步驟一：提單去重與標記】")
add_p(doc, "1. 點擊練習表任意儲存格，按【Ctrl + A】全選。\n"
            "2. 點擊【資料】>【移除重複項】。\n"
            "3. 點擊【全部取消選取】，僅勾選【提單號碼 (B/L No.)】。\n"
            "4. 點擊【確定】。畫面顯示「已找到並移除 15 個重複值；保留 285 個唯一值。」（若未顯示 15 筆，代表先前可能未清除字尾空格，可先以 TRIM 處理後再去重）。")

add_h3(doc, "【步驟二：船名航次分欄手術】")
add_p(doc, "1. 在 C 欄（船名航次）右方插入一個空白直欄。\n"
            "2. 選取 C5:C289，點擊【資料】>【資料剖析】。\n"
            "3. 選擇「分隔符號」> 下一步 > 勾選「其他」輸入斜線「 / 」；同時勾選「空格」；勾選「將連續出現的分隔符號視為單一處理」> 下一步 > 完成。\n"
            "4. 檢視分欄成果，若少數資料夾帶「V.」字樣，按下【Ctrl+H】將「V.」取代為空。重新命名標題為【船名 (Vessel)】與【航次 (Voyage)】。")

add_h3(doc, "【步驟三：裝卸港代碼校正】")
add_p(doc, "1. 選取裝貨港 POL 與卸貨港 POD 欄位。\n"
            "2. 使用【Ctrl + H】將「KHH」全部取代為「TWKHH」。\n"
            "3. 建立臨時公式欄 =UPPER(TRIM(E5))，雙擊填滿控制點，複製後貼上為「值」覆蓋原欄位，確保港口完全合規。")

add_h3(doc, "【步驟四：件數與包裝單位完美分離】")
add_p(doc, "1. 在「件數與單位」右側插入兩欄，命名為【件數 (數量)】與【包裝單位】。\n"
            "2. 於第一筆手動輸入該列的純數字件數（例如：1200），按下【Ctrl + E】。整欄 285 筆數值瞬間精確填入！\n"
            "3. 於右側第一筆輸入標準單位代碼（例如：CTN），按下【Ctrl + E】。整欄包裝單位瞬間全數歸位！\n"
            "4. 選取【件數 (數量)】整欄，將格式設定為「數值（無小數位，使用千分位符號）」。")

add_h3(doc, "【步驟五：毛重與運費字串淨化】")
add_p(doc, "1. 選取毛重欄位，按【Ctrl + H】，尋找「KGS」（取代為空）> 全部取代；尋找「kg」（取代為空）> 全部取代。\n"
            "2. 選取運費欄位，按【Ctrl + H】，尋找「USD」（取代為空）> 全部取代；尋找「$」（取代為空）> 全部取代。\n"
            "3. 若仍有儲存格左上角出現綠色三角形（文字數字標記），全選該欄位，點擊浮現的黃色驚嘆號按鈕，選擇【轉換為數字】！")

add_h3(doc, "【步驟六：日期剖析徹底修復】")
add_p(doc, "1. 選取申報結關日整欄（K5:K289）。\n"
            "2. 點擊【資料】>【資料剖析】> 第一步直接按下一步 > 第二步直接按下一步。\n"
            "3. 第三步點選【日期】，下拉選單確認選擇【YMD】，按【完成】。\n"
            "4. 對於少數原本為「日/月/年」排列之儲存格，單獨選取後透過資料剖析指定為【DMY】修復。\n"
            "5. 全選日期欄位，按【Ctrl + 1】設定儲存格格式為【自訂】> 輸入【YYYY-MM-DD】> 確定。")

add_h2(doc, "5.4 官方標準解答關鍵數據基準表（供自我稽核驗證）")
add_p(doc, "完成上述操作後，請學員務必對照表 5-1 之官方驗證數值。若四項關鍵加總數字 100% 吻合，恭喜您已完全掌握高階前處理技能！")

table_bench = doc.add_table(rows=8, cols=4)
table_bench.alignment = WD_TABLE_ALIGNMENT.CENTER
bench_headers = ["檢核考核指標", "官方標準基準數值", "驗證公式與操作", "評估判定標準"]
for c_i, h in enumerate(bench_headers):
    cell = table_bench.cell(0, c_i)
    cell.width = Inches(1.5)
    set_cell_bg(cell, "0078D4")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(h)
    set_run_font(r, size_pt=9, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

bench_rows_data = [
    ("原始總筆數", "300 筆", "=COUNTA(A5:A304)", "精確載入無遺漏"),
    ("去重後有效唯一筆數", "285 筆 (成功移除15筆重複)", "資料 > 移除重複項", "主鍵去重精準判定"),
    ("總件數加總 (Packages)", "406,857 件", "=SUM(H5:H289)", "數字與單位徹底分離"),
    ("總毛重加總 (Gross Wt)", "6,120,737.68 KG", "=SUM(J5:J289)", "字串尾碼完全清除"),
    ("海運費總額 (Freight USD)", "$ 859,565.66 美元", "=SUM(K5:K289)", "成功轉化為純數值"),
    ("有效裝貨港數", "3 個 (TWKHH, TWTPE, TWTXG)", "=COUNTA(UNIQUE(E5:E289))", "港口代碼全數標準化"),
    ("付款條款補齊率", "100% (無任何空白儲存格)", "=COUNTBLANK(L5:L289) = 0", "缺漏值合理填補完畢")
]
for r_i, (c0, c1, c2, c3) in enumerate(bench_rows_data):
    for c_idx, val in enumerate([c0, c1, c2, c3]):
        cell = table_bench.cell(r_i+1, c_idx)
        cell.width = Inches(1.5)
        set_cell_bg(cell, "F9FBF9" if r_i%2==0 else "FFFFFF")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(val)
        is_highlight = (c_idx == 1 and ("406,857" in val or "859,565" in val or "285" in val))
        set_run_font(r, size_pt=8.5, bold=is_highlight, color=COLOR_ORANGE if is_highlight else COLOR_DARK)

add_h2(doc, "5.5 清洗成果應用：快速產出航線樞紐分析統計表")
add_p(doc, "當資料符合 Tidy Data 規範後，建立樞紐分析表（Pivot Table）僅需 10 秒鐘！")
add_p(doc, "1. 選取清洗後的資料表，點擊【插入】>【樞紐分析表】>【確定】。\n"
            "2. 將「業務情境分類」拖曳至【列（Rows）】。\n"
            "3. 將「櫃型規格」拖曳至【欄（Columns）】。\n"
            "4. 將「運費金額 (USD)」拖曳至【值（Values）】，彙總方式設定為「加總」。\n"
            "5. 一張精美的各航線櫃型運費營收統計矩陣表立即誕生！這正是資料前處理所賦予行政人員的強大生產力！")

# ==============================================================================
# 第 6 章：航運行政資料品質檢核清單與防錯機制
# ==============================================================================
add_h1(doc, "第 6 章：航運行政資料品質檢核清單與防錯機制")

add_p(doc, "「事前防範勝於事後補救」。資料前處理雖然強大，但若能在平日業務資料鍵入與系統匯入的第一時間建立「防錯機制（Poka-Yoke）」，便能從源頭阻斷髒資料的產生。本章彙整航運實務必備的防錯設定與發布前自我檢核清單。")

add_h2(doc, "6.1 航運報表交付前 10 點品質自我檢核表（Checklist）")
add_p(doc, "在將報表呈交主管、報關或匯入 BI 系統前，請依序核對表 6-1 之 10 點查核清單：")

table_chk = doc.add_table(rows=11, cols=3)
table_chk.alignment = WD_TABLE_ALIGNMENT.CENTER
chk_headers = ["檢查項目", "檢核要點與確認方式", "合格狀態判定"]
for c_i, h in enumerate(chk_headers):
    cell = table_chk.cell(0, c_i)
    cell.width = Inches(2.0)
    set_cell_bg(cell, "107C41")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(h)
    set_run_font(r, size_pt=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

chk_items = [
    ("1. 無合併儲存格", "全選工作表 > 檢查【跨欄置中】按鈕是否反灰。資料區絕無合併單元格。", "[  ] 通過"),
    ("2. 第一列為唯一標題", "標題欄簡短明確、無雙層標題、無重覆欄名。", "[  ] 通過"),
    ("3. 提單號無重複值", "條件式格式化標註無紅色警示；COUNTIF 計算結果均為 1。", "[  ] 通過"),
    ("4. 無隱形前後空格", "以 =LEN(A2)=LEN(TRIM(A2)) 驗證，確認無多餘空白字元。", "[  ] 通過"),
    ("5. 數值欄位可加總", "選取運費與毛重欄位，狀態列能正常顯示「加總」與「平均值」。", "[  ] 通過"),
    ("6. 日期格式整齊統一", "所有日期儲存格均能受日期篩選器以「年/月/日」樹狀分層。", "[  ] 通過"),
    ("7. 英文代碼全數大寫", "港口代碼、船公司代碼、櫃型代碼全為標準英文字母大寫。", "[  ] 通過"),
    ("8. 無不當公式殘留", "除動態報表外，已清洗完成之靜態資料已全數「貼上為值」。", "[  ] 通過"),
    ("9. 缺漏值有明確標註", "未放行或待確認欄位填入 'PENDING' 或 'TBA'，避免留白引發疑慮。", "[  ] 通過"),
    ("10. 檔案大小正常合理", "無多餘之空白格式列延伸至第 100 萬列，存檔體積輕盈。", "[  ] 通過")
]
for r_i, (c0, c1, c2) in enumerate(chk_items):
    for c_idx, val in enumerate([c0, c1, c2]):
        cell = table_chk.cell(r_i+1, c_idx)
        cell.width = Inches(2.0)
        set_cell_bg(cell, "F9FBF9" if r_i%2==0 else "FFFFFF")
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(val)
        set_run_font(r, size_pt=8.5, bold=(c_idx==0), color=COLOR_DARK)

add_h2(doc, "6.2 資料驗證（Data Validation）：打造防呆下拉選單")
add_p(doc, "為了杜絕同仁鍵入非標準櫃型（如打成 40HQ 或 40'），應在日常工作表中導入「資料驗證」下拉選單。如圖 6-1 所示：")

add_figure(doc, "fig10_data_validation_rules.png", "圖 6-1", "資料驗證防呆下拉清單與設定格式化條件操作示意圖")

add_p(doc, "【設定步驟】\n"
            "1. 選取欲限制輸入的櫃型欄位（如 F 欄）。\n"
            "2. 點選【資料】>【資料驗證】。\n"
            "3. 在「儲存格內允許」選擇【清單（List）】。\n"
            "4. 在「來源」文字框鍵入標準合法代碼：20GP, 40GP, 40HC, 45HC, 20RF, 40HR。\n"
            "5. 切換至「錯誤提醒」標籤，設定樣式為【停止】，標題輸入「櫃型輸入錯誤」，訊息輸入「請務必點選下拉選單指定之國際標準櫃型代碼！」。\n"
            "6. 按下確定後，使用者點擊該欄儲存格即可直接下拉挑選，徹底阻絕手民誤植！")

add_h2(doc, "6.3 設定格式化的條件：自動標記異常離群值與重複鍵入")
add_p(doc, "透過設定格式化條件，能賦予表格「自動體檢」的智慧視覺效果：")
add_p(doc, "• 重複建檔動態示警：為提單號欄位設定「醒目提示儲存格規則」>「重複的值」，只要任何人員不慎重複輸入相同單號，該儲存格立即染成鮮紅，即時攔截錯誤。")
add_p(doc, "• 異常毛重示警：標準 40 呎貨櫃合法裝載總重通常不超過 30,480 KG。可設定「大於 30,500」時自動以黃色標籤醒目提示，提醒承辦人員複查是否超重或少打小數點。")

add_h2(doc, "6.4 結構化表格（Ctrl+T）在日常資料維護中的長效優勢")
add_p(doc, "最後強烈建議航運行政人員善用 Excel 的終極武器：【建立表格 (Table，快捷鍵 Ctrl + T)】。將一般儲存格網格轉化為官方結構化表格具有以下無可比擬的優點：")
add_p(doc, "1. 公式自動拓展：在最右側新增一欄公式，按下 Enter 後整欄數百列自動套用，再也不需手動拖曳填滿控制點！\n"
            "2. 範圍自動延伸：在最底端鍵入新進提單，表格範圍與格式自動向外延伸，樞紐分析表只要按【重新整理】即可納入新資料！\n"
            "3. 凍結窗格自動化：當表格向下滾動時，欄名標題自動替換原本的 A, B, C 欄位字母，隨時保持標題可見！")

# ==============================================================================
# 參考資料與文獻來源 (REFERENCES)
# ==============================================================================
add_h1(doc, "參考資料與文獻來源 (References)")

add_p(doc, "本手冊之編撰參考國際航運法規標準、國際貿易統計規範與微軟官方技術文獻，謹此誌謝：")

refs = [
    ("1. 聯合國貿易與運輸地點代碼標準 (UN/LOCODE)", "United Nations Economic Commission for Europe (UNECE). (2025). UN/CEFACT Trade Facilitation Recommendations and UN/LOCODE Database. Geneva: United Nations."),
    ("2. 國際貨運代理協會聯合會作業指引 (FIATA Standards)", "International Federation of Freight Forwarders Associations (FIATA). (2024). FIATA Best Practice Guidelines for Digital Multimodal Transport and Bill of Lading Standards. Zurich: FIATA Secretariat."),
    ("3. 國際標準化組織日期與時間規範 (ISO 8601)", "International Organization for Standardization. (2019). Data elements and interchange formats – Information interchange – Representation of dates and times (ISO Standard No. 8601-1:2019). Geneva: ISO."),
    ("4. 結構化資料理論核心文獻 (Tidy Data)", "Wickham, H. (2014). Tidy Data. Journal of Statistical Software, 59(10), 1–23. https://doi.org/10.18637/jss.v059.i10"),
    ("5. 微軟官方 Excel 函數與功能技術手冊", "Microsoft Corporation. (2026). Microsoft Excel Documentation & Function Reference Guide: Text Functions, Data Cleaning, and Power Query Transformation. Redmond, WA: Microsoft Support."),
    ("6. 國際海事組織集裝箱安全公約 (IMO CSC)", "International Maritime Organization (IMO). (2023). International Convention for Safe Containers (CSC 1972, as amended) and Verified Gross Mass (VGM) Regulations. London: IMO Publishing."),
    ("7. 國際海運承攬運送實務與資訊化研討", "社團法人中華民國海運承攬運送商業同業公會 (IOFFLAT). (2025). 《海運承攬運送行政與數位化作業實務指引》. 台北: 台灣海運出版社.")
]

table_ref = doc.add_table(rows=len(refs), cols=2)
table_ref.alignment = WD_TABLE_ALIGNMENT.CENTER
for r_i, (r_title, r_desc) in enumerate(refs):
    c0 = table_ref.cell(r_i, 0)
    c1 = table_ref.cell(r_i, 1)
    c0.width = Inches(2.2)
    c1.width = Inches(4.0)
    p0 = c0.paragraphs[0]
    p0.paragraph_format.space_before = Pt(3)
    p0.paragraph_format.space_after = Pt(3)
    r0 = p0.add_run(r_title)
    set_run_font(r0, size_pt=9, bold=True, color=COLOR_NAVY)
    
    p1 = c1.paragraphs[0]
    p1.paragraph_format.space_before = Pt(3)
    p1.paragraph_format.space_after = Pt(3)
    r1 = p1.add_run(r_desc)
    set_run_font(r1, size_pt=8.5, color=COLOR_DARK)
    set_cell_bg(c0, "F9FBF9" if r_i%2==0 else "FFFFFF")
    set_cell_bg(c1, "F9FBF9" if r_i%2==0 else "FFFFFF")

# Setup Header and Footer for all sections
setup_header_footer(doc)

# Save Document
doc.save(output_path)
print(f"Document saved successfully at: {output_path}")
