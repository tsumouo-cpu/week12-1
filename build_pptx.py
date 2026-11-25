# -*- coding: utf-8 -*-
import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

output_path = "C:/proj/week12-1/航運行政人員_Excel資料前處理實務簡報.pptx"
assets_dir = "C:/proj/week12-1/assets"

prs = Presentation()
# 16:9 Widescreen layout
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_slide_layout = prs.slide_layouts[6] # Blank slide

# Color Palette (Maritime & Modern Dark Executive Theme)
C_BG_DARK = RGBColor(15, 23, 42)       # Slate 900
C_CARD_BG = RGBColor(30, 41, 59)      # Slate 800
C_CARD_BORDER = RGBColor(51, 65, 85)  # Slate 700
C_EMERALD = RGBColor(16, 185, 129)    # Emerald 500
C_EMERALD_DARK = RGBColor(16, 124, 65)# Excel Green
C_CYAN = RGBColor(6, 182, 212)        # Cyan 500
C_BLUE = RGBColor(59, 130, 246)       # Blue 500
C_ROSE = RGBColor(244, 63, 94)        # Rose 500
C_AMBER = RGBColor(245, 158, 11)      # Amber 500
C_PURPLE = RGBColor(168, 85, 247)     # Purple 500
C_WHITE = RGBColor(255, 255, 255)
C_SLATE_200 = RGBColor(226, 232, 240)
C_SLATE_400 = RGBColor(148, 163, 184)

FONT_MAIN = "微軟正黑體"
FONT_ENG = "Segoe UI"

# Helper: Set slide background color
def set_slide_background(slide, color=C_BG_DARK):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg

# Helper: Add header banner to content slides
def add_slide_header(slide, tag_text, title_text, slide_num):
    # Tag
    tb_tag = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8), Inches(0.3))
    tf_tag = tb_tag.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = tag_text.upper()
    p_tag.font.name = FONT_ENG
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = C_EMERALD

    # Title
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(10), Inches(0.6))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = FONT_MAIN
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = C_WHITE

    # Slide Number & Watermark
    tb_num = slide.shapes.add_textbox(Inches(11.2), Inches(0.45), Inches(1.3), Inches(0.5))
    tf_num = tb_num.text_frame
    p_num = tf_num.paragraphs[0]
    p_num.alignment = PP_ALIGN.RIGHT
    p_num.text = f"{slide_num:02d} / 12"
    p_num.font.name = FONT_ENG
    p_num.font.size = Pt(13)
    p_num.font.bold = True
    p_num.font.color.rgb = C_EMERALD

    # Dividing Line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.3), Inches(11.733), Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = C_CARD_BORDER
    line.line.fill.background()

# Helper: Add Card Box
def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_CARD_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
    else:
        card.line.fill.background()
    return card

print("Generating 12 PPTX Slides...")

# ==============================================================================
# SLIDE 1: Cover Page
# ==============================================================================
s1 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s1, C_BG_DARK)

# Badge
badge = add_card(s1, Inches(4.5), Inches(1.2), Inches(4.333), Inches(0.45), bg_color=RGBColor(6, 78, 59), border_color=C_EMERALD)
tf_b = badge.text_frame
p_b = tf_b.paragraphs[0]
p_b.alignment = PP_ALIGN.CENTER
p_b.text = "🚢 國際航運與船務行政資訊化實戰系列"
p_b.font.name = FONT_MAIN
p_b.font.size = Pt(12)
p_b.font.bold = True
p_b.font.color.rgb = C_EMERALD

# Title
tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(1.8))
tf1 = tb1.text_frame
tf1.word_wrap = True
p1_1 = tf1.paragraphs[0]
p1_1.alignment = PP_ALIGN.CENTER
p1_1.text = "航運行政人員 MS Excel"
p1_1.font.name = FONT_MAIN
p1_1.font.size = Pt(36)
p1_1.font.bold = True
p1_1.font.color.rgb = C_WHITE

p1_2 = tf1.add_paragraph()
p1_2.alignment = PP_ALIGN.CENTER
p1_2.text = "資料前處理實務手冊"
p1_2.font.name = FONT_MAIN
p1_2.font.size = Pt(40)
p1_2.font.bold = True
p1_2.font.color.rgb = C_EMERALD

# Subtitle
tb_sub = s1.shapes.add_textbox(Inches(1.5), Inches(3.7), Inches(10.333), Inches(0.8))
tf_sub = tb_sub.text_frame
tf_sub.word_wrap = True
p_sub = tf_sub.paragraphs[0]
p_sub.alignment = PP_ALIGN.CENTER
p_sub.text = "從「髒資料」到「分析就緒」：結構化準則、文字數值清理技巧與 300 筆多情境實戰指南"
p_sub.font.name = FONT_MAIN
p_sub.font.size = Pt(15)
p_sub.font.color.rgb = C_SLATE_200

# 4 Info Badges
cards_data = [
    ("🎯 適用對象", "船務行政 / 報關專員 / 貨代初學者", C_WHITE),
    ("🛠️ 核心工具", "Microsoft Excel 全版本", C_EMERALD),
    ("📁 實務案例", "30 筆示範檔 + 300 筆實戰檔", C_CYAN),
    ("⚡ 核心技能", "欄位剖析 / 提單去重 / 型態轉換", C_AMBER)
]
for idx, (title, desc, val_color) in enumerate(cards_data):
    cx = Inches(1.2 + idx * 2.8)
    card = add_card(s1, cx, Inches(4.8), Inches(2.6), Inches(1.3))
    tf_c = card.text_frame
    tf_c.word_wrap = True
    p_t = tf_c.paragraphs[0]
    p_t.text = title
    p_t.font.name = FONT_MAIN
    p_t.font.size = Pt(11)
    p_t.font.bold = True
    p_t.font.color.rgb = C_SLATE_400
    
    p_d = tf_c.add_paragraph()
    p_d.text = desc
    p_d.font.name = FONT_MAIN
    p_d.font.size = Pt(12)
    p_d.font.bold = True
    p_d.font.color.rgb = val_color

# Footer note
tb_foot = s1.shapes.add_textbox(Inches(1.0), Inches(6.6), Inches(11.333), Inches(0.4))
tf_foot = tb_foot.text_frame
p_f = tf_foot.paragraphs[0]
p_f.alignment = PP_ALIGN.CENTER
p_f.text = "國立航運物流實務研討小組 編印  |  2026 年初版 (Release v1.0)"
p_f.font.name = FONT_MAIN
p_f.font.size = Pt(10)
p_f.font.color.rgb = C_SLATE_400

# ==============================================================================
# SLIDE 2: 痛點與緣起 (Garbage In, Garbage Out)
# ==============================================================================
s2 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s2, C_BG_DARK)
add_slide_header(s2, "Problem & Motivation", "為什麼 80% 的分析時間花在「整理資料」？", 2)

# Left Box: The Pain Points
card_left = add_card(s2, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2), bg_color=RGBColor(38, 16, 26), border_color=RGBColor(159, 18, 57))
tf_l = card_left.text_frame
tf_l.word_wrap = True
p_lt = tf_l.paragraphs[0]
p_lt.text = "❌ 航運新手常遭遇的「致命挫折」"
p_lt.font.name = FONT_MAIN
p_lt.font.size = Pt(16)
p_lt.font.bold = True
p_lt.font.color.rgb = C_ROSE

items_l = [
    ("SUM 加總算出來竟然是 0：", "金額欄夾帶「$」或「USD」，被 Excel 識別為字串！"),
    ("VLOOKUP 永遠顯示 #N/A 錯誤：", "提單號肉眼看似相同，頭尾其實暗藏看不見的半形空格！"),
    ("港口貨量統計破碎發散：", "高雄港分別被鍵入為「TWKHH」、「twkhh」、「KHH」，無法整合！"),
    ("重複建檔虛增雙倍帳款：", "協同作業導致同一提單整列重複登錄，報關與運費嚴重高估！")
]
for t, d in items_l:
    p = tf_l.add_paragraph()
    p.text = f"• {t}"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    p.space_before = Pt(8)
    
    p2 = tf_l.add_paragraph()
    p2.text = f"  {d}"
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(11)
    p2.font.color.rgb = C_SLATE_200

# Bottom alert
p_la = tf_l.add_paragraph()
p_la.text = "\n💡 資料科學公理：Garbage In, Garbage Out（垃圾進，垃圾出）"
p_la.font.name = FONT_MAIN
p_la.font.size = Pt(11)
p_la.font.bold = True
p_la.font.color.rgb = C_AMBER

# Right Box: The Benefits
card_right = add_card(s2, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2), bg_color=RGBColor(6, 44, 33), border_color=RGBColor(5, 150, 105))
tf_r = card_right.text_frame
tf_r.word_wrap = True
p_rt = tf_r.paragraphs[0]
p_rt.text = "✅ 掌握「資料前處理」的關鍵回報"
p_rt.font.name = FONT_MAIN
p_rt.font.size = Pt(16)
p_rt.font.bold = True
p_rt.font.color.rgb = C_EMERALD

items_r = [
    ("標準結構化（Tidy Data）：", "落實一欄一屬性，資料無縫對接樞紐分析表與 Power BI。"),
    ("告別手動一格格修改：", "活用「資料剖析」與「快速填入 (Ctrl+E)」，千筆資料瞬間分欄！"),
    ("建立自動防呆把關：", "透過「資料驗證」下拉清單與格式化條件，從源頭阻斷打錯字。"),
    ("主管與會計 100% 信任：", "報表數據嚴謹清晰、加總分毫不差，展現航運行政專業度！")
]
for t, d in items_r:
    p = tf_r.add_paragraph()
    p.text = f"• {t}"
    p.font.name = FONT_MAIN
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    p.space_before = Pt(8)
    
    p2 = tf_r.add_paragraph()
    p2.text = f"  {d}"
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(11)
    p2.font.color.rgb = C_SLATE_200

p_ra = tf_r.add_paragraph()
p_ra.text = "\n🌟 前處理不是排版美化，而是將資料轉化為可運算資產的核心技術！"
p_ra.font.name = FONT_MAIN
p_ra.font.size = Pt(11)
p_ra.font.bold = True
p_ra.font.color.rgb = C_CYAN

# ==============================================================================
# SLIDE 3: 認識 Excel 操作介面與重點工具區
# ==============================================================================
s3 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s3, C_BG_DARK)
add_slide_header(s3, "Excel UI Tour", "認識 Excel 操作介面與航運前處理重點工具區", 3)

# 4 UI Components on Top Row
ui_items_top = [
    ("1. 功能區 (Ribbon)", "【資料】分頁為前處理核心大本營，必用資料工具與篩選。", C_EMERALD),
    ("2. 名稱方塊 (Name Box)", "顯示目前游標儲存格（如 B4），亦可鍵入代碼瞬間跳轉。", C_BLUE),
    ("3. 公式列 (Formula Bar)", "還原底層真實數據與公式，揪出前後空格與格式偽裝。", C_PURPLE),
    ("4. 狀態列 (Status Bar)", "框選數值即時顯示加總/平均；未顯示加總即代表誤判為文字！", C_AMBER)
]
for i, (title, desc, col) in enumerate(ui_items_top):
    card = add_card(s3, Inches(0.8 + i * 2.95), Inches(1.6), Inches(2.85), Inches(2.3))
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = FONT_MAIN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = col
    
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(11)
    p2.font.color.rgb = C_SLATE_200
    p2.space_before = Pt(8)

# Bottom Row: 3 Focus Areas
ui_items_bot = [
    ("📌 核心群組：資料工具 (Data Tools)", "包含「資料剖析」、「快速填入」、「移除重複項」、「資料驗證」四劍客。解決 90% 航運格式混亂！", C_EMERALD),
    ("📌 核心群組：排序與篩選 (Sort & Filter)", "快捷鍵 Ctrl + Shift + L 快速啟動自動篩選器，瞬間揪出異常空值、未放行提單與特定船公司明細。", C_CYAN),
    ("📌 核心功能：設定格式化的條件 (Conditional Formatting)", "位於【常用】分頁。能設定「重複的值」高亮標註，讓重複輸入的提單號立即以醒目紅色警示。", C_ROSE)
]
for i, (title, desc, col) in enumerate(ui_items_bot):
    card = add_card(s3, Inches(0.8 + i * 3.95), Inches(4.2), Inches(3.85), Inches(2.6))
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = FONT_MAIN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = col
    
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(11)
    p2.font.color.rgb = C_SLATE_200
    p2.space_before = Pt(8)

# ==============================================================================
# SLIDE 4: 結構化資料三大黃金原則 (Tidy Data)
# ==============================================================================
s4 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s4, C_BG_DARK)
add_slide_header(s4, "Tidy Data Principles", "結構化資料（Tidy Data）的三大黃金原則", 4)

tidy_rules = [
    ("原則一：一欄只放一個屬性", "Every column is a single variable.",
     "❌ 錯誤：合寫「EVER GIVEN / 0124E」或「1,200 CTNS」\n\n✅ 正確：拆分為【船名】與【航次】；【件數】與【包裝單位】兩欄",
     "效益：可單獨對特定航次或包裝單位進行統計匯總與篩選。", C_EMERALD),
    
    ("原則二：一列只放一筆紀錄", "Every row is an observation.",
     "❌ 錯誤：為求美觀在資料中間插入空白分隔列、或手動加小計行\n\n✅ 正確：每一列代表唯一一張提單或貨櫃，連續紀錄無空行",
     "效益：確保排序、自動篩選與公式向下拖曳時不會中途斷裂。", C_BLUE),
    
    ("原則三：儲存格絕對不合併", "Cells must never be merged.",
     "❌ 錯誤：將同航次的 10 筆資料以「跨欄置中 / 合併儲存格」合併\n\n✅ 正確：每格皆填入完整航次，嚴禁在資料區使用合併儲存格！",
     "效益：合併儲存格是樞紐分析與 VLOOKUP 運算失敗的萬惡之源！", C_PURPLE)
]

for idx, (title, sub, body, benefit, col) in enumerate(tidy_rules):
    card = add_card(s4, Inches(0.8 + idx * 3.95), Inches(1.6), Inches(3.85), Inches(5.2))
    tf = card.text_frame
    tf.word_wrap = True
    
    p_num = tf.paragraphs[0]
    p_num.text = f"RULE 0{idx+1}"
    p_num.font.name = FONT_ENG
    p_num.font.size = Pt(11)
    p_num.font.bold = True
    p_num.font.color.rgb = col
    
    p_t = tf.add_paragraph()
    p_t.text = title
    p_t.font.name = FONT_MAIN
    p_t.font.size = Pt(16)
    p_t.font.bold = True
    p_t.font.color.rgb = C_WHITE
    p_t.space_before = Pt(4)
    
    p_sub = tf.add_paragraph()
    p_sub.text = sub
    p_sub.font.name = FONT_ENG
    p_sub.font.size = Pt(10)
    p_sub.font.color.rgb = C_SLATE_400
    
    p_b = tf.add_paragraph()
    p_b.text = f"\n{body}"
    p_b.font.name = FONT_MAIN
    p_b.font.size = Pt(11)
    p_b.font.color.rgb = C_SLATE_200
    
    p_bf = tf.add_paragraph()
    p_bf.text = f"\n💡 {benefit}"
    p_bf.font.name = FONT_MAIN
    p_bf.font.size = Pt(10)
    p_bf.font.bold = True
    p_bf.font.color.rgb = col

# ==============================================================================
# SLIDE 5: 航運 6 大常見「髒資料」型態
# ==============================================================================
s5 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s5, C_BG_DARK)
add_slide_header(s5, "Dirty Data Taxonomy", "航運日常業務 6 大常見「髒資料」型態與處方", 5)

dirty_items = [
    ("1. 隱藏空白字元", "「  EVER102938  」", "=TRIM() 清除頭尾空格", "提單號肉眼看似相同但比對失敗，VLOOKUP 永遠回傳 #N/A 錯誤。"),
    ("2. 欄位未原子化", "「EVER GIVEN / 0124E」", "資料剖析 / Ctrl+E 拆分", "船名與航次黏在一起，無法單獨統計個別航次貨櫃量。"),
    ("3. 數值夾帶文字單位", "「1,200 CTNS」/「$ 2,450」", "=SUBSTITUTE() / VALUE()", "帶有單位或貨幣符號被判定為文字，無法以 SUM 加總運算。"),
    ("4. 日期格式混亂", "2026/03/01 與 01/03/2026", "資料剖析 (YMD/DMY)", "美制英制混雜，導致時間軸錯亂，無法依月份/季分析。"),
    ("5. 港口代碼大小寫", "twkhh vs TWKHH vs KHH", "=UPPER() / UN/LOCODE", "大小寫未統一，樞紐分析時同一高雄港被分散成三筆。"),
    ("6. 重複建檔與鬼影", "同一提單號整列重複登錄", "資料 > 移除重複項", "協同作業失誤，導致重複計費、雙倍虛增運費與毛重。")
]

for idx, (title, sample, remedy, impact) in enumerate(dirty_items):
    row = idx // 3
    col = idx % 3
    cx = Inches(0.8 + col * 3.95)
    cy = Inches(1.6 + row * 2.65)
    
    card = add_card(s5, cx, cy, Inches(3.85), Inches(2.45))
    tf = card.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = FONT_MAIN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    
    p_s = tf.add_paragraph()
    p_s.text = f"案例：{sample}"
    p_s.font.name = FONT_ENG
    p_s.font.size = Pt(10)
    p_s.font.bold = True
    p_s.font.color.rgb = C_ROSE
    p_s.space_before = Pt(2)
    
    p_r = tf.add_paragraph()
    p_r.text = f"處方：{remedy}"
    p_r.font.name = FONT_MAIN
    p_r.font.size = Pt(10)
    p_r.font.bold = True
    p_r.font.color.rgb = C_EMERALD
    
    p_i = tf.add_paragraph()
    p_i.text = f"影響：{impact}"
    p_i.font.name = FONT_MAIN
    p_i.font.size = Pt(10)
    p_i.font.color.rgb = C_SLATE_200
    p_i.space_before = Pt(4)

# ==============================================================================
# SLIDE 6: 前處理標準作業流程 (SOP 六步法)
# ==============================================================================
s6 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s6, C_BG_DARK)
add_slide_header(s6, "Standard Operating Procedure", "航運資料前處理標準作業流程（SOP 六步法）", 6)

sop_steps = [
    ("STEP 01", "結構診斷與環境清理", ["• 取消所有合併儲存格", "• 刪除裝飾性空行與小計", "• 確認第 1 列為單一欄位標題"], C_EMERALD),
    ("STEP 02", "提單去重與空值查核", ["• 格式化條件高亮重複提單", "• 【移除重複項】僅勾選提單號", "• 標記缺失欄位為 'TBA'"], C_CYAN),
    ("STEP 03", "欄位剖析與屬性拆分", ["• 預先插入空白欄防覆蓋", "• 【資料剖析】依斜線拆分船機", "• 【Ctrl+E】分離件數與包裝"], C_BLUE),
    ("STEP 04", "文字淨化與代碼大寫", ["• =TRIM() 剔除頭尾多餘空格", "• =UPPER() 統一英文代碼大寫", "• 校正 KHH 為標準 TWKHH"], C_PURPLE),
    ("STEP 05", "數值與日期型態校正", ["• 【Ctrl+H】移除 USD、$、KGS", "• 確保數值欄能即時 SUM 加總", "• 資料剖析將文字秒轉標準日期"], C_AMBER),
    ("STEP 06", "品質驗證與鎖定交付", ["• 複製全表【選擇性貼上為值】", "• 建立櫃型資料驗證下拉選單", "• 產出對帳指標並匯入樞紐分析"], C_EMERALD)
]

for idx, (step_no, title, bullet_list, col) in enumerate(sop_steps):
    row = idx // 3
    col_idx = idx % 3
    cx = Inches(0.8 + col_idx * 3.95)
    cy = Inches(1.6 + row * 2.65)
    
    card = add_card(s6, cx, cy, Inches(3.85), Inches(2.45))
    tf = card.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = step_no
    p.font.name = FONT_ENG
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = col
    
    p_t = tf.add_paragraph()
    p_t.text = title
    p_t.font.name = FONT_MAIN
    p_t.font.size = Pt(14)
    p_t.font.bold = True
    p_t.font.color.rgb = C_WHITE
    p_t.space_before = Pt(2)
    
    for b in bullet_list:
        p_b = tf.add_paragraph()
        p_b.text = b
        p_b.font.name = FONT_MAIN
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = C_SLATE_200
        p_b.space_before = Pt(3)

# ==============================================================================
# SLIDE 7: 30 筆示範檔逐步實戰演練 (Before vs. After)
# ==============================================================================
s7 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s7, C_BG_DARK)
add_slide_header(s7, "Hands-on 30 Demo Walkthrough", "30 筆示範檔：清洗前 (Dirty) vs 清洗後 (Clean) 對照", 7)

# Left Column: Before
card_b = add_card(s7, Inches(0.8), Inches(1.6), Inches(5.7), Inches(4.2), bg_color=RGBColor(38, 16, 26), border_color=RGBColor(159, 18, 57))
tf_b = card_b.text_frame
tf_b.word_wrap = True
p_bt = tf_b.paragraphs[0]
p_bt.text = "🔴 【清洗前】原始髒資料（無法直接統計分析）"
p_bt.font.name = FONT_MAIN
p_bt.font.size = Pt(14)
p_bt.font.bold = True
p_bt.font.color.rgb = C_ROSE

items_dirty_samples = [
    "1. 提單號帶大量空格與小寫：'  EVER10928301  '、'whlu2839102'",
    "2. 船名與航次黏在同格：'EVER GIVEN / 0124E'",
    "3. 件數欄混雜英文單位：'1,200 CTNS' (文字型態)",
    "4. 毛重帶 KGS 尾碼：'18,450.00 KGS' (無法參與加總)",
    "5. 運費欄夾帶幣別符號：'$ 2,450.00'、'USD 850'",
    "6. 日期格式四種混用：2026/03/01、02/03/2026、20260303",
    "7. 存在整列重複提單：第 5 筆與第 15 筆（ONEY77281920）完全重複"
]
for item in items_dirty_samples:
    p = tf_b.add_paragraph()
    p.text = item
    p.font.name = FONT_MAIN
    p.font.size = Pt(10)
    p.font.color.rgb = C_SLATE_200
    p.space_before = Pt(4)

# Right Column: After
card_a = add_card(s7, Inches(6.8), Inches(1.6), Inches(5.7), Inches(4.2), bg_color=RGBColor(6, 44, 33), border_color=RGBColor(5, 150, 105))
tf_a = card_a.text_frame
tf_a.word_wrap = True
p_at = tf_a.paragraphs[0]
p_at.text = "🟢 【清洗後】標準結構化資料（分析就緒）"
p_at.font.name = FONT_MAIN
p_at.font.size = Pt(14)
p_at.font.bold = True
p_at.font.color.rgb = C_EMERALD

items_clean_samples = [
    "1. 提單號無任何頭尾空格，全數轉為合規大寫標準碼",
    "2. 船名 (Vessel) 與 航次 (Voyage) 獨立成兩欄",
    "3. 件數拆分為【純數值: 1200】與【包裝單位: CTN】",
    "4. 毛重轉化為純數值 18450.00 KG，可加總與平均",
    "5. 運費轉化為純數值 2450.00，設定貨幣顯示格式",
    "6. 日期全數校正為 ISO 8601 標準日期格式 (YYYY-MM-DD)",
    "7. 重複單號精準剔除，資料由 30 筆還原為 28 筆唯一真實案"
]
for item in items_clean_samples:
    p = tf_a.add_paragraph()
    p.text = item
    p.font.name = FONT_MAIN
    p.font.size = Pt(10)
    p.font.color.rgb = C_SLATE_200
    p.space_before = Pt(4)

# Bottom Audit Box
box_sum = add_card(s7, Inches(0.8), Inches(5.95), Inches(11.7), Inches(1.1), bg_color=C_CARD_BG, border_color=C_EMERALD)
tf_s = box_sum.text_frame
tf_s.word_wrap = True
p_st = tf_s.paragraphs[0]
p_st.text = "🏆 示範檔最終官方檢核指標（加總 100% 吻合）"
p_st.font.name = FONT_MAIN
p_st.font.size = Pt(12)
p_st.font.bold = True
p_st.font.color.rgb = C_WHITE

p_sd = tf_s.add_paragraph()
p_sd.text = "• 唯一提單筆數：28 筆 (移除2筆重複)   |   • 總件數：31,340 件   |   • 總毛重：467,750.50 KG   |   • 海運費總額：$66,900.00 USD"
p_sd.font.name = FONT_ENG
p_sd.font.size = Pt(12)
p_sd.font.bold = True
p_sd.font.color.rgb = C_EMERALD
p_sd.space_before = Pt(3)

# ==============================================================================
# SLIDE 8: 核心功能黑科技 (Ctrl+E 快速填入 & Ctrl+H 尋找取代)
# ==============================================================================
s8 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s8, C_BG_DARK)
add_slide_header(s8, "Productivity Superpowers", "Excel 兩大神技：快速填入 (Ctrl+E) 與尋找與取代 (Ctrl+H)", 8)

# Left: Flash Fill
card_ff = add_card(s8, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.2))
tf_ff = card_ff.text_frame
tf_ff.word_wrap = True
p_fft = tf_ff.paragraphs[0]
p_fft.text = "⚡ 快速填入 (Flash Fill)  【Ctrl + E】"
p_fft.font.name = FONT_MAIN
p_fft.font.size = Pt(16)
p_fft.font.bold = True
p_fft.font.color.rgb = C_EMERALD

p_ffd = tf_ff.add_paragraph()
p_ffd.text = "微軟研究院 AI 機器學習黑科技！無需寫複雜公式，只需在相鄰直欄手動輸入第 1 格範本，按下快捷鍵瞬間自動推導全欄規律。"
p_ffd.font.name = FONT_MAIN
p_ffd.font.size = Pt(11)
p_ffd.font.color.rgb = C_SLATE_200
p_ffd.space_before = Pt(6)

ff_demo = [
    ("原始欄：1,200 CTNS", "手動輸入：1200 (示範範本)"),
    ("原始欄：450 PLTS", "按 Ctrl+E ➔ 450 (秒填)"),
    ("原始欄：850 PKGS", "按 Ctrl+E ➔ 850 (秒填)"),
    ("原始欄：2,100 CTNS", "按 Ctrl+E ➔ 2100 (秒填)")
]
for o, r in ff_demo:
    p = tf_ff.add_paragraph()
    p.text = f"• {o}  ➔  {r}"
    p.font.name = FONT_ENG
    p.font.size = Pt(11)
    p.font.color.rgb = C_CYAN
    p.space_before = Pt(4)

p_ffw = tf_ff.add_paragraph()
p_ffw.text = "\n⚠️ 注意事項：左側必須有連續資料行；填入產出為純文字靜態值，若原欄位資料變動需重新執行。"
p_ffw.font.name = FONT_MAIN
p_ffw.font.size = Pt(10)
p_ffw.font.color.rgb = C_AMBER

# Right: Find & Replace
card_fr = add_card(s8, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.2))
tf_fr = card_fr.text_frame
tf_fr.word_wrap = True
p_frt = tf_fr.paragraphs[0]
p_frt.text = "🔍 尋找與取代 (Find & Replace)  【Ctrl + H】"
p_frt.font.name = FONT_MAIN
p_frt.font.size = Pt(16)
p_frt.font.bold = True
p_frt.font.color.rgb = C_CYAN

p_frd = tf_fr.add_paragraph()
p_frd.text = "批量清洗字串雜質之外科手術刀！將「USD」、「$」、「KGS」等字串一次徹底移除，使文字數字瞬間還原為可加總的純數值。"
p_frd.font.name = FONT_MAIN
p_frd.font.size = Pt(11)
p_frd.font.color.rgb = C_SLATE_200
p_frd.space_before = Pt(6)

fr_tricks = [
    ("秘技 1：移除運費文字符號", "尋找「USD」取代為「空」；尋找「$」取代為「空」"),
    ("秘技 2：清除惡性強制換行符", "尋找目標按住【Ctrl + J】取代為「空」，瞬間抹去亂行！"),
    ("秘技 3：萬用字元清除括號備註", "尋找目標輸入「(*)」取代為「空」，清除櫃型備註文字")
]
for t, d in fr_tricks:
    p = tf_fr.add_paragraph()
    p.text = f"• {t}："
    p.font.name = FONT_MAIN
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    p.space_before = Pt(4)
    
    p2 = tf_fr.add_paragraph()
    p2.text = f"  {d}"
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(10)
    p2.font.color.rgb = C_SLATE_200

p_frtip = tf_fr.add_paragraph()
p_frtip.text = "\n💡 提示：取代完成後，選取整欄設定為「貨幣」或「數值」格式即恢復加總！"
p_frtip.font.name = FONT_MAIN
p_frtip.font.size = Pt(10)
p_frtip.font.color.rgb = C_EMERALD

# ==============================================================================
# SLIDE 9: 航運行政必備 8 大清理函數速查庫
# ==============================================================================
s9 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s9, C_BG_DARK)
add_slide_header(s9, "Formula Toolkit", "航運行政必備 8 大前處理函數速查庫", 9)

formulas_data = [
    ("=TRIM(text)", "文字清洗", "去除前後所有空格，單字間多餘空格僅留一個", "=TRIM(\"  TWKHH  \") → \"TWKHH\"", C_EMERALD),
    ("=CLEAN(text)", "控制字元", "移除自系統匯出時夾帶之不可列印字元與奇怪跳行", "=CLEAN(A2) 解決無法對齊亂碼", C_CYAN),
    ("=UPPER(text)", "大小寫", "將英文全數強制轉為大寫標準規格", "=UPPER(\"whlu\") → \"WHLU\"", C_BLUE),
    ("=VALUE(text)", "型態轉換", "將文字型態的數字轉換為可計算之真數值", "=VALUE(\"1500\") → 1500 (可加總)", C_AMBER),
    ("=SUBSTITUTE(t,o,n)", "字串替換", "將字串中的特定字元(如KGS、USD)替換為空", "=SUBSTITUTE(G2,\"KGS\",\"\")", C_PURPLE),
    ("=TEXT(val,fmt)", "格式鎖定", "將數值或日期鎖定為指定顯示字串格式", "=TEXT(TODAY(),\"YYYY-MM-DD\")", C_CYAN),
    ("=IFERROR(val,alt)", "容錯保護", "當計算或查詢出錯(#N/A)時，給予替代友好值", "=IFERROR(VLOOKUP(...),\"代碼異常\")", C_ROSE),
    ("=XLOOKUP(k,s,r)", "關聯查詢", "依港口或櫃型代碼跨表抓取標準名稱，無左側限制", "=XLOOKUP(D2,代碼!A:A,代碼!B:B)", C_EMERALD)
]

for idx, (fn, cat, desc, eg, col) in enumerate(formulas_data):
    row = idx // 4
    col_idx = idx % 4
    cx = Inches(0.8 + col_idx * 2.95)
    cy = Inches(1.6 + row * 2.65)
    
    card = add_card(s9, cx, cy, Inches(2.85), Inches(2.45))
    tf = card.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = fn
    p.font.name = FONT_ENG
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = col
    
    p_c = tf.add_paragraph()
    p_c.text = f"【{cat}】"
    p_c.font.name = FONT_MAIN
    p_c.font.size = Pt(10)
    p_c.font.color.rgb = C_SLATE_400
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = FONT_MAIN
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = C_SLATE_200
    p_d.space_before = Pt(4)
    
    p_e = tf.add_paragraph()
    p_e.text = f"範例：{eg}"
    p_e.font.name = FONT_ENG
    p_e.font.size = Pt(9)
    p_e.font.bold = True
    p_e.font.color.rgb = col
    p_e.space_before = Pt(4)

# ==============================================================================
# SLIDE 10: 300 筆多情境實戰練習題挑戰地圖
# ==============================================================================
s10 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s10, C_BG_DARK)
add_slide_header(s10, "300 Practice Scenarios", "300 筆航運實戰練習檔：4 大情境與官方基準指標", 10)

scen_data = [
    ("情境 A：美加遠洋出口 (80筆)", "長榮、陽明主力航線。考驗：美元千分位、提單重號、洛杉磯 (USLAX/USLGB) 空格。", C_BLUE),
    ("情境 B：歐地航線進口 (75筆)", "馬士基、赫伯羅德。考驗：鹿特丹/漢堡日期反轉 (DD/MM)、運費含 USD 符號文字。", C_EMERALD),
    ("情境 C：東南亞散貨併櫃 (75筆)", "萬海近洋線。考驗：單位多樣 (BAGS/DRUMS/PLTS)、毛重含 KGS 尾碼、條款空白。", C_AMBER),
    ("情境 D：高時效特殊冷鏈 (70筆)", "生鮮冷藏特種櫃 (20RF/40HR)。考驗：純數字日期 (20260315)、通關狀態大小寫混亂。", C_PURPLE)
]

for idx, (title, desc, col) in enumerate(scen_data):
    cx = Inches(0.8 + idx * 2.95)
    card = add_card(s10, cx, Inches(1.6), Inches(2.85), Inches(2.3))
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = FONT_MAIN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = col
    
    p2 = tf.add_paragraph()
    p2.text = desc
    p2.font.name = FONT_MAIN
    p2.font.size = Pt(10)
    p2.font.color.rgb = C_SLATE_200
    p2.space_before = Pt(6)

# Bottom Box: Official Benchmarks
bench_box = add_card(s10, Inches(0.8), Inches(4.2), Inches(11.733), Inches(2.6))
tf_bb = bench_box.text_frame
tf_bb.word_wrap = True
p_bbt = tf_bb.paragraphs[0]
p_bbt.text = "🏆 官方標準考核檢核指標（完成 300 筆清洗後對照核對，數值 100% 吻合即通關！）"
p_bbt.font.name = FONT_MAIN
p_bbt.font.size = Pt(14)
p_bbt.font.bold = True
p_bbt.font.color.rgb = C_WHITE

bench_badges = [
    ("唯一提單筆數", "285 筆", "剔除 15 筆重複"),
    ("總件數加總 (Packages)", "406,857 件", "=SUM(H5:H289)"),
    ("總毛重加總 (Gross Wt)", "6,120,737.68 KG", "=SUM(J5:J289)"),
    ("海運費總額 (Freight USD)", "$ 859,565.66", "=SUM(K5:K289)")
]
for idx, (lbl, val, note) in enumerate(bench_badges):
    bx = Inches(1.1 + idx * 2.8)
    b_card = add_card(s10, bx, Inches(4.8), Inches(2.6), Inches(1.7), bg_color=RGBColor(15, 23, 42), border_color=C_EMERALD)
    tf_bc = b_card.text_frame
    tf_bc.word_wrap = True
    p1 = tf_bc.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    p1.text = lbl
    p1.font.name = FONT_MAIN
    p1.font.size = Pt(10)
    p1.font.color.rgb = C_SLATE_400
    
    p2 = tf_bc.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.text = val
    p2.font.name = FONT_ENG
    p2.font.size = Pt(14)
    p2.font.bold = True
    p2.font.color.rgb = C_EMERALD
    p2.space_before = Pt(4)
    
    p3 = tf_bc.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.text = note
    p3.font.name = FONT_ENG
    p3.font.size = Pt(9)
    p3.font.color.rgb = C_SLATE_400

# ==============================================================================
# SLIDE 11: 航運報表交付前 10 點品質自我檢核清單
# ==============================================================================
s11 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s11, C_BG_DARK)
add_slide_header(s11, "Quality Assurance Checklist", "航運報表交付前 10 點品質自我檢核清單", 11)

chk_items_left = [
    ("1. 無合併儲存格", "資料網格區絕對無跨欄置中，每一單元格各自獨立。"),
    ("2. 第 1 列為唯一標題", "無雙層表頭，欄位名稱精簡明確、無重覆欄名。"),
    ("3. 提單號無重複值", "以條件格式化確認無紅色警示，COUNTIF 結果均為 1。"),
    ("4. 無隱形前後空格", "以 =LEN(A2)=LEN(TRIM(A2)) 驗證確認無多餘空格。"),
    ("5. 數值欄位可加總", "框選運費與毛重欄位，狀態列能正常顯示「加總」與「平均」。")
]

chk_items_right = [
    ("6. 日期格式整齊統一", "篩選器能正確以「年/月/日」樹狀分層展開，無純文字。"),
    ("7. 英文代碼全數大寫", "港口與船公司代碼皆為標準英文字母大寫 (TWKHH)。"),
    ("8. 公式已貼上為值", "靜態分析資料已全選「複製」並「貼上為值」消除動態依賴。"),
    ("9. 缺漏值有明確標註", "未放行欄位明確標示 'PENDING' 或 'TBA'，絕不留白。"),
    ("10. 導入下拉清單防錯", "設定資料驗證 (Data Validation)，從源頭杜絕手民誤植。")
]

for col_idx, items in enumerate([chk_items_left, chk_items_right]):
    cx = Inches(0.8 + col_idx * 5.95)
    card = add_card(s11, cx, Inches(1.6), Inches(5.78), Inches(5.2))
    tf = card.text_frame
    tf.word_wrap = True
    
    p0 = tf.paragraphs[0]
    p0.text = f"檢核項目 0{col_idx*5+1} 至 0{col_idx*5+5}"
    p0.font.name = FONT_ENG
    p0.font.size = Pt(11)
    p0.font.bold = True
    p0.font.color.rgb = C_EMERALD
    
    for t, d in items:
        p = tf.add_paragraph()
        p.text = f"☑ {t}"
        p.font.name = FONT_MAIN
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.space_before = Pt(8)
        
        p2 = tf.add_paragraph()
        p2.text = f"   {d}"
        p2.font.name = FONT_MAIN
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = C_SLATE_200

# ==============================================================================
# SLIDE 12: 結語、教材資源與參考國際規範
# ==============================================================================
s12 = prs.slides.add_slide(blank_slide_layout)
set_slide_background(s12, C_BG_DARK)
add_slide_header(s12, "Summary & Resources", "恭喜完成訓練！航運前處理教材配套與國際規範", 12)

# 3 Resources Cards
res_data = [
    ("📄 完整 Word 實務手冊 (.docx)", "航運行政人員_Excel資料前處理實務教學手冊.docx", "含 12 章完整解說、10 幅操作截圖、12 組表格、排版目錄與頁碼", C_EMERALD),
    ("📊 30 筆示範檔案 (.xlsx)", "航運事務資料_30筆示範檔.xlsx", "含「原始未清洗資料」與「標準化清洗完成表」，隨手比對演練", C_CYAN),
    ("🎯 300 筆實戰練習檔 (.xlsx)", "航運案例情境_300筆實戰練習檔.xlsx", "含 4 大航線情境、15 筆重複陷阱、參考標準解答與加總對帳表", C_PURPLE)
]

for idx, (title, filename, desc, col) in enumerate(res_data):
    cx = Inches(0.8 + idx * 3.95)
    card = add_card(s12, cx, Inches(1.6), Inches(3.85), Inches(2.5))
    tf = card.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = FONT_MAIN
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = col
    
    p_fn = tf.add_paragraph()
    p_fn.text = filename
    p_fn.font.name = FONT_ENG
    p_fn.font.size = Pt(10)
    p_fn.font.bold = True
    p_fn.font.color.rgb = C_WHITE
    p_fn.space_before = Pt(4)
    
    p_d = tf.add_paragraph()
    p_d.text = desc
    p_d.font.name = FONT_MAIN
    p_d.font.size = Pt(10)
    p_d.font.color.rgb = C_SLATE_200
    p_d.space_before = Pt(4)

# Bottom Box: Standards Reference
box_ref = add_card(s12, Inches(0.8), Inches(4.4), Inches(11.733), Inches(2.4))
tf_ref = box_ref.text_frame
tf_ref.word_wrap = True
p_reft = tf_ref.paragraphs[0]
p_reft.text = "📚 本教材編撰嚴格引用之國際航運法規與學術文獻"
p_reft.font.name = FONT_MAIN
p_reft.font.size = Pt(13)
p_reft.font.bold = True
p_reft.font.color.rgb = C_WHITE

refs_list = [
    "1. 聯合國貿易與運輸地點代碼標準 (UN/LOCODE) - UNECE Trade Facilitation Recommendations",
    "2. 國際貨運代理協會聯合會作業指引 (FIATA Standards for Digital B/L and Multimodal Transport)",
    "3. 國際標準化組織日期與時間規範 (ISO 8601-1:2019 Representation of dates and times)",
    "4. 結構化資料核心文獻 (Wickham, H. 2014. Tidy Data. Journal of Statistical Software)",
    "5. 微軟官方 Excel 函數與功能技術手冊 (Microsoft Excel Function & Data Cleaning Reference)"
]
for r in refs_list:
    p = tf_ref.add_paragraph()
    p.text = f"• {r}"
    p.font.name = FONT_ENG
    p.font.size = Pt(10)
    p.font.color.rgb = C_SLATE_200
    p.space_before = Pt(2)

prs.save(output_path)
print(f"Presentation saved successfully at: {output_path}")
