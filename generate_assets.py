# -*- coding: utf-8 -*-
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.font_manager import FontProperties
import numpy as np

font_path = "C:/Windows/Fonts/msjh.ttc"
font_bold_path = "C:/Windows/Fonts/msjhbd.ttc"
font_prop = FontProperties(fname=font_path)
font_bold_prop = FontProperties(fname=font_bold_path)

output_dir = "C:/proj/week12-1/assets"
os.makedirs(output_dir, exist_ok=True)

# -------------------------------------------------------------
# Figure 1: Excel 操作介面全覽圖 (Figure 1: Excel Interface Overview)
# -------------------------------------------------------------
def make_fig1():
    fig, ax = plt.subplots(figsize=(12, 7.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Background window
    bg = patches.Rectangle((2, 2), 96, 96, linewidth=1.5, edgecolor='#107C41', facecolor='#F3F2F1')
    ax.add_patch(bg)

    # Title bar (Green)
    title_bar = patches.Rectangle((2, 92), 96, 6, linewidth=0, facecolor='#107C41')
    ax.add_patch(title_bar)
    ax.text(5, 95, "AutoSave [ON]    航運管理系統_報關明細表.xlsx - Excel", color='white', fontproperties=font_bold_prop, fontsize=12, va='center')
    ax.text(93, 95, "—  □  ✕", color='white', fontproperties=font_bold_prop, fontsize=12, va='center')

    # Ribbon Tabs
    ribbon_tab_bg = patches.Rectangle((2, 85), 96, 7, linewidth=0, facecolor='#FFFFFF')
    ax.add_patch(ribbon_tab_bg)
    tabs = ["檔案", "常用 (Home)", "插入", "繪圖", "版面配置", "公式 (Formulas)", "資料 (Data)", "校閱", "檢視", "說明"]
    x_tab = 4
    for i, tab in enumerate(tabs):
        is_active = (i == 6) # 資料 (Data)
        if is_active:
            t_rect = patches.Rectangle((x_tab-1, 85), 13, 7, facecolor='#E1DFDD')
            ax.add_patch(t_rect)
            ax.text(x_tab+5.5, 88.5, tab, color='#107C41', fontproperties=font_bold_prop, fontsize=10, ha='center', va='center')
            x_tab += 14
        else:
            ax.text(x_tab+3, 88.5, tab, color='#323130', fontproperties=font_prop, fontsize=10, ha='center', va='center')
            x_tab += 9

    # Ribbon Body (Data tab)
    ribbon_body = patches.Rectangle((2, 67), 96, 18, linewidth=0.5, edgecolor='#E1DFDD', facecolor='#F8F9FA')
    ax.add_patch(ribbon_body)
    
    # Ribbon Groups inside Data tab
    groups = [
        ("取得與轉換資料", 4, 22, ["從文字/CSV", "從Web", "取得資料"]),
        ("排序與篩選", 28, 20, ["A-Z 遞增", "Z-A 遞減", "篩選 (Filter)", "清除"]),
        ("資料工具 (核心重點)", 50, 28, ["資料剖析", "快速填入", "移除重複項", "資料驗證"]),
        ("預測與分析", 80, 16, ["模擬分析", "預測工作表"])
    ]
    for g_title, gx, gw, items in groups:
        is_core = "核心重點" in g_title
        gborder = '#107C41' if is_core else '#D2D0CE'
        gface = '#EBF6EF' if is_core else '#FFFFFF'
        g_box = patches.Rectangle((gx, 70), gw, 14, linewidth=1.5 if is_core else 0.8, edgecolor=gborder, facecolor=gface, linestyle='--' if is_core else '-')
        ax.add_patch(g_box)
        ax.text(gx + gw/2, 72, g_title, color='#107C41' if is_core else '#605E5C', fontproperties=font_bold_prop if is_core else font_prop, fontsize=9, ha='center', va='center')
        for idx, item in enumerate(items):
            iy = 81 - (idx % 2) * 4
            ix = gx + 1 + (idx // 2) * (gw/2)
            i_btn = patches.Rectangle((ix, iy-1.5), gw/2 - 2, 3.5, facecolor='#EDEBE9', edgecolor='#C8C6C4', linewidth=0.5)
            ax.add_patch(i_btn)
            ax.text(ix + (gw/4 - 1), iy, item, color='#201F1E', fontproperties=font_prop, fontsize=8, ha='center', va='center')

    # Formula Bar and Name Box
    ax.add_patch(patches.Rectangle((2, 59), 96, 7, facecolor='#FFFFFF', edgecolor='#EDEBE9'))
    # Name Box
    ax.add_patch(patches.Rectangle((4, 60.5), 14, 4.5, facecolor='#FFFFFF', edgecolor='#8A8886', linewidth=1))
    ax.text(11, 62.75, "B4", color='#201F1E', fontproperties=font_bold_prop, fontsize=11, ha='center', va='center')
    ax.text(20, 62.75, "✕  ✓  fx", color='#605E5C', fontproperties=font_prop, fontsize=10, ha='center', va='center')
    # Formula input
    ax.add_patch(patches.Rectangle((24, 60.5), 72, 4.5, facecolor='#FFFFFF', edgecolor='#8A8886', linewidth=1))
    ax.text(26, 62.75, "=TRIM(CLEAN(C4))", color='#004578', fontproperties=font_bold_prop, fontsize=11, ha='left', va='center')

    # Grid Header
    ax.add_patch(patches.Rectangle((2, 53), 96, 5, facecolor='#E1DFDD', edgecolor='#C8C6C4'))
    cols = ["", "A", "B (提單號碼)", "C (船名航次)", "D (裝貨港)", "E (卸貨港)", "F (件數)", "G (運費USD)"]
    widths = [4, 12, 16, 20, 10, 10, 10, 14]
    cur_x = 2
    for c_idx, (c_name, w) in enumerate(zip(cols, widths)):
        ax.add_patch(patches.Rectangle((cur_x, 53), w, 5, facecolor='#F3F2F1', edgecolor='#C8C6C4', linewidth=0.5))
        ax.text(cur_x + w/2, 55.5, c_name, color='#323130', fontproperties=font_bold_prop, fontsize=9, ha='center', va='center')
        cur_x += w

    # Grid Rows
    sample_data = [
        ("1", "BK-901", "  ONEY240981  ", "ONE TRIUMPH / 024W", "TWKHH", "USLAX", "1,200 CTN", "$ 2,450.00"),
        ("2", "BK-902", "EVER981240", "EVER GIVEN / 0812E", "twkhh", "DEHAM", "450 PLT", "1850"),
        ("3", "BK-903", "MAEU334190  ", "MAERSK MC-KINNEY", "SGSIN", "NLRTM", "800", "USD 3200"),
        ("4", "BK-904", "YMLU881720", "YM WELLHEAD / 221S", "TWKHH", "JPTYO", "620 CTNS", "1,420.50"),
    ]
    for r_idx, row in enumerate(sample_data):
        ry = 47 - r_idx * 6
        cur_x = 2
        for c_idx, (val, w) in enumerate(zip(row, widths)):
            is_active_cell = (r_idx == 3 and c_idx == 1) # B4 cell
            bg_col = '#C7E0F4' if is_active_cell else ('#FFFFFF' if r_idx % 2 == 0 else '#FAFAFA')
            ax.add_patch(patches.Rectangle((cur_x, ry), w, 6, facecolor=bg_col, edgecolor='#107C41' if is_active_cell else '#E1DFDD', linewidth=2 if is_active_cell else 0.5))
            align = 'center' if c_idx in [0, 1, 4, 5] else ('right' if c_idx in [6, 7] else 'left')
            pos_x = cur_x + w/2 if align == 'center' else (cur_x + w - 1 if align == 'right' else cur_x + 1)
            ax.text(pos_x, ry + 3, val, color='#107C41' if is_active_cell else '#201F1E', fontproperties=font_bold_prop if is_active_cell else font_prop, fontsize=8.5, ha=align, va='center')
            cur_x += w

    # Sheet Tabs & Status Bar
    ax.add_patch(patches.Rectangle((2, 16), 96, 6, facecolor='#F3F2F1', edgecolor='#EDEBE9'))
    sheets = ["01_航運原始資料", "02_清理後標準表", "代碼對照表"]
    sx = 5
    for s_idx, s_name in enumerate(sheets):
        s_act = (s_idx == 0)
        ax.add_patch(patches.Rectangle((sx, 16), 22, 6, facecolor='#FFFFFF' if s_act else '#E1DFDD', edgecolor='#C8C6C4'))
        if s_act:
            ax.add_patch(patches.Rectangle((sx, 16), 22, 1, facecolor='#107C41'))
        ax.text(sx + 11, 19, s_name, color='#107C41' if s_act else '#605E5C', fontproperties=font_bold_prop if s_act else font_prop, fontsize=8.5, ha='center', va='center')
        sx += 23
    ax.text(sx + 2, 19, "+", color='#605E5C', fontproperties=font_bold_prop, fontsize=12, va='center')

    # Status Bar
    ax.add_patch(patches.Rectangle((2, 8), 96, 7, facecolor='#107C41'))
    ax.text(5, 11.5, "就緒 (Ready)   |   項目個數: 30   |   平均值: 2,230.12   |   加總: 66,903.60", color='white', fontproperties=font_prop, fontsize=9, va='center')
    ax.text(85, 11.5, "標準模式  [— 100% +]", color='white', fontproperties=font_prop, fontsize=9, va='center')

    # Callout Annotations pointing to interface elements
    callouts = [
        (75, 96, 80, 85, "1. 快速存取工具列 / 視窗標題", '#D83B01'),
        (12, 85, 15, 92, "2. 功能區索引標籤 (Ribbon Tabs)", '#0078D4'),
        (65, 68, 65, 59, "3. 資料工具群組 (前處理核心專用)", '#D83B01'),
        (10, 65, 11, 72, "4. 名稱方塊 (顯示當前儲存格 B4)", '#0078D4'),
        (60, 61, 60, 52, "5. 公式編輯列 (輸入/檢視清洗公式)", '#D83B01'),
        (10, 42, 10, 32, "6. 欄名列與儲存格工作區", '#0078D4'),
        (25, 16, 25, 7, "7. 工作表索引標籤 (Sheet Tabs)", '#0078D4'),
        (65, 8, 65, -1, "8. 狀態列 (數值摘要與檢視模式)", '#D83B01')
    ]
    for px, py, tx, ty, text, col in [
        (8, 88, 8, 97, "① 功能區索引標籤 (Ribbon)", '#0078D4'),
        (64, 76, 64, 88, "② 資料工具區 (資料剖析、移除重複項、快速填入)", '#D83B01'),
        (11, 65, 11, 74, "③ 名稱方塊", '#0078D4'),
        (55, 65, 55, 74, "④ 公式編輯列 (檢視原始公式與參照)", '#D83B01'),
        (18, 47, 18, 38, "⑤ 活頁簿資料網格 (含選取中儲存格)", '#0078D4'),
        (25, 16, 25, 6, "⑥ 工作表標籤 (Sheet Tabs)", '#0078D4'),
        (75, 8, 75, 3, "⑦ 狀態列 (即時統計加總/平均)", '#0078D4')
    ]:
        ax.annotate(text, xy=(px, py), xytext=(tx, ty),
                    arrowprops=dict(facecolor=col, edgecolor=col, width=1.5, headwidth=6, shrink=0.05),
                    bbox=dict(boxstyle="round,pad=0.3", fc="#FFF4CE" if col=='#D83B01' else "#EBF3FC", ec=col, lw=1.2),
                    fontproperties=font_bold_prop, fontsize=8.5, ha='center', va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig1_excel_interface.png"), dpi=200, bbox_inches='tight')
    plt.close()

# -------------------------------------------------------------
# Figure 2: 航運資料前處理標準作業流程 (Preprocessing Pipeline)
# -------------------------------------------------------------
def make_fig2():
    fig, ax = plt.subplots(figsize=(12, 6.5), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    steps = [
        ("步驟一：資料結構診斷", "• 檢視第一列是否為標題欄\n• 檢查合併儲存格與多餘空行\n• 確認各欄為單一屬性", "#E1DFDD", "#323130"),
        ("步驟二：重複與空值處理", "• 依提單號(B/L)移除重複值\n• 識別缺失關鍵資訊(POD/櫃型)\n• 標記或填補合理預設值", "#D2E9AF", "#107C41"),
        ("步驟三：欄位剖析與拆分", "• 船名航次拆分 (Vessel & Voy)\n• 件數與包裝單位拆分 (Qty & Unit)\n• 工具：資料剖析 / 快速填入", "#C7E0F4", "#0078D4"),
        ("步驟四：文字與代碼標準化", "• 去除前後隱藏空白 (TRIM)\n• 英文字母統一轉大寫 (UPPER)\n• 港口代碼與UN/LOCODE對齊", "#FFE8A3", "#D83B01"),
        ("步驟五：數值與日期清洗", "• 運費金額去貨幣符號轉數字\n• 重量千分位去除轉為純數值\n• 日期文字標準化為 YYYY-MM-DD", "#D8C2E7", "#5C2D91"),
        ("步驟六：品質驗證與鎖定", "• 設定資料驗證 (下拉清單防錯)\n• 複製整表貼上為「值」去除公式\n• 建立樞紐分析或匯出報表", "#B4E1FA", "#004578")
    ]

    for idx, (title, desc, bg_col, border_col) in enumerate(steps):
        row = idx // 3
        col = idx % 3
        x = 5 + col * 32
        y = 55 - row * 45
        w = 28
        h = 36
        
        box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=1.5",
                                     linewidth=2, edgecolor=border_col, facecolor=bg_col)
        ax.add_patch(box)
        
        # Header banner inside card
        banner = patches.Rectangle((x+0.5, y+h-7.5), w-1, 7, linewidth=0, facecolor=border_col)
        ax.add_patch(banner)
        ax.text(x + w/2, y + h - 4, title, color='white', fontproperties=font_bold_prop, fontsize=10.5, ha='center', va='center')
        
        # Desc
        ax.text(x + 2, y + h - 12, desc, color='#201F1E', fontproperties=font_prop, fontsize=9, ha='left', va='top', linespacing=1.6)

        # Draw arrows between steps
        if idx < 2:
            ax.annotate("", xy=(x + w + 3.5, y + h/2), xytext=(x + w + 0.5, y + h/2),
                        arrowprops=dict(facecolor='#605E5C', edgecolor='#605E5C', width=2, headwidth=7))
        elif idx == 2:
            ax.annotate("", xy=(x + w/2, y - 8), xytext=(x + w/2, y - 1),
                        arrowprops=dict(facecolor='#605E5C', edgecolor='#605E5C', width=2, headwidth=7))
        elif idx > 3:
            ax.annotate("", xy=(x - 3.5, y + h/2), xytext=(x - 0.5, y + h/2),
                        arrowprops=dict(facecolor='#605E5C', edgecolor='#605E5C', width=2, headwidth=7))

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig2_preprocessing_pipeline.png"), dpi=200, bbox_inches='tight')
    plt.close()

# -------------------------------------------------------------
# Figure 3: 航運行政常見 6 大髒資料類型 (Common Dirty Data Types)
# -------------------------------------------------------------
def make_fig3():
    fig, ax = plt.subplots(figsize=(12, 6.8), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    types = [
        ("1. 隱藏空白與非列印字元", "「  MAEU10928  」或字尾空格", "TRIM(A2) / CLEAN(A2)", "提單號查無資料、VLOOKUP失靈"),
        ("2. 欄位混雜未單一屬性", "「EVER GIVEN / 012E」船名+航次", "資料剖析 / 快速填入 / TEXTSPLIT", "無法單獨依航次篩選統計"),
        ("3. 數值儲存為文字或含單位", "「1,500 CTNS」或「$ 2,400.00」", "SUBSTITUTE() / VALUE() / 快速填入", "無法加總(SUM=0)、平均值計算錯誤"),
        ("4. 日期格式混亂多樣", "2026/3/5、26-03-05、20260305", "資料剖析(YMD) / DATE() / TEXT()", "無法依月份/季排序，時間軸錯亂"),
        ("5. 英文大小寫與代碼不一", "twkhh vs TWKHH vs Kao", "UPPER() / VLOOKUP代碼標準表", "統計港口貨量時被分散為不同項目"),
        ("6. 鍵入重複提單與幽靈資料", "同一提單號整列重複輸入兩次", "條件式格式化標註 / 移除重複項", "重複計算貨量、運費雙倍高估")
    ]

    ax.text(50, 96, "航運行政常見 6 大「髒資料」陷阱與對應處方", fontproperties=font_bold_prop, fontsize=14, ha='center', color='#107C41')

    # Draw table
    headers = ["髒資料類型", "航運實務常見案例範例", "Excel 建議處方 (功能/函數)", "對後續分析的負面影響"]
    xs = [5, 30, 60, 80]
    widths = [24, 29, 19, 17]

    # Header row
    ax.add_patch(patches.Rectangle((4, 87), 92, 6, facecolor='#107C41'))
    for h_txt, x, w in zip(headers, xs, widths):
        ax.text(x + w/2, 90, h_txt, color='white', fontproperties=font_bold_prop, fontsize=10, ha='center', va='center')

    for r_idx, (t_name, t_ex, t_fix, t_imp) in enumerate(types):
        y = 75 - r_idx * 13
        bg_col = '#F9F9F9' if r_idx % 2 == 0 else '#FFFFFF'
        ax.add_patch(patches.Rectangle((4, y), 92, 11, facecolor=bg_col, edgecolor='#E1DFDD', linewidth=1))

        # Col 1
        ax.text(xs[0] + 1, y + 5.5, t_name, color='#004578', fontproperties=font_bold_prop, fontsize=9.5, va='center')
        # Col 2
        ax.text(xs[1] + 1, y + 5.5, t_ex, color='#A80000', fontproperties=font_bold_prop, fontsize=9, va='center')
        # Col 3
        ax.text(xs[2] + 1, y + 5.5, t_fix, color='#107C41', fontproperties=font_bold_prop, fontsize=8.5, va='center')
        # Col 4
        ax.text(xs[3] + 1, y + 5.5, t_imp, color='#323130', fontproperties=font_prop, fontsize=8, va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig3_dirty_data_types.png"), dpi=200, bbox_inches='tight')
    plt.close()

# -------------------------------------------------------------
# Figure 4: 資料剖析三步驟對話方塊示意圖 (Text to Columns UI)
# -------------------------------------------------------------
def make_fig4():
    fig, ax = plt.subplots(figsize=(12, 7.2), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Main Dialog Window
    ax.add_patch(patches.FancyBboxPatch((10, 8), 80, 84, boxstyle="round,pad=1", edgecolor='#0078D4', facecolor='#F3F2F1', linewidth=2))
    
    # Dialog Title
    ax.add_patch(patches.Rectangle((10, 86), 80, 6, facecolor='#0078D4'))
    ax.text(13, 89, "文字匯入精靈 - 步驟 3 之 1 (分隔符號剖析船名航次)", color='white', fontproperties=font_bold_prop, fontsize=11, va='center')

    # Step 1 Pane
    ax.add_patch(patches.Rectangle((14, 62), 72, 21, facecolor='white', edgecolor='#C8C6C4'))
    ax.text(16, 80, "請選擇最符合您資料的原始資料類型：", fontproperties=font_bold_prop, fontsize=9.5, color='#323130')
    
    # Radio options
    ax.text(18, 74, "◉ 分隔符號 (Delimited) - 例如逗號、定位點、斜線 ( / ) 鍵隔開每個欄位", fontproperties=font_bold_prop, fontsize=9, color='#107C41')
    ax.text(18, 67, "○ 固定寬度 (Fixed Width) - 每欄皆對齊至具有空格的直欄", fontproperties=font_prop, fontsize=9, color='#605E5C')

    # Step 2 Pane (Delimiters)
    ax.add_patch(patches.Rectangle((14, 38), 72, 21, facecolor='white', edgecolor='#C8C6C4'))
    ax.text(16, 56, "選取分隔符號 (Delimiters)：", fontproperties=font_bold_prop, fontsize=9.5, color='#323130')
    ax.text(18, 50, "☐ Tab鍵      ☑ 空白 (Space)      ☑ 其他 (Other): [  /  ]", fontproperties=font_bold_prop, fontsize=9.5, color='#D83B01')
    ax.text(18, 43, "☑ 將連續出現的分隔符號視為單一處理 (Treat consecutive delimiters as one)", fontproperties=font_prop, fontsize=8.5, color='#323130')

    # Step 3 Data Preview
    ax.add_patch(patches.Rectangle((14, 15), 72, 20, facecolor='#FAFAFA', edgecolor='#8A8886'))
    ax.text(16, 32, "資料預覽 (Data Preview)：將原本單一欄位「EVER GIVEN / 0123E」自動分割為兩欄", fontproperties=font_bold_prop, fontsize=8.5, color='#004578')
    
    # Preview Grid
    ax.add_patch(patches.Rectangle((16, 17), 33, 12, facecolor='white', edgecolor='#0078D4', linewidth=1.5))
    ax.text(18, 26, "【直欄 1：船名 Vessel】\nEVER GIVEN\nONE TRIUMPH\nYM WELLHEAD", fontproperties=font_bold_prop, fontsize=8, color='#107C41')

    ax.add_patch(patches.Rectangle((51, 17), 33, 12, facecolor='white', edgecolor='#0078D4', linewidth=1.5))
    ax.text(53, 26, "【直欄 2：航次 Voyage】\n0123E\n024W\n221S", fontproperties=font_bold_prop, fontsize=8, color='#107C41')

    # Bottom Buttons
    buttons = ["< 上一步", "下一步 >", "完成 (Finish)", "取消"]
    bx = 45
    for b in buttons:
        is_fin = "完成" in b
        ax.add_patch(patches.Rectangle((bx, 9.5), 10, 4, facecolor='#107C41' if is_fin else '#E1DFDD', edgecolor='#8A8886'))
        ax.text(bx + 5, 11.5, b, color='white' if is_fin else '#201F1E', fontproperties=font_bold_prop if is_fin else font_prop, fontsize=8, ha='center', va='center')
        bx += 11

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig4_text_to_columns.png"), dpi=200, bbox_inches='tight')
    plt.close()

# -------------------------------------------------------------
# Figure 5: 移除重複項與快速填入操作圖 (Remove Duplicates & Flash Fill)
# -------------------------------------------------------------
def make_fig5():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6.2), dpi=200)
    
    # Subplot 1: Remove Duplicates Dialog
    ax1.set_xlim(0, 100)
    ax1.set_ylim(0, 100)
    ax1.axis('off')

    ax1.add_patch(patches.FancyBboxPatch((5, 5), 90, 90, boxstyle="round,pad=1", edgecolor='#107C41', facecolor='#F3F2F1', linewidth=1.5))
    ax1.add_patch(patches.Rectangle((5, 87), 90, 8, facecolor='#107C41'))
    ax1.text(8, 91, "移除重複項 (Remove Duplicates)", color='white', fontproperties=font_bold_prop, fontsize=11, va='center')

    ax1.text(8, 80, "若要刪除重複的值，請選取包含重複值的欄：", fontproperties=font_prop, fontsize=9, color='#201F1E')
    ax1.text(8, 73, "☑ 我的資料有標題 (My data has headers)", fontproperties=font_bold_prop, fontsize=9.5, color='#004578')

    # Button Bar
    ax1.add_patch(patches.Rectangle((8, 64), 22, 5, facecolor='#EDEBE9', edgecolor='#8A8886'))
    ax1.text(19, 66.5, "全選 (Select All)", fontproperties=font_prop, fontsize=8, ha='center', va='center')
    ax1.add_patch(patches.Rectangle((33, 64), 28, 5, facecolor='#EDEBE9', edgecolor='#8A8886'))
    ax1.text(47, 66.5, "全部取消選取 (Unselect)", fontproperties=font_bold_prop, fontsize=8, ha='center', va='center', color='#D83B01')

    # Column Check List
    ax1.add_patch(patches.Rectangle((8, 16), 84, 45, facecolor='white', edgecolor='#C8C6C4'))
    cols_check = [
        ("☑ 提單號碼 (B/L No.) 【關鍵識別碼：唯一勾選】", True),
        ("☐ 船名航次 (Vessel/Voyage)", False),
        ("☐ 裝貨港 (POL)", False),
        ("☐ 卸貨港 (POD)", False),
        ("☐ 運費金額 (Freight)", False),
        ("☐ 申報日期 (Date)", False)
    ]
    for c_i, (c_label, c_checked) in enumerate(cols_check):
        cy = 54 - c_i * 7
        ax1.text(12, cy, c_label, fontproperties=font_bold_prop if c_checked else font_prop, fontsize=8.5, color='#107C41' if c_checked else '#605E5C')

    # Result notification box
    ax1.add_patch(patches.Rectangle((8, 7), 84, 7, facecolor='#DFF6DD', edgecolor='#107C41'))
    ax1.text(50, 10.5, "已找到並移除 2 個重複值；保留 28 個唯一值。", fontproperties=font_bold_prop, fontsize=9, color='#107C41', ha='center', va='center')

    # Subplot 2: Flash Fill Concept
    ax2.set_xlim(0, 100)
    ax2.set_ylim(0, 100)
    ax2.axis('off')

    ax2.add_patch(patches.FancyBboxPatch((5, 5), 90, 90, boxstyle="round,pad=1", edgecolor='#0078D4', facecolor='#FFFFFF', linewidth=1.5))
    ax2.add_patch(patches.Rectangle((5, 87), 90, 8, facecolor='#0078D4'))
    ax2.text(8, 91, "快速填入 (Flash Fill) 快捷鍵：Ctrl + E", color='white', fontproperties=font_bold_prop, fontsize=11, va='center')

    ax2.text(8, 80, "AI 智慧感應規律：只要輸入第一個示範，按下 Ctrl+E 自動補齊", fontproperties=font_prop, fontsize=8.5, color='#323130')

    # Flash fill demonstration grid
    grid_y = 70
    headers2 = ["原始資料 (件數與包裝)", "人工示範欄", "按 Ctrl+E 自動填入結果"]
    ax2.add_patch(patches.Rectangle((8, 68), 84, 6, facecolor='#0078D4'))
    ax2.text(20, 71, headers2[0], color='white', fontproperties=font_bold_prop, fontsize=8.5, ha='center', va='center')
    ax2.text(50, 71, headers2[1], color='white', fontproperties=font_bold_prop, fontsize=8.5, ha='center', va='center')
    ax2.text(76, 71, headers2[2], color='white', fontproperties=font_bold_prop, fontsize=8.5, ha='center', va='center')

    ff_rows = [
        ("1,200 CTNS", "1200", "(手動鍵入範本)", '#FFF4CE'),
        ("450 PLTS", "450", "← Excel 瞬間自動辨識填入", '#DFF6DD'),
        ("800 PKGS", "800", "← Excel 瞬間自動辨識填入", '#DFF6DD'),
        ("2,500 BAGS", "2500", "← Excel 瞬間自動辨識填入", '#DFF6DD'),
        ("160 DRUMS", "160", "← Excel 瞬間自動辨識填入", '#DFF6DD')
    ]
    for r_i, (c1, c2, c3, bgc) in enumerate(ff_rows):
        ry = 59 - r_i * 9
        ax2.add_patch(patches.Rectangle((8, ry), 84, 8, facecolor=bgc, edgecolor='#E1DFDD'))
        ax2.text(20, ry + 4, c1, fontproperties=font_prop, fontsize=9, ha='center', va='center')
        ax2.text(50, ry + 4, c2, fontproperties=font_bold_prop, fontsize=9, ha='center', va='center', color='#004578')
        ax2.text(76, ry + 4, c3, fontproperties=font_bold_prop, fontsize=8, ha='center', va='center', color='#107C41')

    ax2.text(50, 10, "★ 提示：航運編號、港口字首、金額幣別拆分均適用！", fontproperties=font_bold_prop, fontsize=9, color='#D83B01', ha='center')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig5_remove_duplicates.png"), dpi=200, bbox_inches='tight')
    plt.close()

# -------------------------------------------------------------
# Figure 6: 尋找與取代及進階批次清理 (Find & Replace)
# -------------------------------------------------------------
def make_fig6():
    fig, ax = plt.subplots(figsize=(11, 6), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.add_patch(patches.FancyBboxPatch((10, 10), 80, 80, boxstyle="round,pad=1.2", edgecolor='#107C41', facecolor='#F3F2F1', linewidth=2))
    ax.add_patch(patches.Rectangle((10, 82), 80, 8, facecolor='#107C41'))
    ax.text(13, 86, "尋找與取代 (Find and Replace) - 快捷鍵：Ctrl + H", color='white', fontproperties=font_bold_prop, fontsize=12, va='center')

    # Tabs inside Find & Replace
    ax.add_patch(patches.Rectangle((14, 74), 20, 6, facecolor='white', edgecolor='#C8C6C4'))
    ax.text(24, 77, "取代 (Replace)", fontproperties=font_bold_prop, fontsize=10, ha='center', va='center', color='#107C41')
    ax.add_patch(patches.Rectangle((35, 74), 20, 6, facecolor='#EDEBE9', edgecolor='#C8C6C4'))
    ax.text(45, 77, "尋找 (Find)", fontproperties=font_prop, fontsize=10, ha='center', va='center', color='#605E5C')

    # Input Fields
    ax.text(15, 65, "尋找目標 (Find what)：", fontproperties=font_bold_prop, fontsize=10, color='#323130')
    ax.add_patch(patches.Rectangle((38, 61), 48, 7, facecolor='white', edgecolor='#8A8886', linewidth=1.5))
    ax.text(40, 64.5, "USD$  (或字元逗號 、貨幣符號)", fontproperties=font_bold_prop, fontsize=9.5, color='#D83B01', va='center')

    ax.text(15, 50, "取代為 (Replace with)：", fontproperties=font_bold_prop, fontsize=10, color='#323130')
    ax.add_patch(patches.Rectangle((38, 46), 48, 7, facecolor='white', edgecolor='#8A8886', linewidth=1.5))
    ax.text(40, 49.5, "(留空白表示直接刪除該字串)", fontproperties=font_prop, fontsize=9.5, color='#605E5C', va='center')

    # Buttons
    ax.add_patch(patches.Rectangle((15, 28), 24, 7, facecolor='#107C41', edgecolor='#0E6B37'))
    ax.text(27, 31.5, "全部取代 (Replace All)", fontproperties=font_bold_prop, fontsize=9.5, color='white', ha='center', va='center')

    ax.add_patch(patches.Rectangle((42, 28), 18, 7, facecolor='#EDEBE9', edgecolor='#8A8886'))
    ax.text(51, 31.5, "取代 (Replace)", fontproperties=font_prop, fontsize=9, color='#201F1E', ha='center', va='center')

    ax.add_patch(patches.Rectangle((63, 28), 23, 7, facecolor='#EDEBE9', edgecolor='#8A8886'))
    ax.text(74.5, 31.5, "尋找全部 (Find All)", fontproperties=font_prop, fontsize=9, color='#201F1E', ha='center', va='center')

    # Explanation callout
    ax.add_patch(patches.Rectangle((14, 15), 72, 9, facecolor='#FFF4CE', edgecolor='#C8C6C4'))
    ax.text(16, 20.5, "【航運妙用】：清理運費欄位時，尋找「USD」取代為空；尋找「,」取代為空，\n數值立即從「純文字字串」恢復為可加總的「真數值」！", fontproperties=font_bold_prop, fontsize=8.5, color='#7A3600', va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig6_flash_fill_find_replace.png"), dpi=200, bbox_inches='tight')
    plt.close()

# -------------------------------------------------------------
# Figure 7: 常用前處理函數大補帖 (Formula Toolkit)
# -------------------------------------------------------------
def make_fig7():
    fig, ax = plt.subplots(figsize=(12, 7), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.text(50, 96, "航運行政人員必備 8 大資料前處理函數速查表", fontproperties=font_bold_prop, fontsize=14, ha='center', color='#004578')

    formulas = [
        ("=TRIM(text)", "文字清洗", "去除文字前後所有空格，單字間多餘空格僅留一個", "=TRIM(\"  TWKHH  \") → \"TWKHH\""),
        ("=CLEAN(text)", "控制字元", "移除從系統匯出時夾帶之換行符、不可列印字元", "=CLEAN(A2) 解決無法對齊與奇怪斷行"),
        ("=UPPER(text)", "大小寫", "將英文代碼全數轉為大寫標準規格", "=UPPER(\"cnshe\") → \"CNSHE\""),
        ("=VALUE(text)", "型態轉換", "將文字型態的數字轉換為可計算的真數值", "=VALUE(\"1500\") → 1500 (可SUM加總)"),
        ("=SUBSTITUTE(t,o,n)", "字串替換", "將字串中的特定字元(如KGS、USD)替換為空", "=SUBSTITUTE(G2,\"KGS\",\"\")"),
        ("=TEXT(val,fmt)", "格式轉換", "將數值或日期鎖定為指定顯示字串格式", "=TEXT(TODAY(),\"YYYY-MM-DD\")"),
        ("=IFERROR(val,alt)", "容錯保護", "當計算或查詢出錯(#N/A)時，給予替代友好值", "=IFERROR(VLOOKUP(...),\"代碼異常\")"),
        ("=XLOOKUP(key,s,r)", "關聯查詢", "依港口或櫃型代碼，跨表抓取中文名稱或標準代碼", "=XLOOKUP(D2,代碼!A:A,代碼!B:B)")
    ]

    ax.add_patch(patches.Rectangle((4, 87), 92, 6, facecolor='#004578'))
    headers = ["函數語法", "類別", "核心功能說明", "航運實務應用範例"]
    xs = [5, 27, 40, 70]
    ws = [21, 12, 29, 25]
    for h, x, w in zip(headers, xs, ws):
        ax.text(x + w/2, 90, h, color='white', fontproperties=font_bold_prop, fontsize=10, ha='center', va='center')

    for r_idx, (fn, cat, desc, eg) in enumerate(formulas):
        y = 77 - r_idx * 9.5
        bg_col = '#F2F6FA' if r_idx % 2 == 0 else '#FFFFFF'
        ax.add_patch(patches.Rectangle((4, y), 92, 8.5, facecolor=bg_col, edgecolor='#D2D0CE', linewidth=0.8))
        ax.text(xs[0] + 1, y + 4.2, fn, color='#004578', fontproperties=font_bold_prop, fontsize=9, va='center')
        ax.text(xs[1] + ws[1]/2, y + 4.2, cat, color='#107C41', fontproperties=font_bold_prop, fontsize=9, ha='center', va='center')
        ax.text(xs[2] + 1, y + 4.2, desc, color='#201F1E', fontproperties=font_prop, fontsize=8.5, va='center')
        ax.text(xs[3] + 1, y + 4.2, eg, color='#D83B01', fontproperties=font_bold_prop, fontsize=8, va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig7_formula_toolkit.png"), dpi=200, bbox_inches='tight')
    plt.close()

# -------------------------------------------------------------
# Figure 8: 30筆示範檔清洗前後對照圖 (30 Demo Before & After)
# -------------------------------------------------------------
def make_fig8():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6.5), dpi=200)

    # Before (Dirty)
    ax1.set_xlim(0, 100)
    ax1.set_ylim(0, 100)
    ax1.axis('off')
    ax1.add_patch(patches.Rectangle((2, 2), 96, 96, facecolor='#FFF8F8', edgecolor='#D83B01', linewidth=2))
    ax1.add_patch(patches.Rectangle((2, 90), 96, 8, facecolor='#D83B01'))
    ax1.text(50, 94, "【清洗前】原始髒亂航運資料 (無法直接統計分析)", color='white', fontproperties=font_bold_prop, fontsize=10.5, ha='center', va='center')

    dirty_samples = [
        "1. 提單號前後大量空格：'  EVER102938  '",
        "2. 船名與航次黏在一起：'EVER GIVEN / 012E'",
        "3. 港口代碼大小寫雜亂：'twkhh'、'USlax'",
        "4. 件數欄混雜英文單位：'1,200 CTNS' (文字型)",
        "5. 運費欄夾帶幣別符號：'$ 2,450.00' (文字型)",
        "6. 日期格式四種混用：2026/3/5、26-03-05",
        "7. 存在整列重複提單號：EVER102938 建檔兩次",
        "8. 異常空值未標記：付款條款空白、櫃型遺漏"
    ]
    for idx, s in enumerate(dirty_samples):
        sy = 78 - idx * 9.5
        ax1.add_patch(patches.Rectangle((5, sy), 90, 8, facecolor='#FFFFFF', edgecolor='#E1DFDD'))
        ax1.text(8, sy + 4, s, fontproperties=font_prop, fontsize=9, color='#A80000', va='center')

    # After (Clean)
    ax2.set_xlim(0, 100)
    ax2.set_ylim(0, 100)
    ax2.axis('off')
    ax2.add_patch(patches.Rectangle((2, 2), 96, 96, facecolor='#F6FFF8', edgecolor='#107C41', linewidth=2))
    ax2.add_patch(patches.Rectangle((2, 90), 96, 8, facecolor='#107C41'))
    ax2.text(50, 94, "【清洗後】標準結構化資料 (分析/樞紐即刻可用)", color='white', fontproperties=font_bold_prop, fontsize=10.5, ha='center', va='center')

    clean_samples = [
        "1. 提單號去除所有頭尾空格，乾淨唯一碼",
        "2. 船名(Vessel) 與 航次(Voyage) 獨立成兩欄",
        "3. 裝卸港全數標準大寫 UN/LOCODE (TWKHH)",
        "4. 件數拆為【純數值: 1200】與【單位: CTN】",
        "5. 運費轉為純數值 2450.00，設定貨幣顯示格式",
        "6. 日期全數轉為標準日期格式 (YYYY-MM-DD)",
        "7. 重複單號精準剔除，維持 28 筆真實航運案",
        "8. 缺失欄位補齊或標註 'PENDING' / 'TBA'"
    ]
    for idx, s in enumerate(clean_samples):
        sy = 78 - idx * 9.5
        ax2.add_patch(patches.Rectangle((5, sy), 90, 8, facecolor='#FFFFFF', edgecolor='#E1DFDD'))
        ax2.text(8, sy + 4, s, fontproperties=font_bold_prop, fontsize=9, color='#107C41', va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig8_30_demo_before_after.png"), dpi=200, bbox_inches='tight')
    plt.close()

# -------------------------------------------------------------
# Figure 9: 300筆實戰練習題庫架構 (300 Practice Scenarios)
# -------------------------------------------------------------
def make_fig9():
    fig, ax = plt.subplots(figsize=(12, 6.8), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    ax.text(50, 96, "300 筆多情境航運實戰練習檔——架構與挑戰地圖", fontproperties=font_bold_prop, fontsize=14, ha='center', color='#004578')

    scenarios = [
        ("情境 A：美加遠洋出口櫃 (80筆)", "• 挑戰：美元千分位、提單重號、美東美西港口代碼混淆\n• 涵蓋船司：長榮(EMC)、陽明(YML)、地中海(MSC)", "#EBF3FC", "#0078D4"),
        ("情境 B：歐地近遠洋進口櫃 (75筆)", "• 挑戰：歐元/美金匯率符號、鹿特丹/漢堡日期反轉(DD/MM)\n• 涵蓋船司：馬士基(MSK)、赫伯羅德(Hapag)", "#EBF6EF", "#107C41"),
        ("情境 C：東南亞海運散貨併櫃 (75筆)", "• 挑戰：件數單位混用(CTN/PLT/PKG)、毛重含文字KGS\n• 涵蓋船司：萬海(WHL)、德翔(TSL)", "#FFF4CE", "#D83B01"),
        ("情境 D：高時效航運特殊冷鏈 (70筆)", "• 挑戰：溫控要求註記混雜、付款條款空白、異常字元\n• 涵蓋特種櫃型：20RF、40HR、危險品(DG)註記", "#F4EDF9", "#5C2D91")
    ]

    for idx, (title, content, bg_col, border_col) in enumerate(scenarios):
        row = idx // 2
        col = idx % 2
        x = 5 + col * 46
        y = 52 - row * 38
        w = 44
        h = 32

        box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=1.2",
                                     linewidth=1.8, edgecolor=border_col, facecolor=bg_col)
        ax.add_patch(box)
        banner = patches.Rectangle((x+0.5, y+h-7), w-1, 6.5, facecolor=border_col)
        ax.add_patch(banner)
        ax.text(x + w/2, y + h - 3.7, title, color='white', fontproperties=font_bold_prop, fontsize=10, ha='center', va='center')
        ax.text(x + 2, y + h - 11, content, color='#201F1E', fontproperties=font_prop, fontsize=8.5, ha='left', va='top', linespacing=1.6)

    ax.add_patch(patches.Rectangle((5, 4), 90, 8, facecolor='#EDEBE9', edgecolor='#8A8886'))
    ax.text(50, 8, "★ 解題考核核心：學員完成清洗後，加總總運費、加總毛重、唯一提單數必須 100% 吻合參考解答！", fontproperties=font_bold_prop, fontsize=9.5, color='#D83B01', ha='center', va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig9_300_practice_scenarios.png"), dpi=200, bbox_inches='tight')
    plt.close()

# -------------------------------------------------------------
# Figure 10: 資料驗證與條件式格式化操作圖 (Data Validation & Formatting)
# -------------------------------------------------------------
def make_fig10():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6.2), dpi=200)

    # Subplot 1: Data Validation Dropdown
    ax1.set_xlim(0, 100)
    ax1.set_ylim(0, 100)
    ax1.axis('off')
    ax1.add_patch(patches.FancyBboxPatch((5, 5), 90, 90, boxstyle="round,pad=1", edgecolor='#107C41', facecolor='#F3F2F1', linewidth=1.5))
    ax1.add_patch(patches.Rectangle((5, 87), 90, 8, facecolor='#107C41'))
    ax1.text(8, 91, "資料驗證 (Data Validation) - 建立防呆清單", color='white', fontproperties=font_bold_prop, fontsize=11, va='center')

    ax1.text(8, 80, "設定儲存格僅允許輸入預先定義好的標準櫃型：", fontproperties=font_prop, fontsize=8.5, color='#201F1E')

    ax1.add_patch(patches.Rectangle((8, 30), 84, 46, facecolor='white', edgecolor='#C8C6C4'))
    ax1.text(12, 70, "儲存格內允許 (Allow)：", fontproperties=font_bold_prop, fontsize=9, color='#323130')
    ax1.add_patch(patches.Rectangle((42, 66), 46, 6, facecolor='#FAFAFA', edgecolor='#8A8886'))
    ax1.text(45, 69, "清單 (List)  ▼", fontproperties=font_bold_prop, fontsize=9, color='#004578', va='center')

    ax1.text(12, 54, "來源 (Source)：", fontproperties=font_bold_prop, fontsize=9, color='#323130')
    ax1.add_patch(patches.Rectangle((42, 50), 46, 6, facecolor='#FAFAFA', edgecolor='#8A8886'))
    ax1.text(45, 53, "20GP, 40GP, 40HC, 45HC, 20RF", fontproperties=font_bold_prop, fontsize=8, color='#D83B01', va='center')

    ax1.text(12, 38, "☑ 提供儲存格內下拉式清單 (In-cell dropdown)", fontproperties=font_prop, fontsize=8.5, color='#107C41')

    # Visual dropdown effect
    ax1.add_patch(patches.Rectangle((8, 10), 84, 16, facecolor='#EBF3FC', edgecolor='#0078D4'))
    ax1.text(12, 20, "實際操作效果：使用者點擊儲存格時，\n自動彈出合法選項，杜絕輸入「40HQ」或打錯字！", fontproperties=font_bold_prop, fontsize=8.5, color='#004578', va='center')

    # Subplot 2: Conditional Formatting Duplicates
    ax2.set_xlim(0, 100)
    ax2.set_ylim(0, 100)
    ax2.axis('off')
    ax2.add_patch(patches.FancyBboxPatch((5, 5), 90, 90, boxstyle="round,pad=1", edgecolor='#D83B01', facecolor='#F3F2F1', linewidth=1.5))
    ax2.add_patch(patches.Rectangle((5, 87), 90, 8, facecolor='#D83B01'))
    ax2.text(8, 91, "設定格式化的條件 - 高亮重複提單號", color='white', fontproperties=font_bold_prop, fontsize=11, va='center')

    ax2.text(8, 80, "路徑：常用 > 設定格式化的條件 > 醒目提示儲存格規則 > 重複的值", fontproperties=font_prop, fontsize=8.5, color='#201F1E')

    # Dialog mock
    ax2.add_patch(patches.Rectangle((8, 30), 84, 46, facecolor='white', edgecolor='#C8C6C4'))
    ax2.text(12, 68, "設定包含下列值的儲存格格式：", fontproperties=font_bold_prop, fontsize=9, color='#323130')
    ax2.add_patch(patches.Rectangle((12, 54), 36, 7, facecolor='#FAFAFA', edgecolor='#8A8886'))
    ax2.text(15, 57.5, "重複 (Duplicate)  ▼", fontproperties=font_bold_prop, fontsize=9, color='#004578', va='center')

    ax2.text(50, 57.5, "設定為：", fontproperties=font_prop, fontsize=9, color='#323130', va='center')
    ax2.add_patch(patches.Rectangle((62, 54), 28, 7, facecolor='#FFC7CE', edgecolor='#9C0006'))
    ax2.text(76, 57.5, "淺紅色填滿與\n深紅色文字", fontproperties=font_bold_prop, fontsize=7.5, color='#9C0006', ha='center', va='center')

    # Grid effect
    ax2.add_patch(patches.Rectangle((8, 10), 84, 16, facecolor='#FFF8F8', edgecolor='#D83B01'))
    ax2.text(12, 20, "視覺效果：重複建檔的提單立即可見(紅色警示)！\n方便行政人員在送審前即時攔截錯誤。", fontproperties=font_bold_prop, fontsize=8.5, color='#A80000', va='center')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "fig10_data_validation_rules.png"), dpi=200, bbox_inches='tight')
    plt.close()

if __name__ == "__main__":
    print("Generating Figure 1...")
    make_fig1()
    print("Generating Figure 2...")
    make_fig2()
    print("Generating Figure 3...")
    make_fig3()
    print("Generating Figure 4...")
    make_fig4()
    print("Generating Figure 5...")
    make_fig5()
    print("Generating Figure 6...")
    make_fig6()
    print("Generating Figure 7...")
    make_fig7()
    print("Generating Figure 8...")
    make_fig8()
    print("Generating Figure 9...")
    make_fig9()
    print("Generating Figure 10...")
    make_fig10()
    print("All 10 figures generated successfully in", output_dir)
