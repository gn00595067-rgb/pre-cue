import streamlit as st
import pandas as pd
import math
import io
import os
import shutil
import tempfile
import subprocess
import re
import requests
import base64
from datetime import timedelta, datetime, date
from copy import copy
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill, Color

# =========================================================
# 0. 初始化 Session State
# =========================================================
if "is_supervisor" not in st.session_state:
    st.session_state.is_supervisor = False

# =========================================================
# 1. 基礎工具
# =========================================================
def parse_count_to_int(x):
    if x is None: return 0
    if isinstance(x, (int, float)): return int(x)
    s = str(x)
    m = re.findall(r"[\d,]+", s)
    if not m: return 0
    return int(m[0].replace(",", ""))

def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

def html_escape(s):
    if s is None: return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;")

# =========================================================
# 2. 頁面設定
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v91.0")

# =========================================================
# 3. PDF 策略
# =========================================================
def find_soffice_path():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice: return soffice
    if os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in candidates:
            if os.path.exists(p): return p
    return None

def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    soffice = find_soffice_path()
    if not soffice: 
        return None, "Fail", "無可用的 LibreOffice 引擎"

    try:
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, "cue.xlsx")
            with open(xlsx_path, "wb") as f: f.write(xlsx_bytes)
            
            subprocess.run([
                soffice, "--headless", "--nologo", "--convert-to", "pdf:calc_pdf_Export", 
                "--outdir", tmp, xlsx_path
            ], capture_output=True, timeout=60)
            
            pdf_path = os.path.join(tmp, "cue.pdf")
            if not os.path.exists(pdf_path):
                for fn in os.listdir(tmp):
                    if fn.endswith(".pdf"): pdf_path = os.path.join(tmp, fn); break
            
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f: return f.read(), "LibreOffice", ""
            return None, "Fail", "LibreOffice 轉檔無輸出"
    except Exception as e: return None, "Fail", str(e)

def html_to_pdf_weasyprint(html_str):
    try:
        from weasyprint import HTML, CSS
        from weasyprint.text.fonts import FontConfiguration
        font_config = FontConfiguration()
        css = CSS(string="@page { size: A4 landscape; margin: 0.5cm; } body { font-family: sans-serif; }")
        pdf_bytes = HTML(string=html_str).write_pdf(stylesheets=[css], font_config=font_config)
        return pdf_bytes, ""
    except Exception as e: return None, str(e)

# =========================================================
# 4. 核心資料設定 (雲端 Google Sheet 版)
# =========================================================
GSHEET_SHARE_URL = "https://docs.google.com/spreadsheets/d/1bzmG-N8XFsj8m3LUPqA8K70AcIqaK4Qhq1VPWcK0w_s/edit?usp=sharing"

@st.cache_data(ttl=300)
def load_config_from_cloud(share_url):
    try:
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", share_url)
        if not match: return None, None, None, None, "連結格式錯誤"
        file_id = match.group(1)
        
        def read_sheet(sheet_name):
            url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
            return pd.read_csv(url)

        df_store = read_sheet("Stores")
        df_store.columns = [c.strip() for c in df_store.columns]
        store_counts = dict(zip(df_store['Key'], df_store['Display_Name']))
        store_counts_num = dict(zip(df_store['Key'], df_store['Count']))

        df_fact = read_sheet("Factors")
        df_fact.columns = [c.strip() for c in df_fact.columns]
        sec_factors = {}
        for _, row in df_fact.iterrows():
            if row['Media'] not in sec_factors: sec_factors[row['Media']] = {}
            sec_factors[row['Media']][int(row['Seconds'])] = float(row['Factor'])
        
        name_map = {"全家新鮮視": "新鮮視", "全家廣播": "全家廣播", "家樂福": "家樂福"}
        for k, v in name_map.items():
            if k in sec_factors and v not in sec_factors:
                sec_factors[v] = sec_factors[k]

        df_price = read_sheet("Pricing")
        df_price.columns = [c.strip() for c in df_price.columns]
        pricing_db = {}
        for _, row in df_price.iterrows():
            m = row['Media']
            r = row['Region']
            if m == "家樂福":
                if m not in pricing_db: pricing_db[m] = {}
                pricing_db[m][r] = {
                    "List": int(row['List_Price']),
                    "Net": int(row['Net_Price']),
                    "Std_Spots": int(row['Std_Spots']),
                    "Day_Part": row['Day_Part']
                }
            else:
                if m not in pricing_db:
                    pricing_db[m] = {"Std_Spots": int(row['Std_Spots']), "Day_Part": row['Day_Part']}
                pricing_db[m][r] = [int(row['List_Price']), int(row['Net_Price'])]
            
        return store_counts, store_counts_num, pricing_db, sec_factors, None

    except Exception as e:
        return None, None, None, None, f"讀取失敗: {str(e)}"

with st.spinner("正在連線 Google Sheet 載入最新價格表..."):
    STORE_COUNTS, STORE_COUNTS_NUM, PRICING_DB, SEC_FACTORS, err_msg = load_config_from_cloud(GSHEET_SHARE_URL)

if err_msg:
    st.error(f"❌ 設定檔載入失敗: {err_msg}")
    st.stop()

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

REGION_DISPLAY_MAP = {
    "北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東",
    "全省量販": "全省量販", "全省超市": "全省超市"
}
def region_display(region): return REGION_DISPLAY_MAP.get(region, region)

def get_sec_factor(media_type, seconds):
    factors = SEC_FACTORS.get(media_type)
    if not factors:
        if media_type == "新鮮視": factors = SEC_FACTORS.get("全家新鮮視")
        elif media_type == "全家廣播": factors = SEC_FACTORS.get("全家廣播")
    
    if not factors: return 1.0
    if seconds in factors: return factors[seconds]
    for base in [10, 20, 15, 30]:
        if base in factors: return (seconds / base) * factors[base]
    return 1.0

def calculate_schedule(total_spots, days):
    if days <= 0: return []
    if total_spots % 2 != 0: total_spots += 1
    half_spots = total_spots // 2
    base, rem = divmod(half_spots, days)
    sch = [base + (1 if i < rem else 0) for i in range(days)]
    return [x * 2 for x in sch]

def get_remarks_text(sign_deadline, billing_month, payment_date):
    d_str = sign_deadline.strftime("%Y/%m/%d (%a) %H:%M") if sign_deadline else "____/__/__ (__) 12:00"
    p_str = payment_date.strftime("%Y/%m/%d") if payment_date else "____/__/__"
    return [
        f"1.請於 {d_str}前 回簽及進單，方可順利上檔。",
        "2.以上節目名稱如有異動，以上檔時節目名稱為主，如遇時段滿檔，上檔時間挪後或更換至同級時段。",
        "3.通路店鋪數與開機率至少七成(以上)。每日因加盟數調整，或遇店舖年度季度改裝、設備維護升級及保修等狀況，會有一定幅度增減。",
        "4.託播方需於上檔前 5 個工作天，提供廣告帶(mp3)、影片/影像 1920x1080 (mp4)。",
        f"5.雙方同意費用請款月份 : {billing_month}，如有修正必要，將另行E-Mail告知，並視為正式合約之一部分。",
        f"6.付款兌現日期：{p_str}"
    ]

# =========================================================
# 5. 核心計算函式
# =========================================================
def calculate_plan_data(config, total_budget, days_count):
    rows = []
    total_list_accum = 0
    debug_logs = []

    for m, cfg in config.items():
        m_budget_total = total_budget * (cfg["share"] / 100.0)
        
        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget_total * (sec_pct / 100.0)
            if s_budget <= 0: continue
            
            factor = get_sec_factor(m, sec)
            
            if m in ["全家廣播", "新鮮視"]:
                db = PRICING_DB[m]
                calc_regs = ["全省"] if cfg["is_national"] else cfg["regions"]
                display_regs = REGIONS_ORDER if cfg["is_national"] else cfg["regions"]
                
                unit_net_sum = 0
                for r in calc_regs:
                    unit_net_sum += (db[r][1] / db["Std_Spots"]) * factor
                if unit_net_sum == 0: continue
                
                spots_init = math.ceil(s_budget / unit_net_sum)
                is_under_target = spots_init < db["Std_Spots"]
                calc_penalty = 1.1 if is_under_target else 1.0 
                
                if cfg["is_national"]:
                    row_display_penalty = 1.0 
                    total_display_penalty = 1.1 if is_under_target else 1.0
                    status_msg = "全省(分區豁免/總價懲罰)" if is_under_target else "達標"
                else:
                    row_display_penalty = 1.1 if is_under_target else 1.0
                    total_display_penalty = 1.0 
                    status_msg = "未達標 x1.1" if is_under_target else "達標"

                spots_final = math.ceil(s_budget / (unit_net_sum * calc_penalty))
                if spots_final % 2 != 0: spots_final += 1
                if spots_final == 0: spots_final = 2

                log_details = []
                sch = calculate_schedule(spots_final, days_count)
                nat_pkg_display = 0
                
                if cfg["is_national"]:
                    nat_list = db["全省"][0]
                    nat_unit_price = int((nat_list / db["Std_Spots"]) * factor * total_display_penalty)
                    nat_pkg_display = nat_unit_price * spots_final
                    total_list_accum += nat_pkg_display
                    log_details.append(f"**全省總價**: ${nat_pkg_display:,} (單價 ${nat_unit_price:,} x {spots_final})")

                for i, r in enumerate(display_regs):
                    list_price_region = db[r][0]
                    unit_rate_display = int((list_price_region / db["Std_Spots"]) * factor * row_display_penalty)
                    total_rate_display = unit_rate_display * spots_final 
                    row_pkg_display = total_rate_display
                    if not cfg["is_national"]:
                        total_list_accum += row_pkg_display
                        log_details.append(f"**{r}**: ${total_rate_display:,} (單價 ${unit_rate_display:,} x {spots_final})")

                    rows.append({
                        "media": m, "region": r,
                        "program_num": STORE_COUNTS_NUM.get(f"新鮮視_{r}" if m=="新鮮視" else r, 0),
                        "daypart": db["Day_Part"], "seconds": sec,
                        "spots": spots_final, "schedule": sch,
                        "rate_display": total_rate_display, 
                        "pkg_display": row_pkg_display,
                        "is_pkg_member": cfg["is_national"],
                        "nat_pkg_display": nat_pkg_display
                    })
                
                debug_logs.append({"Media": f"{m} ({sec}s)", "Budget": f"${s_budget:,.0f}", "Status": f"執行 {spots_final} 檔 ({status_msg})", "Details": log_details})

            elif m == "家樂福":
                db = PRICING_DB["家樂福"]
                base_std = db["量販_全省"]["Std_Spots"]
                unit_net = (db["量販_全省"]["Net"] / base_std) * factor
                spots_init = math.ceil(s_budget / unit_net)
                penalty = 1.1 if spots_init < base_std else 1.0
                spots_final = math.ceil(s_budget / (unit_net * penalty))
                if spots_final % 2 != 0: spots_final += 1
                sch_h = calculate_schedule(spots_final, days_count)
                
                base_list = db["量販_全省"]["List"]
                unit_rate_h = int((base_list / base_std) * factor * penalty)
                total_rate_h = unit_rate_h * spots_final
                total_list_accum += total_rate_h
                
                log_details = [f"**量販總價**: ${total_rate_h:,} (單價 ${unit_rate_h:,} x {spots_final})"]
                debug_logs.append({"Media": f"家樂福 ({sec}s)", "Budget": f"${s_budget:,.0f}", "Status": f"執行 {spots_final} 檔", "Details": log_details})
                
                rows.append({"media": m, "region": "全省量販", "program_num": STORE_COUNTS_NUM["家樂福_量販"], "daypart": db["量販_全省"]["Day_Part"], "seconds": sec, "spots": spots_final, "schedule": sch_h, "rate_display": total_rate_h, "pkg_display": total_rate_h, "is_pkg_member": False})
                
                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({"media": m, "region": "全省超市", "program_num": STORE_COUNTS_NUM["家樂福_超市"], "daypart": db["超市_全省"]["Day_Part"], "seconds": sec, "spots": spots_s, "schedule": sch_s, "rate_display": "計量販", "pkg_display": "計量販", "is_pkg_member": False})

    return rows, total_list_accum, debug_logs

# =========================================================
# 6. OpenPyXL 規格重建引擎
# =========================================================
FONT_MAIN = "微軟正黑體"
SIDE_THIN = Side(style='thin')
SIDE_MEDIUM = Side(style='medium')
SIDE_THICK = Side(style='medium') # Reusing medium as thick for consistency
SIDE_HAIR = Side(style='hair')

def style_range(ws, cell_range, border=Border(), fill=None, font=None, alignment=None):
    rows = list(ws[cell_range])
    for row in rows:
        for cell in row:
            if border: cell.border = border
            if fill: cell.fill = fill
            if font: cell.font = font
            if alignment: cell.alignment = alignment

def apply_borders(ws, range_string, style='thin'):
    min_col, min_row, max_col, max_row = openpyxl.utils.range_boundaries(range_string)
    border_side = Side(style=style, color="000000")
    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(r, c).border = border

def draw_outer_border(ws, min_r, max_r, min_c, max_c):
    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            cell = ws.cell(r, c)
            new_border = copy(cell.border)
            top = SIDE_MEDIUM if r == min_r else new_border.top
            bottom = SIDE_MEDIUM if r == max_r else new_border.bottom
            left = SIDE_MEDIUM if c == min_c else new_border.left
            right = SIDE_MEDIUM if c == max_c else new_border.right
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

# ----------------- Dongwu Engine (No changes, stable) -----------------
def render_dongwu(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val):
    COL_WIDTHS = {'A': 19.6, 'B': 22.8, 'C': 14.6, 'D': 20.0, 'E': 13.0, 'F': 19.6, 'G': 17.9}
    ROW_HEIGHTS = {1: 61.0, 2: 29.0, 3: 40.0, 4: 40.0, 5: 40.0, 6: 40.0, 7: 40.0, 8: 40.0}
    for k, v in COL_WIDTHS.items(): ws.column_dimensions[k].width = v
    for i in range(8, 40): ws.column_dimensions[get_column_letter(i)].width = 8.5
    ws.column_dimensions['AM'].width = 13.0
    for r, h in ROW_HEIGHTS.items(): ws.row_dimensions[r].height = h
    ws['A1'] = "Media Schedule"; ws.merge_cells("A1:AM1")
    style_range(ws, "A1:AM1", font=Font(name=FONT_MAIN, size=48, bold=True), alignment=Alignment(horizontal='center', vertical='center'))
    for c in range(1, 40): ws.cell(3, c).border = Border(top=SIDE_MEDIUM)
    info_map = {"A3": ("客戶名稱：", client_name), "A4": ("Product：", product_display_str), "A5": ("Period :", f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"), "A6": ("Medium :", "全家廣播/新鮮視/家樂福")}
    for addr, (lbl, val) in info_map.items():
        ws[addr] = lbl; ws[addr].font = Font(name=FONT_MAIN, size=14, bold=True); ws[addr].alignment = Alignment(vertical='center')
        val_cell = ws.cell(ws[addr].row, 2); val_cell.value = val; val_cell.font = Font(name=FONT_MAIN, size=14, bold=True); val_cell.alignment = Alignment(vertical='center')
    ws['H6'] = f"{start_dt.month}月"; ws['H6'].font = Font(name=FONT_MAIN, size=16, bold=True); ws['H6'].alignment = Alignment(horizontal='center', vertical='center')
    headers = [("A","Station"), ("B","Location"), ("C","Program"), ("D","Day-part"), ("E","Size"), ("F","rate\n(Net)"), ("G","Package-cost\n(Net)")]
    for col, txt in headers:
        ws[f"{col}7"] = txt; ws.merge_cells(f"{col}7:{col}8")
        style_range(ws, f"{col}7:{col}8", font=Font(name=FONT_MAIN, size=14), alignment=Alignment(horizontal='center', vertical='center', wrap_text=True), border=Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=SIDE_THIN, right=SIDE_THIN))
    curr = start_dt; eff_days = (end_dt - start_dt).days + 1
    for i in range(31):
        col_idx = 8 + i; d_cell = ws.cell(7, col_idx); w_cell = ws.cell(8, col_idx)
        if i < eff_days:
            d_cell.value = curr; d_cell.number_format = 'm/d'; w_cell.value = ["一","二","三","四","五","六","日"][curr.weekday()]
            if curr.weekday() >= 5: w_cell.fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
            curr += timedelta(days=1)
        d_cell.font = Font(name=FONT_MAIN, size=12); w_cell.font = Font(name=FONT_MAIN, size=12)
        d_cell.alignment = w_cell.alignment = Alignment(horizontal='center', vertical='center')
        d_cell.border = Border(left=SIDE_THIN, right=SIDE_THIN, top=SIDE_MEDIUM, bottom=SIDE_THIN)
        w_cell.border = Border(left=SIDE_THIN, right=SIDE_THIN, bottom=SIDE_MEDIUM, top=SIDE_THIN)
    ws['AM7'] = "檔次"; ws.merge_cells("AM7:AM8")
    style_range(ws, "AM7:AM8", font=Font(name=FONT_MAIN, size=14), alignment=Alignment(horizontal='center', vertical='center'), border=Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=SIDE_THIN, right=SIDE_THIN))
    return render_data_rows(ws, rows, 9, final_budget_val, eff_days, "Dongwu")

# ----------------- Shenghuo Engine (Dynamic Cols) -----------------
def render_shenghuo(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val, prod_cost):
    days_n = (end_dt - start_dt).days + 1
    
    # 1. Config Cols
    # Fixed: A=22.5, B=24.5, C=13.8, D=19.4, E=13
    ws.column_dimensions['A'].width = 22.5
    ws.column_dimensions['B'].width = 24.5
    ws.column_dimensions['C'].width = 13.8
    ws.column_dimensions['D'].width = 19.4
    ws.column_dimensions['E'].width = 13.0
    
    # Date Cols (F to ...)
    for i in range(days_n):
        col_letter = get_column_letter(6 + i)
        ws.column_dimensions[col_letter].width = 13.0
    
    # End Cols (3 cols after dates)
    end_c_start = 6 + days_n
    ws.column_dimensions[get_column_letter(end_c_start)].width = 13.0   # Spots
    ws.column_dimensions[get_column_letter(end_c_start+1)].width = 59.0 # List
    ws.column_dimensions[get_column_letter(end_c_start+2)].width = 13.2 # Net
    
    total_cols = 5 + days_n + 3

    # Row Heights
    ROW_H_MAP = {1:46, 2:46, 3:46, 4:46.5, 5:40, 6:40, 7:40, 8:40}
    for r, h in ROW_H_MAP.items(): ws.row_dimensions[r].height = h
    
    # 2. Header Content
    ws['A3'] = "聲活數位科技股份有限公司 統編 28710100"
    ws['A3'].font = Font(name=FONT_MAIN, size=20); ws['A3'].alignment = Alignment(vertical='center')
    ws['A4'] = "蔡伊閔"
    ws['A4'].font = Font(name=FONT_MAIN, size=16); ws['A4'].alignment = Alignment(vertical='center')

    # Row 5-6 Info (Grey)
    grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    for r in [5, 6]:
        for c in range(1, total_cols + 1):
            cell = ws.cell(r, c)
            cell.fill = grey_fill
            cell.font = Font(name=FONT_MAIN, size=14, bold=True)
            top = SIDE_MEDIUM if r==5 else Side()
            bottom = SIDE_MEDIUM if r==6 else Side()
            left = SIDE_MEDIUM if c==1 else Side()
            right = SIDE_MEDIUM if c==total_cols else Side()
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)

    ws['A5'] = "客戶名稱："; ws['B5'] = client_name
    ws['F5'] = "廣告規格："; ws['H5'] = "20秒/15秒"
    # Date Range at 2nd last col (List Price Col)
    date_range_col = total_cols - 1
    ws.cell(5, date_range_col).value = f"執行期間：: {start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"
    ws.cell(5, date_range_col).alignment = Alignment(horizontal='right', vertical='center')

    ws['A6'] = "廣告名稱："; ws['B6'] = product_display_str
    
    # Month Labels
    ws.cell(6, 6).value = f"{start_dt.month}月"
    for i in range(days_n):
        d = start_dt + timedelta(days=i)
        if d.month != start_dt.month and d.day == 1:
            ws.cell(6, 6+i).value = f"{d.month}月"

    # Row 7 & 8 (Table Header)
    headers = ["頻道", "播出地區", "播出店數", "播出時間", "秒數\n規格"]
    header_blue = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    
    for i, h in enumerate(headers):
        ws.merge_cells(start_row=7, start_column=i+1, end_row=8, end_column=i+1)
        cell = ws.cell(7, i+1); cell.value = h
        style_range(ws, f"{get_column_letter(i+1)}7:{get_column_letter(i+1)}8", 
                    font=Font(name=FONT_MAIN, size=14, bold=True), 
                    alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
                    fill=header_blue,
                    border=Border(top=SIDE_MEDIUM, bottom=SIDE_HAIR, left=SIDE_HAIR, right=SIDE_HAIR))
    
    # A7 Left Medium
    ws.cell(7,1).border = Border(top=SIDE_MEDIUM, left=SIDE_MEDIUM, right=SIDE_HAIR)
    ws.cell(8,1).border = Border(bottom=SIDE_HAIR, left=SIDE_MEDIUM, right=SIDE_HAIR)

    # Date Cols
    curr = start_dt
    for i in range(days_n):
        c = 6 + i
        cell7 = ws.cell(7, c); cell7.value = curr; cell7.number_format = 'd'
        cell7.fill = header_blue
        cell7.font = Font(name=FONT_MAIN, size=14, bold=True); cell7.alignment = Alignment(horizontal='center', vertical='center')
        cell7.border = Border(top=SIDE_MEDIUM, bottom=SIDE_HAIR, left=SIDE_HAIR, right=SIDE_HAIR)
        
        cell8 = ws.cell(8, c); cell8.value = f'=MID("日一二三四五六",WEEKDAY({get_column_letter(c)}7,1),1)'
        cell8.font = Font(name=FONT_MAIN, size=14, bold=True); cell8.alignment = Alignment(horizontal='center', vertical='center')
        cell8.border = Border(top=SIDE_HAIR, bottom=SIDE_HAIR, left=SIDE_HAIR, right=SIDE_HAIR)
        
        curr += timedelta(days=1)

    # End Cols
    end_headers = ["檔次", "定價", "專案價"]
    for i, h in enumerate(end_headers):
        c = end_c_start + i
        ws.merge_cells(start_row=7, start_column=c, end_row=8, end_column=c)
        ws.cell(7, c).value = h
        style_range(ws, f"{get_column_letter(c)}7:{get_column_letter(c)}8",
                    font=Font(name=FONT_MAIN, size=14, bold=True),
                    alignment=Alignment(horizontal='center', vertical='center'),
                    fill=header_blue,
                    border=Border(top=SIDE_MEDIUM, bottom=SIDE_HAIR, left=SIDE_HAIR, right=SIDE_HAIR))
    
    # Last Col Right Medium
    ws.cell(7, total_cols).border = Border(top=SIDE_MEDIUM, right=SIDE_MEDIUM, left=SIDE_HAIR)
    ws.cell(8, total_cols).border = Border(bottom=SIDE_HAIR, right=SIDE_MEDIUM, left=SIDE_HAIR)

    return render_data_rows(ws, rows, 9, final_budget_val, days_n, "Shenghuo")

# ----------------- Bolin Engine (Dynamic Cols + B Start) -----------------
def render_bolin(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val, prod_cost):
    days_n = (end_dt - start_dt).days + 1
    total_cols = 1 + 5 + days_n + 3 # Spacer(A) + Fixed(B-F) + Dates + End(3)
    
    # Col Widths
    ws.column_dimensions['A'].width = 1.76 # Spacer
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 10
    
    # Date Cols (G to ...)
    for i in range(days_n):
        col_letter = get_column_letter(7 + i)
        ws.column_dimensions[col_letter].width = 5
        
    # End Cols
    end_c_start = 7 + days_n
    ws.column_dimensions[get_column_letter(end_c_start)].width = 8
    ws.column_dimensions[get_column_letter(end_c_start+1)].width = 12
    ws.column_dimensions[get_column_letter(end_c_start+2)].width = 12

    # Row Heights
    # Bolin doesn't use A1 huge title anymore
    ROW_H_MAP = {1:15, 2:25, 3:25, 4:25, 5:25, 6:25, 7:35} # Row 1 spacer? No, start from Row 2 data?
    # Spec says: A3, A4... wait, start from B3/B4.
    # Let's map Rows 3-6 as Header Info.
    for r in range(1, 8): ws.row_dimensions[r].height = 25
    ws.row_dimensions[7].height = 35 # Header

    # Meta Info (Left Labels at B, Values at C)
    # TO
    ws['B2'] = "TO："; ws['B2'].font = Font(name=FONT_MAIN, size=13, bold=True); ws['B2'].alignment = Alignment(horizontal='right')
    ws['C2'] = client_name; ws['C2'].font = Font(name=FONT_MAIN, size=13)
    
    # FROM
    ws['B3'] = "FROM："; ws['B3'].font = Font(name=FONT_MAIN, size=13, bold=True); ws['B3'].alignment = Alignment(horizontal='right')
    ws['C3'] = "鉑霖行動行銷 許雅婷 TINA"; ws['C3'].font = Font(name=FONT_MAIN, size=13)
    
    # Client
    ws['B4'] = "客戶名稱："; ws['B4'].font = Font(name=FONT_MAIN, size=13, bold=True); ws['B4'].alignment = Alignment(horizontal='right')
    ws['C4'] = client_name; ws['C4'].font = Font(name=FONT_MAIN, size=13)
    
    # Product
    ws['B5'] = "廣告名稱："; ws['B5'].font = Font(name=FONT_MAIN, size=13, bold=True); ws['B5'].alignment = Alignment(horizontal='right')
    ws['C5'] = product_display_str; ws['C5'].font = Font(name=FONT_MAIN, size=13)

    # Right Side Info
    # Spec: G4
    ws['G4'] = "廣告規格："; ws['G4'].font = Font(name=FONT_MAIN, size=13, bold=True)
    ws['H4'] = "20秒/15秒"; ws['H4'].font = Font(name=FONT_MAIN, size=13)
    
    # Date Range at End-1
    dr_lbl_col = end_c_start + 1 # End col is Total Price, so maybe shift left?
    # Use explicit col for date range? Let's put it near end.
    # Spec said "AE4" in fixed 31 layout. Dynamic: End-1 works.
    date_lbl_col = total_cols - 2
    date_val_col = total_cols - 1
    ws.cell(4, date_lbl_col).value = "執行期間："; ws.cell(4, date_lbl_col).font = Font(name=FONT_MAIN, size=13, bold=True)
    ws.cell(4, date_val_col).value = f"{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')}"; ws.cell(4, date_val_col).font = Font(name=FONT_MAIN, size=13)

    # Row 7 Table Header
    header_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    
    # Fixed Cols B-F
    headers = ["頻道", "播出地區", "播出店數", "播出時間", "規格"]
    for i, h in enumerate(headers):
        c = 2 + i # Start B=2
        cell = ws.cell(7, c); cell.value = h
        cell.fill = header_fill
        cell.font = Font(name=FONT_MAIN, size=12, bold=True); cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=SIDE_THIN, right=SIDE_THIN)
        if c==2: cell.border = Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=SIDE_MEDIUM, right=SIDE_THIN)

    # Date Cols
    curr = start_dt
    for i in range(days_n):
        c = 7 + i
        cell = ws.cell(7, c); cell.value = curr; cell.number_format = 'm/d'
        cell.fill = header_fill
        cell.font = Font(name=FONT_MAIN, size=10, bold=True); cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=SIDE_THIN, right=SIDE_THIN)
        curr += timedelta(days=1)

    # End Cols
    end_h = ["總檔次", "單價", "金額"]
    for i, h in enumerate(end_h):
        c = end_c_start + i
        cell = ws.cell(7, c); cell.value = h
        cell.fill = header_fill
        cell.font = Font(name=FONT_MAIN, size=12, bold=True); cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=SIDE_THIN, right=SIDE_THIN)
        if i==2: cell.border = Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=SIDE_THIN, right=SIDE_MEDIUM)

    return render_data_rows(ws, rows, 8, final_budget_val, days_n, "Bolin")

# Common Data Renderer
def render_data_rows(ws, rows, start_row, final_budget_val, eff_days, mode):
    curr_row = start_row
    font_content = Font(name=FONT_MAIN, size=14 if mode in ["Dongwu","Shenghuo"] else 12)
    row_height = 40 if mode in ["Dongwu","Shenghuo"] else 25

    grouped_data = {
        "全家廣播": sorted([r for r in rows if r["media"] == "全家廣播"], key=lambda x: x["seconds"]),
        "新鮮視": sorted([r for r in rows if r["media"] == "新鮮視"], key=lambda x: x["seconds"]),
        "家樂福": sorted([r for r in rows if r["media"] == "家樂福"], key=lambda x: x["seconds"]),
    }

    # Define Max Col based on mode
    if mode == "Dongwu": max_c = 39
    elif mode == "Shenghuo": max_c = 5 + eff_days + 3
    else: max_c = 1 + 5 + eff_days + 3 # Bolin

    for m_key, data in grouped_data.items():
        if not data: continue
        start_merge_row = curr_row
        
        # Section Top Border (Medium)
        start_c = 1 if mode != "Bolin" else 2
        for c in range(start_c, max_c + 1):
            cell = ws.cell(curr_row, c)
            l = SIDE_MEDIUM if c==start_c else SIDE_THIN if mode != "Shenghuo" else SIDE_HAIR
            r = SIDE_MEDIUM if c==max_c else SIDE_THIN if mode != "Shenghuo" else SIDE_HAIR
            cell.border = Border(top=SIDE_MEDIUM, left=Side(style=l), right=Side(style=r), bottom=SIDE_THIN if mode!="Shenghuo" else SIDE_HAIR)

        display_name = f"全家便利商店\n{m_key if m_key!='家樂福' else ''}廣告"
        if m_key == "家樂福": display_name = "家樂福"
        elif m_key == "全家廣播": display_name = "全家便利商店\n通路廣播廣告"
        elif m_key == "新鮮視": display_name = "全家便利商店\n新鮮視廣告"

        for idx, r_data in enumerate(data):
            ws.row_dimensions[curr_row].height = row_height
            
            # Fixed Cols
            base_c = 1 if mode != "Bolin" else 2
            ws.cell(curr_row, base_c).value = display_name
            ws.cell(curr_row, base_c+1).value = r_data["region"]
            ws.cell(curr_row, base_c+2).value = int(r_data.get("program_num", 0))
            ws.cell(curr_row, base_c+3).value = r_data["daypart"]
            ws.cell(curr_row, base_c+4).value = f"{r_data['seconds']}秒"
            
            rate_val = r_data["rate_display"]; pkg_val = r_data["pkg_display"]
            if r_data.get("is_pkg_member") and idx == 0: pkg_val = r_data["nat_pkg_display"]
            elif r_data.get("is_pkg_member"): pkg_val = ""

            if mode == "Dongwu":
                ws.cell(curr_row, 6).value = rate_val; ws.cell(curr_row, 7).value = pkg_val
                sch_start_col = 8; total_col = 39
            elif mode == "Shenghuo":
                sch_start_col = 6
                ws.cell(curr_row, 5+eff_days+2).value = rate_val
                ws.cell(curr_row, 5+eff_days+3).value = pkg_val
                total_col = 5+eff_days+1
            else: # Bolin
                sch_start_col = 7
                ws.cell(curr_row, 1+5+eff_days+2).value = rate_val 
                ws.cell(curr_row, 1+5+eff_days+3).value = pkg_val 
                total_col = 1+5+eff_days+1

            sch = r_data["schedule"]; row_sum = 0
            for d_idx in range(eff_days): 
                col_idx = sch_start_col + d_idx
                if d_idx < len(sch):
                    val = sch[d_idx]
                    ws.cell(curr_row, col_idx).value = val; row_sum += val
                
                # Weekend Color (Shenghuo Only)
                if mode == "Shenghuo":
                     # Shenghuo header is Row 7
                     header_date = ws.cell(7, col_idx).value
                     if isinstance(header_date, (datetime, date)) and header_date.weekday() >= 5:
                         ws.cell(curr_row, col_idx).fill = PatternFill(start_color="FFFFD966", end_color="FFFFD966", fill_type="solid")

            ws.cell(curr_row, total_col).value = row_sum

            # Styles
            for c in range(start_c, max_c + 1):
                cell = ws.cell(curr_row, c)
                cell.font = font_content
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                
                t_style = cell.border.top.style if (cell.border.top and cell.border.top.style) else 'thin'
                
                l_style = 'thin'; r_style = 'thin'; b_style = 'thin'
                if mode == "Shenghuo":
                    l_style = 'hair'; r_style = 'hair'; b_style = 'hair'
                    if c==start_c: l_style = 'medium'
                    if c==max_c: r_style = 'medium'
                elif mode == "Bolin":
                    if c==start_c: l_style = 'medium'
                    if c==max_c: r_style = 'medium'
                else: # Dongwu
                    if c==start_c: l_style = 'medium'
                    if c==max_c: r_style = 'medium'

                cell.border = Border(left=Side(style=l_style), right=Side(style=r_style), top=Side(style=t_style), bottom=Side(style=b_style))
                
                if isinstance(cell.value, (int, float)): 
                    cell.number_format = "#,##0_);[Red](#,##0)" if mode=="Shenghuo" else "#,##0"
            curr_row += 1

        if curr_row > start_merge_row:
            ws.merge_cells(start_row=start_merge_row, start_column=start_c, end_row=curr_row-1, end_column=start_c)
        
        # Merge Package Cost
        if data[0].get("is_pkg_member"):
            p_c = 7 if mode == "Dongwu" else (5+eff_days+3) if mode == "Shenghuo" else (1+5+eff_days+3)
            ws.merge_cells(start_row=start_merge_row, start_column=p_c, end_row=curr_row-1, end_column=p_c)
        
        if mode == "Dongwu":
            for col_idx in [4, 5]:
                m_start = start_merge_row
                while m_start < curr_row:
                    m_end = m_start; curr_val = ws.cell(m_start, col_idx).value
                    while m_end + 1 < curr_row:
                        if ws.cell(m_end + 1, col_idx).value == curr_val: m_end += 1
                        else: break
                    if m_end > m_start: ws.merge_cells(start_row=m_start, start_column=col_idx, end_row=m_end, end_column=col_idx)
                    m_start = m_end + 1

        # Bottom Medium for Section End
        for c in range(start_c, max_c + 1):
            cell = ws.cell(curr_row-1, c)
            existing_l = cell.border.left.style if (cell.border.left and cell.border.left.style) else 'thin'
            existing_r = cell.border.right.style if (cell.border.right and cell.border.right.style) else 'thin'
            existing_t = cell.border.top.style if (cell.border.top and cell.border.top.style) else 'thin'
            
            cell.border = Border(top=Side(style=existing_t), bottom=SIDE_MEDIUM, left=Side(style=existing_l), right=Side(style=existing_r))

    # Total Row
    ws.row_dimensions[curr_row].height = 40 if mode=="Shenghuo" else 30
    
    label_col = 6 if mode == "Dongwu" else 5 if mode == "Shenghuo" else 1+5+eff_days+2 # Bolin Total at AM(End-1)?
    # Bolin Spec: "Total" in AM (End-1), Value in AN (End)
    
    total_val_col = 7 if mode == "Dongwu" else 5+eff_days+3 if mode == "Shenghuo" else 1+5+eff_days+3
    
    ws.cell(curr_row, label_col).value = "Total"; ws.cell(curr_row, label_col).alignment = Alignment(horizontal='right', vertical='center')
    ws.cell(curr_row, label_col).font = Font(name=FONT_MAIN, size=14 if mode!="Bolin" else 12, bold=True)
    ws.cell(curr_row, total_val_col).value = final_budget_val; ws.cell(curr_row, total_val_col).number_format = "#,##0"
    ws.cell(curr_row, total_val_col).font = Font(name=FONT_MAIN, size=14 if mode!="Bolin" else 12, bold=True); ws.cell(curr_row, total_val_col).alignment = Alignment(horizontal='center', vertical='center')

    # Daily Sums
    total_spots_all = 0
    sch_start = 8 if mode == "Dongwu" else 6 if mode == "Shenghuo" else 7
    spot_sum_col = 39 if mode == "Dongwu" else 5+eff_days+1 if mode == "Shenghuo" else 1+5+eff_days+1
    
    for d_idx in range(eff_days):
        col_idx = sch_start + d_idx
        s_sum = sum([r["schedule"][d_idx] for r in rows if d_idx < len(r["schedule"])])
        ws.cell(curr_row, col_idx).value = s_sum; total_spots_all += s_sum
        ws.cell(curr_row, col_idx).number_format = "#,##0_);[Red](#,##0)" if mode=="Shenghuo" else "#,##0"
        ws.cell(curr_row, col_idx).font = Font(name=FONT_MAIN, size=14 if mode!="Bolin" else 12, bold=True)
        ws.cell(curr_row, col_idx).alignment = Alignment(horizontal='center', vertical='center')
    
    ws.cell(curr_row, spot_sum_col).value = total_spots_all; ws.cell(curr_row, spot_sum_col).font = Font(name=FONT_MAIN, size=14 if mode!="Bolin" else 12, bold=True)
    ws.cell(curr_row, spot_sum_col).alignment = Alignment(horizontal='center', vertical='center')
    
    # Total Row Style
    start_c = 1 if mode != "Bolin" else 2
    for c in range(start_c, max_c + 1):
        cell = ws.cell(curr_row, c)
        l = SIDE_MEDIUM if c==start_c else SIDE_THIN if mode!="Shenghuo" else SIDE_HAIR
        r = SIDE_MEDIUM if c==max_c else SIDE_THIN if mode!="Shenghuo" else SIDE_HAIR
        if mode == "Bolin": l = SIDE_MEDIUM if c==start_c else SIDE_THIN; r = SIDE_MEDIUM if c==max_c else SIDE_THIN
        cell.border = Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=Side(style=l), right=Side(style=r))
        if mode == "Dongwu" and c==1: cell.border = Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=SIDE_MEDIUM, right=SIDE_THIN) # Dongwu Left
    
    return curr_row

def generate_excel_from_scratch(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val, prod_cost):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "工作表1"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE; ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    
    if format_type == "Dongwu": curr_row = render_dongwu(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val)
    elif format_type == "Shenghuo": curr_row = render_shenghuo(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val, prod_cost)
    else: curr_row = render_bolin(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val, prod_cost)

    if format_type == "Dongwu":
        curr_row += 1
        vat = int(round(final_budget_val * 0.05)); grand_total = final_budget_val + vat
        footer_data = [("製作", prod_cost), ("5% VAT", vat), ("Grand Total", grand_total)]
        label_col = 6; val_col = 7
        for label, val in footer_data:
            ws.row_dimensions[curr_row].height = 30
            ws.cell(curr_row, label_col).value = label; ws.cell(curr_row, label_col).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(curr_row, label_col).font = Font(name=FONT_MAIN, size=14)
            ws.cell(curr_row, val_col).value = val; ws.cell(curr_row, val_col).number_format = "#,##0"; ws.cell(curr_row, val_col).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(curr_row, val_col).font = Font(name=FONT_MAIN, size=14)
            ws.cell(curr_row, label_col).border = Border(left=SIDE_THICK, top=SIDE_THIN, bottom=SIDE_THIN, right=SIDE_THIN)
            ws.cell(curr_row, val_col).border = Border(right=SIDE_THICK, top=SIDE_THIN, bottom=SIDE_THIN, left=SIDE_THIN)
            if label == "Grand Total":
                for c in range(1, 40): ws.cell(curr_row, c).fill = PatternFill(start_color="FFC107", end_color="FFC107", fill_type="solid"); ws.cell(curr_row, c).border = Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM)
            curr_row += 1
        draw_outer_border(ws, 7, curr_row-1, 1, 39)

    if format_type == "Dongwu":
        curr_row += 1
        ws.cell(curr_row, 1).value = "Remarks："
        ws.cell(curr_row, 1).font = Font(name=FONT_MAIN, size=16, bold=True, underline="single", color="000000")
        for c in range(1, 40): ws.cell(curr_row, c).border = Border(top=Side(style=None))
        curr_row += 1
        for rm in remarks_list:
            ws.cell(curr_row, 1).value = rm
            f_color = "FF0000" if (rm.strip().startswith("1.") or rm.strip().startswith("4.")) else "000000"
            ws.cell(curr_row, 1).font = Font(name=FONT_MAIN, size=14, color=f_color)
            curr_row += 1
    elif format_type == "Shenghuo":
        curr_row += 1
        ws.cell(curr_row, 1).value = "Remarks："
        ws.cell(curr_row, 1).font = Font(name=FONT_MAIN, size=14, bold=True, underline="single", color="000000")
        curr_row += 1
        for rm in remarks_list:
            ws.cell(curr_row, 1).value = rm
            f_color = "FF0000" if (rm.strip().startswith("1.") or rm.strip().startswith("4.")) else "000000"
            ws.cell(curr_row, 1).font = Font(name=FONT_MAIN, size=14, color=f_color)
            curr_row += 1
    elif format_type == "Bolin":
        curr_row += 1
        ws.cell(curr_row, 2).value = "Remarks：" # I col in original (9), here B=2? No, I is 9.
        # Original spec said "Remarks at I". But here we have dynamic cols.
        # Let's put at B.
        ws.cell(curr_row, 9).value = "Remarks：" # I
        ws.cell(curr_row, 9).font = Font(name=FONT_MAIN, size=16, bold=True, underline="single")
        curr_row += 1
        for rm in remarks_list:
            ws.cell(curr_row, 9).value = rm
            ws.cell(curr_row, 9).font = Font(name=FONT_MAIN, size=16, bold=True)
            curr_row += 1

    out = io.BytesIO(); wb.save(out); return out.getvalue()

def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks, total_list, grand_total, budget, prod):
    header_cls = "bg-dw-head" if format_type == "Dongwu" else "bg-sh-head"
    if format_type == "Bolin": header_cls = "bg-bolin-head"
    eff_days = min(days_cnt, 31)
    font_b64 = load_font_base64()
    font_face = f"@font-face {{ font-family: 'NotoSansTC'; src: url(data:font/ttf;base64,{font_b64}) format('truetype'); }}" if font_b64 else ""

    date_th1 = ""; date_th2 = ""; curr = start_dt; weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    for i in range(eff_days):
        wd = curr.weekday(); bg = "bg-weekend" if (format_type == "Dongwu" and wd >= 5) else header_cls
        if format_type in ["Shenghuo", "Bolin"]: bg = header_cls 
        date_th1 += f"<th class='{bg} col_day'>{curr.day}</th>"; date_th2 += f"<th class='{bg} col_day'>{weekdays[wd]}</th>"; curr += timedelta(days=1)

    cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(Net)", "Package-cost<br>(Net)"] if format_type == "Dongwu" else ["頻道", "播出地區", "播出店數", "播出時間", "秒數<br>規格", "單價", "金額"]
    th_fixed = "".join([f"<th rowspan='2' class='{header_cls}'>{c}</th>" for c in cols_def])
    
    rows_sorted = sorted(rows, key=lambda x: ({"全家廣播":1,"新鮮視":2,"家樂福":3}.get(x["media"],9), x["seconds"]))
    tbody = ""
    grouped_rows = {}
    for r in rows_sorted: key = (r['media'], r['seconds']); grouped_rows.setdefault(key, []).append(r)

    for (m, sec), group in grouped_rows.items():
        is_nat = group[0].get('is_pkg_member', False); group_size = len(group)
        for k, r_data in enumerate(group):
            tbody += "<tr>"
            if k == 0:
                d_name = "全家便利商店<br>通路廣播廣告" if m == "全家廣播" else "全家便利商店<br>新鮮視廣告" if m == "新鮮視" else "家樂福"
                if format_type in ["Shenghuo", "Bolin"] and m == "全家廣播": d_name = "全家便利商店<br>廣播通路廣告"
                tbody += f"<td class='left' rowspan='{group_size}'>{d_name}</td>"
            tbody += f"<td>{region_display(r_data['region'])}</td><td class='right'>{r_data.get('program_num','')}</td><td>{r_data['daypart']}</td><td>{r_data['seconds']}秒</td>"
            rate = f"{r_data['rate_display']:,}" if isinstance(r_data['rate_display'], int) else r_data['rate_display']
            pkg = f"{r_data['pkg_display']:,}" if isinstance(r_data['pkg_display'], int) else r_data['pkg_display']
            tbody += f"<td class='right'>{rate}</td>"
            if is_nat:
                if k == 0: tbody += f"<td class='right' rowspan='{group_size}'>{r_data['nat_pkg_display']:,}</td>"
            else: tbody += f"<td class='right'>{pkg}</td>"
            for d in r_data['schedule'][:eff_days]: tbody += f"<td>{d}</td>"
            tbody += f"<td class='bg-total'>{r_data['spots']}</td></tr>"

    totals = [sum([r["schedule"][d] for r in rows if d < len(r["schedule"])]) for d in range(eff_days)]
    colspan = 5; empty_td = "<td></td>" if format_type == "Dongwu" else ""
    if format_type != "Dongwu": empty_td = ""
    tfoot = f"<tr class='bg-total'><td colspan='{colspan}' class='right'>Total (List Price)</td>{empty_td}<td class='right'>{total_list:,}</td>"
    for t in totals: tfoot += f"<td>{t}</td>"
    tfoot += f"<td>{sum(totals)}</td></tr>"

    vat = int(round(budget * 0.05))
    footer_rows = f"<tr><td colspan='6' class='right'>製作</td><td class='right'>{prod:,}</td><td colspan='{eff_days+1}'></td></tr>"
    footer_rows += f"<tr><td colspan='6' class='right'>專案優惠價 (Budget)</td><td class='right' style='color:red; font-weight:bold;'>{budget:,}</td><td colspan='{eff_days+1}'></td></tr>"
    footer_rows += f"<tr><td colspan='6' class='right'>5% VAT</td><td class='right'>{vat:,}</td><td colspan='{eff_days+1}'></td></tr>"
    footer_rows += f"<tr class='bg-grand'><td colspan='6' class='right'>Grand Total</td><td class='right'>{grand_total:,}</td><td colspan='{eff_days+1}'></td></tr>"

    return f"""<html><head><style>
    {font_face}
    body {{ font-family: 'NotoSansTC', sans-serif !important; font-size: 10px; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ border: 0.5pt solid #000; padding: 2px; text-align: center; white-space: nowrap; }}
    .bg-dw-head {{ background-color: #4472C4; color: white; -webkit-print-color-adjust: exact; }}
    .bg-sh-head {{ background-color: #BDD7EE; color: black; -webkit-print-color-adjust: exact; }}
    .bg-bolin-head {{ background-color: #F8CBAD; color: black; -webkit-print-color-adjust: exact; }}
    .bg-weekend {{ background-color: #FFD966; -webkit-print-color-adjust: exact; }}
    .bg-total   {{ background-color: #E2EFDA; -webkit-print-color-adjust: exact; }}
    .bg-grand   {{ background-color: #FFC107; -webkit-print-color-adjust: exact; }}
    .left {{ text-align: left; }} .right {{ text-align: right; }}
    .remarks {{ margin-top: 10px; font-size: 9px; text-align: left; white-space: pre-wrap; }}
    </style></head><body>
    <div style="margin-bottom:10px;">
        <div style="font-size:16px; font-weight:bold; text-align:center;">Media Schedule</div>
        <b>客戶名稱：</b>{html_escape(c_name)} &nbsp; <b>Product：</b>{html_escape(p_display)}<br>
        <b>Period：</b>{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')} &nbsp; <b>Medium：</b>全家廣播/新鮮視/家樂福
    </div>
    <table><thead><tr>{th_fixed}{date_th1}<th class='{header_cls}' rowspan='2'>檔次</th></tr><tr>{date_th2}</tr></thead>
    <tbody>{tbody}{tfoot}{footer_rows}</tbody></table>
    <div class="remarks"><b>Remarks：</b><br>{"<br>".join([html_escape(x) for x in remarks])}</div></body></html>"""

def load_font_base64():
    font_path = "NotoSansTC-Regular.ttf"
    if os.path.exists(font_path):
        with open(font_path, "rb") as f: return base64.b64encode(f.read()).decode("utf-8")
    try:
        r = requests.get("https://github.com/googlefonts/noto-cjk/raw/main/Sans/TTF/TraditionalChinese/NotoSansTC-Regular.ttf", timeout=15)
        if r.status_code == 200:
            with open(font_path, "wb") as f: f.write(r.content)
            return base64.b64encode(r.content).decode("utf-8")
    except: pass
    return None

with st.sidebar:
    st.header("🕵️ 主管登入")
    if not st.session_state.is_supervisor:
        pwd = st.text_input("輸入密碼", type="password", key="pwd_input")
        if st.button("登入"):
            if pwd == "1234": st.session_state.is_supervisor = True; st.rerun()
            else: st.error("密碼錯誤")
    else:
        st.success("✅ 目前狀態：主管模式"); 
        if st.button("登出"): st.session_state.is_supervisor = False; st.rerun()

st.title("📺 媒體 Cue 表生成器 (v91.0)")
format_type = st.radio("選擇格式", ["Dongwu", "Shenghuo", "Bolin"], horizontal=True)

c1, c2, c3, c4 = st.columns(4)
with c1: client_name = st.text_input("客戶名稱", "萬國通路")
with c2: product_name = st.text_input("產品名稱", "統一布丁")
with c3: total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)
with c4: prod_cost_input = st.number_input("製作費 (未稅)", value=0, step=1000)

final_budget_val = total_budget_input
if st.session_state.is_supervisor:
    st.markdown("---")
    col_sup1, col_sup2 = st.columns([1, 2])
    with col_sup1: st.error("🔒 [主管] 專案優惠價覆寫")
    with col_sup2:
        override_val = st.number_input("輸入最終成交價", value=total_budget_input)
        if override_val != total_budget_input: final_budget_val = override_val; st.caption(f"⚠️ 使用 ${final_budget_val:,} 結算")
    st.markdown("---")

c5, c6 = st.columns(2)
with c5: start_date = st.date_input("開始日", datetime(2026, 1, 1))
with c6: end_date = st.date_input("結束日", datetime(2026, 1, 31))
days_count = (end_date - start_date).days + 1
st.info(f"📅 走期共 **{days_count}** 天")

with st.expander("📝 備註欄位設定", expanded=False):
    rc1, rc2, rc3 = st.columns(3)
    sign_deadline = rc1.date_input("回簽截止日", datetime.now() + timedelta(days=3))
    billing_month = rc2.text_input("請款月份", "2026年2月")
    payment_date = rc3.date_input("付款兌現日", datetime(2026, 3, 31))

st.markdown("### 3. 媒體投放設定")
if "rad_share" not in st.session_state: st.session_state.rad_share = 100
if "fv_share" not in st.session_state: st.session_state.fv_share = 0
if "cf_share" not in st.session_state: st.session_state.cf_share = 0

def on_media_change():
    active = []
    if st.session_state.get("cb_rad"): active.append("rad_share")
    if st.session_state.get("cb_fv"): active.append("fv_share")
    if st.session_state.get("cb_cf"): active.append("cf_share")
    if not active: return
    share = 100 // len(active)
    for key in active: st.session_state[key] = share
    rem = 100 - sum([st.session_state[k] for k in active])
    st.session_state[active[0]] += rem

def on_slider_change(changed_key):
    active = []
    if st.session_state.get("cb_rad"): active.append("rad_share")
    if st.session_state.get("cb_fv"): active.append("fv_share")
    if st.session_state.get("cb_cf"): active.append("cf_share")
    others = [k for k in active if k != changed_key]
    if not others: st.session_state[changed_key] = 100
    elif len(others) == 1:
        val = st.session_state[changed_key]
        st.session_state[others[0]] = max(0, 100 - val)
    elif len(others) == 2:
        val = st.session_state[changed_key]
        rem = max(0, 100 - val)
        k1, k2 = others[0], others[1]
        sum_others = st.session_state[k1] + st.session_state[k2]
        if sum_others == 0: st.session_state[k1] = rem // 2; st.session_state[k2] = rem - st.session_state[k1]
        else:
            ratio = st.session_state[k1] / sum_others
            st.session_state[k1] = int(rem * ratio)
            st.session_state[k2] = rem - st.session_state[k1]

col_cb1, col_cb2, col_cb3 = st.columns(3)
with col_cb1: is_rad = st.checkbox("全家廣播", value=True, key="cb_rad", on_change=on_media_change)
with col_cb2: is_fv = st.checkbox("新鮮視", value=False, key="cb_fv", on_change=on_media_change)
with col_cb3: is_cf = st.checkbox("家樂福", value=False, key="cb_cf", on_change=on_media_change)

m1, m2, m3 = st.columns(3)
config = {}
if is_rad:
    with m1:
        st.markdown("#### 📻 全家廣播")
        is_nat = st.checkbox("全省聯播", True, key="rad_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=REGIONS_ORDER, key="rad_reg")
        if not is_nat and len(regs) == 6: is_nat = True; regs = ["全省"]; st.info("✅ 已選滿6區，自動轉為全省聯播")
        secs = st.multiselect("秒數", DURATIONS, [20], key="rad_sec")
        st.slider("預算 %", 0, 100, key="rad_share", on_change=on_slider_change, args=("rad_share",))
        sec_shares = {}
        if len(secs) > 1:
            rem = 100; sorted_secs = sorted(secs)
            for i, s in enumerate(sorted_secs):
                if i < len(sorted_secs) - 1: v = st.slider(f"{s}秒 %", 0, rem, int(rem/2), key=f"rs_{s}"); sec_shares[s] = v; rem -= v
                else: sec_shares[s] = rem
        elif secs: sec_shares[secs[0]] = 100
        config["全家廣播"] = {"is_national": is_nat, "regions": regs, "sec_shares": sec_shares, "share": st.session_state.rad_share}

if is_fv:
    with m2:
        st.markdown("#### 📺 新鮮視")
        is_nat = st.checkbox("全省聯播", False, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區"], key="fv_reg")
        if not is_nat and len(regs) == 6: is_nat = True; regs = ["全省"]; st.info("✅ 已選滿6區，自動轉為全省聯播")
        secs = st.multiselect("秒數", DURATIONS, [10], key="fv_sec")
        st.slider("預算 %", 0, 100, key="fv_share", on_change=on_slider_change, args=("fv_share",))
        sec_shares = {}
        if len(secs) > 1:
            rem = 100; sorted_secs = sorted(secs)
            for i, s in enumerate(sorted_secs):
                if i < len(sorted_secs) - 1: v = st.slider(f"{s}秒 %", 0, rem, int(rem/2), key=f"fs_{s}"); sec_shares[s] = v; rem -= v
                else: sec_shares[s] = rem
        elif secs: sec_shares[secs[0]] = 100
        config["新鮮視"] = {"is_national": is_nat, "regions": regs, "sec_shares": sec_shares, "share": st.session_state.fv_share}

if is_cf:
    with m3:
        st.markdown("#### 🛒 家樂福")
        secs = st.multiselect("秒數", DURATIONS, [20], key="cf_sec")
        st.slider("預算 %", 0, 100, key="cf_share", on_change=on_slider_change, args=("cf_share",))
        sec_shares = {}
        if len(secs) > 1:
            rem = 100; sorted_secs = sorted(secs)
            for i, s in enumerate(sorted_secs):
                if i < len(sorted_secs) - 1: v = st.slider(f"{s}秒 %", 0, rem, int(rem/2), key=f"cs_{s}"); sec_shares[s] = v; rem -= v
                else: sec_shares[s] = rem
        elif secs: sec_shares[secs[0]] = 100
        config["家樂福"] = {"regions": ["全省"], "sec_shares": sec_shares, "share": st.session_state.cf_share}

if config:
    rows, total_list_accum, logs = calculate_plan_data(config, total_budget_input, days_count)
    prod_cost = prod_cost_input 
    vat = int(round(final_budget_val * 0.05))
    grand_total = final_budget_val + vat
    p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}"
    rem = get_remarks_text(sign_deadline, billing_month, payment_date)
    html_preview = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem, total_list_accum, grand_total, final_budget_val, prod_cost)
    st.components.v1.html(html_preview, height=700, scrolling=True)
    with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
        for log in logs:
            st.markdown(f"### {log['Media']}"); st.markdown(f"- **預算**: {log['Budget']}"); st.markdown(f"- **狀態**: {log['Status']}")
            if 'Details' in log:
                for detail in log['Details']: st.info(detail)
            st.divider()
    col_dl1, col_dl2 = st.columns(2)
    with col_dl2:
        try:
            xlsx_temp = generate_excel_from_scratch(format_type, start_date, end_date, client_name, p_str, rows, rem, final_budget_val, prod_cost)
            pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx_temp)
            if pdf_bytes: st.download_button(f"📥 下載 PDF ({method})", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf", key="pdf_dl")
            else: 
                pdf_bytes, err = html_to_pdf_weasyprint(html_preview)
                if pdf_bytes: st.download_button("📥 下載 PDF (Web)", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf", key="pdf_dl_web")
        except: pass
    with col_dl1:
        if st.session_state.is_supervisor:
            if rows:
                try:
                    xlsx = generate_excel_from_scratch(format_type, start_date, end_date, client_name, p_str, rows, rem, final_budget_val, prod_cost)
                    st.download_button("📥 下載 Excel (主管權限)", xlsx, f"Cue_{safe_filename(client_name)}.xlsx", key="xlsx_dl")
                except Exception as e: st.error(f"Excel Error: {e}")
        else: st.info("🔒 Excel 下載功能僅限主管使用")
