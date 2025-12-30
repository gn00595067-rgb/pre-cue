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
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

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
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v81.1")

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
DEFAULT_ROW_HEIGHT = 20.5
FOOTER_ROW_HEIGHT = 30.0
FONT_MAIN = "微軟正黑體"

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

# ----------------- Dongwu Engine -----------------
def render_dongwu(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val):
    COL_WIDTHS = {'A': 19.6, 'B': 22.8, 'C': 14.6, 'D': 20.0, 'E': 13.0, 'F': 19.6, 'G': 17.9, 'H': 13.0}
    ROW_HEIGHTS = {1: 61.0, 2: 29.0, 3: 18.5, 4: 18.5, 5: 18.5, 6: 19.0, 7: 40.0, 8: 40.0}
    for k, v in COL_WIDTHS.items(): ws.column_dimensions[k].width = v
    for i in range(8, 40): ws.column_dimensions[get_column_letter(i)].width = 13.0
    ws.column_dimensions['AM'].width = 13.0
    for r, h in ROW_HEIGHTS.items(): ws.row_dimensions[r].height = h

    ws['A1'] = "Media Schedule"; ws.merge_cells("A1:AM1")
    style_range(ws, "A1:AM1", font=Font(name=FONT_MAIN, size=48, bold=True), alignment=Alignment(horizontal='center', vertical='center'))
    
    info_map = {"A3": ("客戶名稱：", client_name), "A4": ("Product：", product_display_str), "A5": ("Period :", f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"), "A6": ("Medium :", "全家廣播/新鮮視/家樂福")}
    for addr, (lbl, val) in info_map.items():
        ws[addr] = lbl; ws[addr].font = Font(name=FONT_MAIN, size=14, bold=True); ws[addr].alignment = Alignment(vertical='center')
        val_cell = ws.cell(ws[addr].row, 2); val_cell.value = val; val_cell.font = Font(name=FONT_MAIN, size=14, bold=True); val_cell.alignment = Alignment(vertical='center')

    ws['H6'] = f"{start_dt.month}月"; ws['H6'].font = Font(name=FONT_MAIN, size=16, bold=True); ws['H6'].alignment = Alignment(horizontal='center', vertical='center')

    headers = [("A","Station"), ("B","Location"), ("C","Program"), ("D","Day-part"), ("E","Size"), ("F","rate\n(Net)"), ("G","Package-cost\n(Net)")]
    for col, txt in headers:
        ws[f"{col}7"] = txt; ws.merge_cells(f"{col}7:{col}8")
        style_range(ws, f"{col}7:{col}8", font=Font(name=FONT_MAIN, size=14), alignment=Alignment(horizontal='center', vertical='center', wrap_text=True), border=Border(bottom=Side(style='hair'), top=Side(style='medium')))

    curr = start_dt; eff_days = (end_dt - start_dt).days + 1
    for i in range(31):
        col_idx = 8 + i; d_cell = ws.cell(7, col_idx); w_cell = ws.cell(8, col_idx)
        if i < eff_days:
            d_cell.value = curr; d_cell.number_format = 'm/d'; w_cell.value = ["一","二","三","四","五","六","日"][curr.weekday()]
            if curr.weekday() >= 5: d_cell.fill = w_cell.fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
            curr += timedelta(days=1)
        d_cell.font = Font(name=FONT_MAIN, size=12); w_cell.font = Font(name=FONT_MAIN, size=12)
        d_cell.alignment = w_cell.alignment = Alignment(horizontal='center', vertical='center')
        d_cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='medium'), bottom=Side(style='hair'))
        w_cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='hair'), bottom=Side(style='medium'))

    ws['AM7'] = "檔次"; ws.merge_cells("AM7:AM8")
    style_range(ws, "AM7:AM8", font=Font(name=FONT_MAIN, size=14), alignment=Alignment(horizontal='center', vertical='center'), border=Border(bottom=Side(style='medium'), top=Side(style='medium'), left=Side(style='thin'), right=Side(style='medium')))

    return render_data_rows(ws, rows, 9, final_budget_val, eff_days, "Dongwu")

# ----------------- Shenghuo Engine -----------------
def render_shenghuo(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val):
    COL_WIDTHS = {'A': 20, 'B': 22, 'C': 10, 'D': 15, 'E': 10, 'F': 5}
    ROW_HEIGHTS = {1: 50, 2: 25, 3: 20, 4: 20, 5: 20, 6: 35}
    for k, v in COL_WIDTHS.items(): ws.column_dimensions[k].width = v
    for i in range(7, 38): ws.column_dimensions[get_column_letter(i)].width = 5
    ws.column_dimensions['AL'].width = 8; ws.column_dimensions['AM'].width = 12; ws.column_dimensions['AN'].width = 12
    for r, h in ROW_HEIGHTS.items(): ws.row_dimensions[r].height = h
    
    ws['A1'] = "Media Schedule"; ws.merge_cells("A1:AN1")
    style_range(ws, "A1:AN1", font=Font(name=FONT_MAIN, size=40, bold=True), alignment=Alignment(horizontal='center', vertical='center'))
    
    info_map = {"A3": ("客戶名稱：", client_name), "A4": ("廣告名稱：", product_display_str), "G4": ("廣告規格：", "20秒/15秒"), "AE4": ("執行期間：", f"{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')}")}
    for addr, (lbl, val) in info_map.items():
        ws[addr] = lbl; ws[addr].font = Font(name=FONT_MAIN, size=12, bold=True); ws[addr].alignment = Alignment(vertical='center')
        val_cell = ws.cell(ws[addr].row, ws[addr].column + 1); val_cell.value = val; val_cell.font = Font(name=FONT_MAIN, size=12); val_cell.alignment = Alignment(vertical='center')

    headers = ["頻道", "播出地區", "播出店數", "播出時間", "秒數\n規格"]
    for i, h in enumerate(headers):
        cell = ws.cell(6, i+1); cell.value = h
        cell.font = Font(name=FONT_MAIN, size=13, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        cell.border = Border(top=Side(style='medium'), bottom=Side(style='medium'))

    curr = start_dt; eff_days = (end_dt - start_dt).days + 1
    for i in range(31):
        col_idx = 6 + i; cell = ws.cell(6, col_idx)
        if i < eff_days:
            cell.value = curr; cell.number_format = 'm/d'; curr += timedelta(days=1)
        cell.font = Font(name=FONT_MAIN, size=10); cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(top=Side(style='medium'), bottom=Side(style='medium'))
    
    for i, h in enumerate(["檔次", "定價", "專案價"]):
        cell = ws.cell(6, 37+i); cell.value = h
        cell.font = Font(name=FONT_MAIN, size=13, bold=True)
        cell.fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        cell.border = Border(top=Side(style='medium'), bottom=Side(style='medium'))

    return render_data_rows(ws, rows, 7, final_budget_val, eff_days, "Shenghuo")

# ----------------- Bolin Engine (NEW) -----------------
def render_bolin(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val):
    COL_WIDTHS = {'A': 20, 'B': 22, 'C': 10, 'D': 15, 'E': 10, 'F': 5}
    ROW_HEIGHTS = {1: 60, 2: 25, 3: 25, 4: 25, 5: 25, 6: 25, 7: 35}
    for k, v in COL_WIDTHS.items(): ws.column_dimensions[k].width = v
    for i in range(7, 38): ws.column_dimensions[get_column_letter(i)].width = 5
    ws.column_dimensions['AL'].width = 8; ws.column_dimensions['AM'].width = 12; ws.column_dimensions['AN'].width = 12
    for r, h in ROW_HEIGHTS.items(): ws.row_dimensions[r].height = h
    
    ws['A1'] = "Media Schedule"; ws.merge_cells("A1:AN1")
    style_range(ws, "A1:AN1", font=Font(name=FONT_MAIN, size=42, bold=True), alignment=Alignment(horizontal='center', vertical='center'))
    
    info_map = {
        "A2": ("TO：", client_name), "A3": ("FROM：", "鉑霖行動行銷 許雅婷 TINA"),
        "A4": ("客戶名稱：", client_name), "A5": ("廣告名稱：", product_display_str),
        "G4": ("廣告規格：", "20秒/15秒"), "AE4": ("執行期間：", f"{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')}")
    }
    for addr, (lbl, val) in info_map.items():
        ws[addr] = lbl; ws[addr].font = Font(name=FONT_MAIN, size=13, bold=True)
        val_cell = ws.cell(ws[addr].row, ws[addr].column + 1); val_cell.value = val; val_cell.font = Font(name=FONT_MAIN, size=13)

    headers = ["頻道", "播出地區", "播出店數", "播出時間", "規格"]
    for i, h in enumerate(headers):
        cell = ws.cell(7, i+1); cell.value = h
        cell.font = Font(name=FONT_MAIN, size=12, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = Border(top=Side(style='medium'), bottom=Side(style='medium'))

    curr = start_dt; eff_days = (end_dt - start_dt).days + 1
    for i in range(31):
        col_idx = 6 + i; cell = ws.cell(7, col_idx)
        if i < eff_days:
            cell.value = curr; cell.number_format = 'm/d'; curr += timedelta(days=1)
        cell.font = Font(name=FONT_MAIN, size=10); cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(top=Side(style='medium'), bottom=Side(style='medium'))
    
    for i, h in enumerate(["總檔次", "單價", "金額"]):
        cell = ws.cell(7, 37+i); cell.value = h
        cell.font = Font(name=FONT_MAIN, size=12, bold=True)
        cell.border = Border(top=Side(style='medium'), bottom=Side(style='medium'))

    return render_data_rows(ws, rows, 8, final_budget_val, eff_days, "Bolin")

# Common Data Renderer
def render_data_rows(ws, rows, start_row, final_budget_val, eff_days, mode):
    curr_row = start_row
    grouped_data = {
        "全家廣播": sorted([r for r in rows if r["media"] == "全家廣播"], key=lambda x: x["seconds"]),
        "新鮮視": sorted([r for r in rows if r["media"] == "新鮮視"], key=lambda x: x["seconds"]),
        "家樂福": sorted([r for r in rows if r["media"] == "家樂福"], key=lambda x: x["seconds"]),
    }
    base_font = Font(name=FONT_MAIN, size=12)
    
    for m_key, data in grouped_data.items():
        if not data: continue
        start_merge_row = curr_row
        
        display_name = f"全家便利商店\n{m_key if m_key!='家樂福' else ''}廣告"
        if m_key == "家樂福": display_name = "家樂福"
        elif m_key == "全家廣播": display_name = "全家便利商店\n通路廣播廣告"
        elif m_key == "新鮮視": display_name = "全家便利商店\n新鮮視廣告"

        for idx, r_data in enumerate(data):
            ws.row_dimensions[curr_row].height = 25
            
            ws.cell(curr_row, 1).value = display_name
            ws.cell(curr_row, 2).value = r_data["region"]
            ws.cell(curr_row, 3).value = int(r_data.get("program_num", 0))
            ws.cell(curr_row, 4).value = r_data["daypart"]
            ws.cell(curr_row, 5).value = f"{r_data['seconds']}秒"
            
            rate_val = r_data["rate_display"]
            pkg_val = r_data["pkg_display"]
            if r_data.get("is_pkg_member") and idx == 0: pkg_val = r_data["nat_pkg_display"]
            elif r_data.get("is_pkg_member"): pkg_val = ""

            if mode == "Dongwu":
                ws.cell(curr_row, 6).value = rate_val
                ws.cell(curr_row, 7).value = pkg_val
                sch_start_col = 8; total_col = 39
            else: 
                sch_start_col = 6; total_col = 37
                ws.cell(curr_row, 38).value = rate_val 
                ws.cell(curr_row, 39).value = pkg_val

            sch = r_data["schedule"]; row_sum = 0
            for d_idx in range(31):
                col_idx = sch_start_col + d_idx; cell = ws.cell(curr_row, col_idx)
                if d_idx < len(sch):
                    cell.value = sch[d_idx]; row_sum += sch[d_idx]
            
            ws.cell(curr_row, total_col).value = row_sum

            for c in range(1, ws.max_column + 1):
                cell = ws.cell(curr_row, c)
                cell.font = base_font
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            curr_row += 1

        if curr_row > start_merge_row:
            ws.merge_cells(start_row=start_merge_row, start_column=1, end_row=curr_row-1, end_column=1)
        
        if data[0].get("is_pkg_member"):
            if mode == "Dongwu": ws.merge_cells(start_row=start_merge_row, start_column=7, end_row=curr_row-1, end_column=7)
            else: ws.merge_cells(start_row=start_merge_row, start_column=39, end_row=curr_row-1, end_column=39)

    ws.row_dimensions[curr_row].height = FOOTER_ROW_HEIGHT
    label_col = 6 if mode == "Dongwu" else 36
    total_val_col = 7 if mode == "Dongwu" else 39
    
    ws.cell(curr_row, label_col).value = "Total"
    ws.cell(curr_row, label_col).alignment = Alignment(horizontal='right', vertical='center')
    ws.cell(curr_row, label_col).font = Font(name=FONT_MAIN, size=14, bold=True)
    
    ws.cell(curr_row, total_val_col).value = final_budget_val
    ws.cell(curr_row, total_val_col).number_format = "#,##0"
    ws.cell(curr_row, total_val_col).font = Font(name=FONT_MAIN, size=14, bold=True)
    ws.cell(curr_row, total_val_col).alignment = Alignment(horizontal='center', vertical='center')

    total_spot_col = 39 if mode == "Dongwu" else 37
    total_spots_all = 0
    sch_start = 8 if mode == "Dongwu" else 6
    
    for d_idx in range(31):
        col_idx = sch_start + d_idx
        daily_sum = sum([r["schedule"][d_idx] for r in rows if d_idx < len(r["schedule"])]) if d_idx < eff_days else 0
        ws.cell(curr_row, col_idx).value = daily_sum
        total_spots_all += daily_sum
        ws.cell(curr_row, col_idx).alignment = Alignment(horizontal='center', vertical='center')
    
    ws.cell(curr_row, total_spot_col).value = total_spots_all
    ws.cell(curr_row, total_spot_col).font = Font(name=FONT_MAIN, size=14, bold=True)
    ws.cell(curr_row, total_spot_col).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(curr_row, total_spot_col).border = Border(right=Side(style='thick'), top=Side(style='medium'), bottom=Side(style='medium'))

    total_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    for c in range(1, 40):
        cell = ws.cell(curr_row, c)
        cell.fill = total_fill
    
    return curr_row

def generate_excel_from_scratch(format_type, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val, prod_cost):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工作表1"
    
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    
    if format_type == "Dongwu":
        curr_row = render_dongwu(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val)
    elif format_type == "Shenghuo":
        curr_row = render_shenghuo(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val)
    else: 
        curr_row = render_bolin(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val)

    curr_row += 1
    vat = int(round(final_budget_val * 0.05))
    grand_total = final_budget_val + vat
    
    footer_data = [
        ("製作", prod_cost),
        ("5% VAT", vat),
        ("Grand Total", grand_total)
    ]
    
    label_col = 6 if format_type == "Dongwu" else 36
    val_col = 7 if format_type == "Dongwu" else 39
    
    for label, val in footer_data:
        ws.row_dimensions[curr_row].height = FOOTER_ROW_HEIGHT
        ws.cell(curr_row, label_col).value = label
        ws.cell(curr_row, label_col).alignment = Alignment(horizontal='right', vertical='center')
        ws.cell(curr_row, label_col).font = Font(name=FONT_MAIN, size=12)
        ws.cell(curr_row, val_col).value = val
        ws.cell(curr_row, val_col).number_format = "#,##0"
        ws.cell(curr_row, val_col).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(curr_row, val_col).font = Font(name=FONT_MAIN, size=12)
        
        if label == "Grand Total":
            grand_fill = PatternFill(start_color="FFC107", end_color="FFC107", fill_type="solid")
            for c in range(1, 40):
                ws.cell(curr_row, c).fill = grand_fill
                ws.cell(curr_row, c).border = Border(top=Side(style='medium'), bottom=Side(style='medium'))
        curr_row += 1

    curr_row += 1
    ws.cell(curr_row, 1).value = "Remarks："
    ws.cell(curr_row, 1).font = Font(name=FONT_MAIN, size=16, bold=True, underline="single", color="FF0000")
    curr_row += 1
    for rm in remarks_list:
        ws.cell(curr_row, 1).value = rm
        ws.cell(curr_row, 1).font = Font(name=FONT_MAIN, size=14)
        curr_row += 1

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 7. UI Main
# =========================================================
with st.sidebar:
    st.header("🕵️ 主管登入")
    if not st.session_state.is_supervisor:
        pwd = st.text_input("輸入密碼", type="password", key="pwd_input")
        if st.button("登入"):
            if pwd == "1234":
                st.session_state.is_supervisor = True
                st.rerun()
            else:
                st.error("密碼錯誤")
    else:
        st.success("✅ 目前狀態：主管模式")
        if st.button("登出"):
            st.session_state.is_supervisor = False
            st.rerun()

st.title("📺 媒體 Cue 表生成器 (v81.1)")

st.markdown("### 1. 選擇格式")
format_type = st.radio("", ["Dongwu", "Shenghuo", "Bolin"], horizontal=True)

st.markdown("### 2. 基本資料設定")
c1, c2, c3, c4 = st.columns(4)
with c1: client_name = st.text_input("客戶名稱", "萬國通路")
with c2: product_name = st.text_input("產品名稱", "統一布丁")
with c3: total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)
with c4: prod_cost_input = st.number_input("製作費 (未稅)", value=0, step=1000)

final_budget_val = total_budget_input
if st.session_state.is_supervisor:
    st.markdown("---")
    col_sup1, col_sup2 = st.columns([1, 2])
    with col_sup1:
        st.error("🔒 [主管] 專案優惠價覆寫")
    with col_sup2:
        override_val = st.number_input("輸入最終成交價 (此數值將取代自動計算的 Total)", value=total_budget_input)
        if override_val != total_budget_input:
            final_budget_val = override_val
            st.caption(f"⚠️ 注意：報表將使用 ${final_budget_val:,} 進行結算")
    st.markdown("---")

c5, c6 = st.columns(2)
with c5: start_date = st.date_input("開始日", datetime(2026, 1, 1))
with c6: end_date = st.date_input("結束日", datetime(2026, 1, 31))
days_count = (end_date - start_date).days + 1
st.info(f"📅 走期共 **{days_count}** 天")

with st.expander("📝 備註欄位設定 (Remarks)", expanded=False):
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

st.write("請勾選要投放的媒體：")
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
        effective_is_nat = is_nat
        if not is_nat and len(regs) == 6:
            effective_is_nat = True
            regs = ["全省"]
            st.info("✅ 已選滿6區，自動轉為全省聯播計價")
        secs = st.multiselect("秒數", DURATIONS, [20], key="rad_sec")
        st.slider("預算 %", 0, 100, key="rad_share", on_change=on_slider_change, args=("rad_share",))
        sec_shares = {}
        if len(secs) > 1:
            st.caption("分配秒數佔比")
            rem = 100
            sorted_secs = sorted(secs)
            for i, s in enumerate(sorted_secs):
                if i < len(sorted_secs) - 1:
                    v = st.slider(f"{s}秒 %", 0, rem, int(rem/2), key=f"rs_{s}")
                    sec_shares[s] = v; rem -= v
                else:
                    sec_shares[s] = rem
                    st.markdown(f"🔹 **{s}秒**: {rem}% (自動計算)")
        elif secs: sec_shares[secs[0]] = 100
        config["全家廣播"] = {"is_national": effective_is_nat, "regions": regs, "sec_shares": sec_shares, "share": st.session_state.rad_share}

if is_fv:
    with m2:
        st.markdown("#### 📺 新鮮視")
        is_nat = st.checkbox("全省聯播", False, key="fv_nat")
        regs = ["全省"] if is_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區"], key="fv_reg")
        effective_is_nat = is_nat
        if not is_nat and len(regs) == 6:
            effective_is_nat = True
            regs = ["全省"]
            st.info("✅ 已選滿6區，自動轉為全省聯播計價")
        secs = st.multiselect("秒數", DURATIONS, [10], key="fv_sec")
        st.slider("預算 %", 0, 100, key="fv_share", on_change=on_slider_change, args=("fv_share",))
        sec_shares = {}
        if len(secs) > 1:
            st.caption("分配秒數佔比")
            rem = 100
            sorted_secs = sorted(secs)
            for i, s in enumerate(sorted_secs):
                if i < len(sorted_secs) - 1:
                    v = st.slider(f"{s}秒 %", 0, rem, int(rem/2), key=f"fs_{s}")
                    sec_shares[s] = v; rem -= v
                else:
                    sec_shares[s] = rem
                    st.markdown(f"🔹 **{s}秒**: {rem}% (自動計算)")
        elif secs: sec_shares[secs[0]] = 100
        config["新鮮視"] = {"is_national": effective_is_nat, "regions": regs, "sec_shares": sec_shares, "share": st.session_state.fv_share}

if is_cf:
    with m3:
        st.markdown("#### 🛒 家樂福")
        secs = st.multiselect("秒數", DURATIONS, [20], key="cf_sec")
        st.slider("預算 %", 0, 100, key="cf_share", on_change=on_slider_change, args=("cf_share",))
        sec_shares = {}
        if len(secs) > 1:
            st.caption("分配秒數佔比")
            rem = 100
            sorted_secs = sorted(secs)
            for i, s in enumerate(sorted_secs):
                if i < len(sorted_secs) - 1:
                    v = st.slider(f"{s}秒 %", 0, rem, int(rem/2), key=f"cs_{s}")
                    sec_shares[s] = v; rem -= v
                else:
                    sec_shares[s] = rem
                    st.markdown(f"🔹 **{s}秒**: {rem}% (自動計算)")
        elif secs: sec_shares[secs[0]] = 100
        config["家樂福"] = {"regions": ["全省"], "sec_shares": sec_shares, "share": st.session_state.cf_share}

if config:
    rows, total_list_accum, logs = calculate_plan_data(config, total_budget_input, days_count)
    
    prod_cost = prod_cost_input 
    vat = int(round(final_budget_val * 0.05))
    grand_total = final_budget_val + vat
    
    p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}"
    rem = get_remarks_text(sign_deadline, billing_month, payment_date)

    # Simplified HTML preview generator for stability
    html_preview = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem, total_list_accum, grand_total, final_budget_val, prod_cost)
    st.components.v1.html(html_preview, height=700, scrolling=True)

    with st.expander("💡 系統運算邏輯說明 (Debug Panel)", expanded=False):
        for log in logs:
            st.markdown(f"### {log['Media']}")
            st.markdown(f"- **預算**: {log['Budget']}")
            st.markdown(f"- **狀態**: {log['Status']}")
            if 'Details' in log:
                for detail in log['Details']:
                    st.info(detail)
            st.divider()

    col_dl1, col_dl2 = st.columns(2)
    with col_dl2:
        try:
            xlsx_temp = generate_excel_from_scratch(format_type, start_date, end_date, client_name, p_str, rows, rem, final_budget_val, prod_cost)
            pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx_temp)
            if pdf_bytes:
                st.download_button(f"📥 下載 PDF ({method})", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf", key="pdf_dl")
            else:
                st.warning(f"本地轉檔失敗，使用網頁版 PDF")
                pdf_bytes, err = html_to_pdf_weasyprint(html_preview)
                if pdf_bytes: st.download_button("📥 下載 PDF (Web)", pdf_bytes, f"Cue_{safe_filename(client_name)}.pdf", key="pdf_dl_web")
        except: pass

    with col_dl1:
        if st.session_state.is_supervisor:
            if rows:
                try:
                    xlsx = generate_excel_from_scratch(format_type, start_date, end_date, client_name, p_str, rows, rem, final_budget_val, prod_cost)
                    st.download_button("📥 下載 Excel (主管權限)", xlsx, f"Cue_{safe_filename(client_name)}.xlsx", key="xlsx_dl")
                except Exception as e:
                    st.error(f"Excel Error: {e}")
        else:
            st.info("🔒 Excel 下載功能僅限主管使用 (請從左側登入)")
