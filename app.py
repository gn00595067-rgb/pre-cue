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
# 1. 頁面設定 (必須在所有 st 指令之前)
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v92.2")

# =========================================================
# 2. 初始化 Session State
# =========================================================
if "is_supervisor" not in st.session_state:
    st.session_state.is_supervisor = False

# =========================================================
# 3. 基礎工具
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
# 4. PDF 策略
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
# 5. 核心資料設定 (雲端 Google Sheet 版)
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
# 6. 核心計算函式
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
# 7. OpenPyXL 規格重建引擎
# =========================================================
FONT_MAIN = "微軟正黑體"
BS_THIN = 'thin'
BS_MEDIUM = 'medium'
BS_THICK = 'medium' # Reusing medium as thick per user request
BS_HAIR = 'hair'

def set_border(cell, top=None, bottom=None, left=None, right=None):
    cur = cell.border
    t = top if top is not None else (cur.top.style if cur.top else None)
    b = bottom if bottom is not None else (cur.bottom.style if cur.bottom else None)
    l = left if left is not None else (cur.left.style if cur.left else None)
    r = right if right is not None else (cur.right.style if cur.right else None)
    
    cell.border = Border(
        top=Side(style=t) if t else Side(),
        bottom=Side(style=b) if b else Side(),
        left=Side(style=l) if l else Side(),
        right=Side(style=r) if r else Side()
    )

def style_range(ws, cell_range, border=Border(), fill=None, font=None, alignment=None):
    rows = list(ws[cell_range])
    for row in rows:
        for cell in row:
            if border: cell.border = border
            if fill: cell.fill = fill
            if font: cell.font = font
            if alignment: cell.alignment = alignment

def draw_outer_border(ws, min_r, max_r, min_c, max_c):
    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            cell = ws.cell(r, c)
            set_border(cell, 
                       top=BS_MEDIUM if r == min_r else None,
                       bottom=BS_MEDIUM if r == max_r else None,
                       left=BS_MEDIUM if c == min_c else None,
                       right=BS_MEDIUM if c == max_c else None)

# ----------------- Dongwu Engine -----------------
def render_dongwu(ws, start_dt, end_dt, client_name, product_display_str, rows, remarks_list, final_budget_val):
    COL_WIDTHS = {'A': 19.6, 'B': 22.8, 'C': 14.6, 'D': 20.0, 'E': 13.0, 'F': 19.6, 'G': 17.9}
    ROW_HEIGHTS = {1: 61.0, 2: 29.0, 3: 40.0, 4: 40.0, 5: 40.0, 6: 40.0, 7: 40.0, 8: 40.0}
    for k, v in COL_WIDTHS.items(): ws.column_dimensions[k].width = v
    for i in range(8, 40): ws.column_dimensions[get_column_letter(i)].width = 8.5
    ws.column_dimensions['AM'].width = 13.0
    for r, h in ROW_HEIGHTS.items(): ws.row_dimensions[r].height = h
    ws['A1'] = "Media Schedule"; ws.merge_cells("A1:AM1")
    style_range(ws, "A1:AM1", font=Font(name=FONT_MAIN, size=48, bold=True), alignment=Alignment(horizontal='center', vertical='center'))
    for c in range(1, 40): ws.cell(3, c).border = Border(top=Side(style=BS_MEDIUM))
    info_map = {"A3": ("客戶名稱：", client_name), "A4": ("Product：", product_display_str), "A5": ("Period :", f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"), "A6": ("Medium :", "全家廣播/新鮮視/家樂福")}
    for addr, (lbl, val) in info_map.items():
        ws[addr] = lbl; ws[addr].font = Font(name=FONT_MAIN, size=14, bold=True); ws[addr].alignment = Alignment(vertical='center')
        val_cell = ws.cell(ws[addr].row, 2); val_cell.value = val; val_cell.font = Font(name=FONT_MAIN, size=14, bold=True); val_cell.alignment = Alignment(vertical='center')
    ws['H6'] = f"{start_dt.month}月"; ws['H6'].font = Font(name=FONT_MAIN, size=16, bold=True); ws['H6'].alignment = Alignment(horizontal='center', vertical='center')
    headers = [("A","Station"), ("B","Location"), ("C","Program"), ("D","Day-part"), ("E","Size"), ("F","rate\n(Net)"), ("G","Package-cost\n(Net)")]
    for col, txt in headers:
        ws[f"{col}7"] = txt; ws.merge_cells(f"{col}7:{col}8")
        style_range(ws, f"{col}7:{col}8", font=Font(name=FONT_MAIN, size=14), alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
        set_border(ws.cell(7, column_index_from_string(col)), top=BS_MEDIUM, bottom=BS_MEDIUM, left=BS_THIN, right=BS_THIN)
    curr = start_dt; eff_days = (end_dt - start_dt).days + 1
    for i in range(31):
        col_idx = 8 + i; d_cell = ws.cell(7, col_idx); w_cell = ws.cell(8, col_idx)
        if i < eff_days:
            d_cell.value = curr; d_cell.number_format = 'm/d'; w_cell.value = ["一","二","三","四","五","六","日"][curr.weekday()]
            if curr.weekday() >= 5: w_cell.fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
            curr += timedelta(days=1)
        d_cell.font = Font(name=FONT_MAIN, size=12); w_cell.font = Font(name=FONT_MAIN, size=12)
        d_cell.alignment = w_cell.alignment = Alignment(horizontal='center', vertical='center')
        set_border(d_cell, top=BS_MEDIUM, bottom=BS_THIN, left=BS_THIN, right=BS_THIN)
        set_border(w_cell, top=BS_THIN, bottom=BS_MEDIUM, left=BS_THIN, right=BS_THIN)
    ws['AM7'] = "檔次"; ws.merge_cells("AM7:AM8")
    style_range(ws, "AM7:AM8", font=Font(name=FONT_MAIN, size=14), alignment=Alignment(horizontal='center', vertical='center'))
    set_border(ws['AM7'], top=BS_MEDIUM, bottom=BS_MEDIUM, left=BS_THIN, right=BS_THIN)
    return render_data_rows(ws, rows, 9, final_budget_val, eff_days, "Dongwu", product_display_str)

# ----------------- Shenghuo Engine (White BG + Infinite) -----------------
def render_shenghuo(ws, start_dt, end_dt, client_name, product_name_raw, rows, remarks_list, final_budget_val, prod_cost):
    days_n = (end_dt - start_dt).days + 1
    
    ws.column_dimensions['A'].width = 22.5
    ws.column_dimensions['B'].width = 24.5
    ws.column_dimensions['C'].width = 13.8
    ws.column_dimensions['D'].width = 19.4
    ws.column_dimensions['E'].width = 13.0
    for i in range(days_n): ws.column_dimensions[get_column_letter(6 + i)].width = 13.0
    end_c_start = 6 + days_n
    ws.column_dimensions[get_column_letter(end_c_start)].width = 13.0   
    ws.column_dimensions[get_column_letter(end_c_start+1)].width = 59.0 
    ws.column_dimensions[get_column_letter(end_c_start+2)].width = 13.2 
    total_cols = 5 + days_n + 3
    ROW_H_MAP = {1:46, 2:46, 3:46, 4:46.5, 5:40, 6:40, 7:40, 8:40}
    for r, h in ROW_H_MAP.items(): ws.row_dimensions[r].height = h
    
    ws['A3'] = "聲活數位科技股份有限公司 統編 28710100"; ws['A3'].font = Font(name=FONT_MAIN, size=20); ws['A3'].alignment = Alignment(vertical='center')
    ws['A4'] = "蔡伊閔"; ws['A4'].font = Font(name=FONT_MAIN, size=16); ws['A4'].alignment = Alignment(vertical='center')

    for r in [5, 6]:
        for c in range(1, total_cols + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name=FONT_MAIN, size=14, bold=True)
            set_border(cell, top=BS_MEDIUM if r==5 else None, bottom=BS_MEDIUM if r==6 else None, left=BS_MEDIUM if c==1 else None, right=BS_MEDIUM if c==total_cols else None)

    ws['A5'] = "客戶名稱："; ws['B5'] = client_name
    ws['F5'] = "廣告規格："; 
    unique_secs = sorted(list(set([r['seconds'] for r in rows])))
    ws['H5'] = " ".join([f"{s}秒廣告" for s in unique_secs])

    date_range_col = total_cols - 1
    ws.cell(5, date_range_col).value = f"執行期間：: {start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"
    ws.cell(5, date_range_col).alignment = Alignment(horizontal='right', vertical='center')

    ws['A6'] = "廣告名稱："; ws['B6'] = product_name_raw
    ws.cell(6, 6).value = f"{start_dt.month}月"
    for i in range(days_n):
        d = start_dt + timedelta(days=i)
        if d.month != start_dt.month and d.day == 1: ws.cell(6, 6+i).value = f"{d.month}月"

    headers = ["頻道", "播出地區", "播出店數", "播出時間", "秒數\n規格"]
    for i, h in enumerate(headers):
        ws.merge_cells(start_row=7, start_column=i+1, end_row=8, end_column=i+1)
        cell = ws.cell(7, i+1); cell.value = h
        style_range(ws, f"{get_column_letter(i+1)}7:{get_column_letter(i+1)}8", font=Font(name=FONT_MAIN, size=14, bold=True), alignment=Alignment(horizontal='center', vertical='center', wrap_text=True))
        set_border(cell, top=BS_MEDIUM, bottom=BS_HAIR, left=BS_HAIR, right=BS_HAIR)
    
    set_border(ws.cell(7,1), top=BS_MEDIUM, left=BS_MEDIUM, right=BS_HAIR)
    set_border(ws.cell(8,1), bottom=BS_HAIR, left=BS_MEDIUM, right=BS_HAIR)

    curr = start_dt
    for i in range(days_n):
        c = 6 + i
        cell7 = ws.cell(7, c); cell7.value = curr; cell7.number_format = 'd'
        cell7.font = Font(name=FONT_MAIN, size=14, bold=True); cell7.alignment = Alignment(horizontal='center', vertical='center')
        set_border(cell7, top=BS_MEDIUM, bottom=BS_HAIR, left=BS_HAIR, right=BS_HAIR)
        
        cell8 = ws.cell(8, c); cell8.value = f'=MID("日一二三四五六",WEEKDAY({get_column_letter(c)}7,1),1)'
        cell8.font = Font(name=FONT_MAIN, size=14, bold=True); cell8.alignment = Alignment(horizontal='center', vertical='center')
        set_border(cell8, top=BS_HAIR, bottom=BS_HAIR, left=BS_HAIR, right=BS_HAIR)
        
        if curr.weekday() >= 5: cell8.fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
        curr += timedelta(days=1)

    end_headers = ["檔次", "定價", "專案價"]
    for i, h in enumerate(end_headers):
        c = end_c_start + i
        ws.merge_cells(start_row=7, start_column=c, end_row=8, end_column=c)
        ws.cell(7, c).value = h
        style_range(ws, f"{get_column_letter(c)}7:{get_column_letter(c)}8", font=Font(name=FONT_MAIN, size=14, bold=True), alignment=Alignment(horizontal='center', vertical='center'))
        set_border(ws.cell(7, c), top=BS_MEDIUM, bottom=BS_HAIR, left=BS_HAIR, right=BS_HAIR)
    
    set_border(ws.cell(7, total_cols), top=BS_MEDIUM, right=BS_MEDIUM, left=BS_HAIR)
    set_border(ws.cell(8, total_cols), bottom=BS_HAIR, right=BS_MEDIUM, left=BS_HAIR)

    return render_data_rows(ws, rows, 9, final_budget_val, days_n, "Shenghuo", product_name_raw)

# ----------------- Bolin Engine (Infinite) -----------------
def render_bolin(ws, start_dt, end_dt, client_name, product_name_raw, rows, remarks_list, final_budget_val, prod_cost):
    days_n = (end_dt - start_dt).days + 1
    total_cols = 1 + 5 + days_n + 3 
    
    ws.column_dimensions['A'].width = 1.76; ws.column_dimensions['B'].width = 20; ws.column_dimensions['C'].width = 22; ws.column_dimensions['D'].width = 10; ws.column_dimensions['E'].width = 15; ws.column_dimensions['F'].width = 10
    for i in range(days_n): ws.column_dimensions[get_column_letter(7 + i)].width = 5
    end_c_start = 7 + days_n
    ws.column_dimensions[get_column_letter(end_c_start)].width = 8; ws.column_dimensions[get_column_letter(end_c_start+1)].width = 12; ws.column_dimensions[get_column_letter(end_c_start+2)].width = 12

    ROW_H_MAP = {1:15, 2:25, 3:25, 4:25, 5:25, 6:25, 7:35}
    for r, h in ROW_H_MAP.items(): ws.row_dimensions[r].height = h
    
    ws['B2'] = "TO："; ws['B2'].font = Font(name=FONT_MAIN, size=13, bold=True); ws['B2'].alignment = Alignment(horizontal='right')
    ws['C2'] = client_name; ws['C2'].font = Font(name=FONT_MAIN, size=13)
    ws['B3'] = "FROM："; ws['B3'].font = Font(name=FONT_MAIN, size=13, bold=True); ws['B3'].alignment = Alignment(horizontal='right')
    ws['C3'] = "鉑霖行動行銷 許雅婷 TINA"; ws['C3'].font = Font(name=FONT_MAIN, size=13)
    ws['B4'] = "客戶名稱："; ws['B4'].font = Font(name=FONT_MAIN, size=13, bold=True); ws['B4'].alignment = Alignment(horizontal='right')
    ws['C4'] = client_name; ws['C4'].font = Font(name=FONT_MAIN, size=13)
    ws['B5'] = "廣告名稱："; ws['B5'].font = Font(name=FONT_MAIN, size=13, bold=True); ws['B5'].alignment = Alignment(horizontal='right')
    ws['C5'] = product_name_raw; ws['C5'].font = Font(name=FONT_MAIN, size=13)

    ws['G4'] = "廣告規格："; ws['G4'].font = Font(name=FONT_MAIN, size=13, bold=True)
    unique_secs = sorted(list(set([r['seconds'] for r in rows])))
    ws['H4'] = " ".join([f"{s}秒廣告" for s in unique_secs]); ws['H4'].font = Font(name=FONT_MAIN, size=13)
    
    date_lbl_col = total_cols - 2; date_val_col = total_cols - 1
    ws.cell(4, date_lbl_col).value = "執行期間："; ws.cell(4, date_lbl_col).font = Font(name=FONT_MAIN, size=13, bold=True)
    ws.cell(4, date_val_col).value = f"{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')}"; ws.cell(4, date_val_col).font = Font(name=FONT_MAIN, size=13)

    header_fill = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")
    headers = ["頻道", "播出地區", "播出店數", "播出時間", "規格"]
    for i, h in enumerate(headers):
        c = 2 + i
        cell = ws.cell(7, c); cell.value = h
        cell.fill = header_fill; cell.font = Font(name=FONT_MAIN, size=12, bold=True); cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        set_border(cell, top=BS_MEDIUM, bottom=BS_MEDIUM, left=BS_THIN, right=BS_THIN)
        if c==2: set_border(cell, left=BS_MEDIUM)

    curr = start_dt
    for i in range(days_n):
        c = 7 + i
        cell = ws.cell(7, c); cell.value = curr; cell.number_format = 'm/d'
        cell.fill = header_fill; cell.font = Font(name=FONT_MAIN, size=10, bold=True); cell.alignment = Alignment(horizontal='center', vertical='center')
        set_border(cell, top=BS_MEDIUM, bottom=BS_MEDIUM, left=BS_THIN, right=BS_THIN)
        curr += timedelta(days=1)

    end_h = ["總檔次", "單價", "金額"]
    for i, h in enumerate(end_h):
        c = end_c_start + i
        cell = ws.cell(7, c); cell.value = h
        cell.fill = header_fill; cell.font = Font(name=FONT_MAIN, size=12, bold=True); cell.alignment = Alignment(horizontal='center', vertical='center')
        set_border(cell, top=BS_MEDIUM, bottom=BS_MEDIUM, left=BS_THIN, right=BS_THIN)
        if i==2: set_border(cell, right=BS_MEDIUM)

    return render_data_rows(ws, rows, 8, final_budget_val, days_n, "Bolin", product_name_raw)

# Common Data Renderer
def render_data_rows(ws, rows, start_row, final_budget_val, eff_days, mode, product_name_raw):
    curr_row = start_row
    font_content = Font(name=FONT_MAIN, size=14 if mode in ["Dongwu","Shenghuo"] else 12)
    row_height = 40 if mode in ["Dongwu","Shenghuo"] else 25

    grouped_data = {
        "全家廣播": sorted([r for r in rows if r["media"] == "全家廣播"], key=lambda x: x["seconds"]),
        "新鮮視": sorted([r for r in rows if r["media"] == "新鮮視"], key=lambda x: x["seconds"]),
        "家樂福": sorted([r for r in rows if r["media"] == "家樂福"], key=lambda x: x["seconds"]),
    }

    max_c = 39 if mode == "Dongwu" else 5 + eff_days + 3
    if mode == "Bolin": max_c = 1 + 5 + eff_days + 3

    for m_key, data in grouped_data.items():
        if not data: continue
        start_merge_row = curr_row
        
        start_c = 1 if mode != "Bolin" else 2
        for c in range(start_c, max_c + 1):
            cell = ws.cell(curr_row, c)
            l = BS_MEDIUM if c==start_c else BS_THIN if mode != "Shenghuo" else BS_HAIR
            r = BS_MEDIUM if c==max_c else BS_THIN if mode != "Shenghuo" else BS_HAIR
            set_border(cell, top=BS_MEDIUM, left=l, right=r, bottom=BS_THIN if mode!="Shenghuo" else BS_HAIR)

        display_name = f"全家便利商店\n{m_key if m_key!='家樂福' else ''}廣告"
        if m_key == "家樂福": display_name = "家樂福"
        elif m_key == "全家廣播": display_name = "全家便利商店\n通路廣播廣告"
        elif m_key == "新鮮視": display_name = "全家便利商店\n新鮮視廣告"

        for idx, r_data in enumerate(data):
            ws.row_dimensions[curr_row].height = row_height
            
            sec_txt = f"{r_data['seconds']}秒"
            store_txt = str(int(r_data.get("program_num", 0)))
            if mode == "Shenghuo":
                if m_key == "新鮮視":
                    sec_txt = f"{r_data['seconds']}秒\n影片/影像 1920x1080 (mp4)"; store_txt = f"{store_txt}面"
                elif m_key == "全家廣播":
                    sec_txt = f"{r_data['seconds']}秒廣告"; store_txt = f"{store_txt}店"
                else: sec_txt = f"{r_data['seconds']}秒廣告"
            
            base_c = 1 if mode != "Bolin" else 2
            ws.cell(curr_row, base_c).value = display_name
            ws.cell(curr_row, base_c+1).value = r_data["region"]
            ws.cell(curr_row, base_c+2).value = store_txt
            ws.cell(curr_row, base_c+3).value = r_data["daypart"]
            ws.cell(curr_row, base_c+4).value = sec_txt
            
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

            ws.cell(curr_row, total_col).value = row_sum

            for c in range(start_c, max_c + 1):
                cell = ws.cell(curr_row, c)
                cell.font = font_content
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                
                l_style = BS_THIN; r_style = BS_THIN; b_style = BS_THIN
                if mode == "Shenghuo":
                    l_style = BS_HAIR; r_style = BS_HAIR; b_style = BS_HAIR
                    if c==start_c: l_style = BS_MEDIUM
                    if c==max_c: r_style = BS_MEDIUM
                elif mode == "Bolin":
                    if c==start_c: l_style = BS_MEDIUM
                    if c==max_c: r_style = BS_MEDIUM
                else: # Dongwu
                    if c==start_c: l_style = BS_MEDIUM
                    if c==max_c: r_style = BS_MEDIUM

                set_border(cell, left=l_style, right=r_style, bottom=b_style)
                if isinstance(cell.value, (int, float)): cell.number_format = "#,##0_);[Red](#,##0)" if mode=="Shenghuo" else "#,##0"
            curr_row += 1

        if curr_row > start_merge_row:
            ws.merge_cells(start_row=start_merge_row, start_column=start_c, end_row=curr_row-1, end_column=start_c)
        
        if data[0].get("is_pkg_member"):
            if mode == "Dongwu": ws.merge_cells(start_row=start_merge_row, start_column=7, end_row=curr_row-1, end_column=7)
            elif mode == "Shenghuo": 
                p_c = 5+eff_days+3
                ws.merge_cells(start_row=start_merge_row, start_column=p_c, end_row=curr_row-1, end_column=p_c)
            else: # Bolin
                ws.merge_cells(start_row=start_merge_row, start_column=40, end_row=curr_row-1, end_column=40)
        
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

        for c in range(start_c, max_c + 1):
            cell = ws.cell(curr_row-1, c)
            set_border(cell, bottom=BS_MEDIUM)

    # Total Row
    ws.row_dimensions[curr_row].height = 40 if mode=="Shenghuo" else 30
    label_col = 6 if mode == "Dongwu" else 5 if mode == "Shenghuo" else 1+5+eff_days+2
    total_val_col = 7 if mode == "Dongwu" else 5+eff_days+3 if mode == "Shenghuo" else 1+5+eff_days+3
    
    ws.cell(curr_row, label_col).value = "Total"; ws.cell(curr_row, label_col).alignment = Alignment(horizontal='right', vertical='center')
    ws.cell(curr_row, label_col).font = Font(name=FONT_MAIN, size=14 if mode!="Bolin" else 12, bold=True)
    ws.cell(curr_row, total_val_col).value = final_budget_val; ws.cell(curr_row, total_val_col).number_format = "#,##0"
    ws.cell(curr_row, total_val_col).font = Font(name=FONT_MAIN, size=14 if mode!="Bolin" else 12, bold=True); ws.cell(curr_row, total_val_col).alignment = Alignment(horizontal='center', vertical='center')

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
    
    start_c = 1 if mode != "Bolin" else 2
    for c in range(start_c, max_c + 1):
        cell = ws.cell(curr_row, c)
        l = BS_MEDIUM if c==start_c else BS_THIN if mode!="Shenghuo" else BS_HAIR
        r = BS_MEDIUM if c==max_c else BS_THIN if mode!="Shenghuo" else BS_HAIR
        set_border(cell, top=BS_MEDIUM, bottom=BS_MEDIUM, left=l, right=r)
        if mode == "Dongwu" and c==1: set_border(cell, left=BS_MEDIUM)
    
    return curr_row

def generate_excel_from_scratch(format_type, start_dt, end_dt, client_name, product_name, rows, remarks_list, final_budget_val, prod_cost):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "工作表1"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE; ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.fitToPage = True; ws.page_setup.fitToWidth = 1
    
    unique_secs = sorted(list(set([r['seconds'] for r in rows])))
    product_display_str_dongwu = f"{'、'.join([f'{s}秒' for s in unique_secs])} {product_name}"
    
    if format_type == "Dongwu": curr_row = render_dongwu(ws, start_dt, end_dt, client_name, product_display_str_dongwu, rows, remarks_list, final_budget_val)
    elif format_type == "Shenghuo": curr_row = render_shenghuo(ws, start_dt, end_dt, client_name, product_name, rows, remarks_list, final_budget_val, prod_cost)
    else: curr_row = render_bolin(ws, start_dt, end_dt, client_name, product_name, rows, remarks_list, final_budget_val, prod_cost)

    if format_type == "Dongwu":
        curr_row += 1
        vat = int(round(final_budget_val * 0.05)); grand_total = final_budget_val + vat
        footer_data = [("製作", prod_cost), ("5% VAT", vat), ("Grand Total", grand_total)]
        label_col = 6; val_col = 7
        for label, val in footer_data:
            ws.row_dimensions[curr_row].height = 30
            ws.cell(curr_row, label_col).value = label; ws.cell(curr_row, label_col).alignment = Alignment(horizontal='right', vertical='center'); ws.cell(curr_row, label_col).font = Font(name=FONT_MAIN, size=14)
            ws.cell(curr_row, val_col).value = val; ws.cell(curr_row, val_col).number_format = "#,##0"; ws.cell(curr_row, val_col).alignment = Alignment(horizontal='center', vertical='center'); ws.cell(curr_row, val_col).font = Font(name=FONT_MAIN, size=14)
            set_border(ws.cell(curr_row, label_col), left=BS_MEDIUM, top=BS_THIN, bottom=BS_THIN, right=BS_THIN)
            set_border(ws.cell(curr_row, val_col), right=BS_MEDIUM, top=BS_THIN, bottom=BS_THIN, left=BS_THIN)
            if label == "Grand Total":
                for c in range(1, 40): ws.cell(curr_row, c).fill = PatternFill(start_color="FFC107", end_color="FFC107", fill_type="solid"); set_border(ws.cell(curr_row, c), top=BS_MEDIUM, bottom=BS_MEDIUM)
            curr_row += 1
        draw_outer_border(ws, 7, curr_row-1, 1, 39)

    if format_type == "Dongwu":
        curr_row += 1
        ws.cell(curr_row, 1).value = "Remarks："
        ws.cell(curr_row, 1).font = Font(name=FONT_MAIN, size=16, bold=True, underline="single", color="000000")
        for c in range(1, 40): set_border(ws.cell(curr_row, c), top=None)
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
        ws.cell(curr_row, 9).value = "Remarks："
        ws.cell(curr_row, 9).font = Font(name=FONT_MAIN, size=16, bold=True, underline="single")
        curr_row += 1
        for rm in remarks_list:
            ws.cell(curr_row, 9).value = rm
            ws.cell(curr_row, 9).font = Font(name=FONT_MAIN, size=16, bold=True)
            curr_row += 1

    out = io.BytesIO(); wb.save(out); return out.getvalue()
