import streamlit as st
import traceback
import time
import gc
from itertools import groupby
import requests

# =========================================================
# 1. 頁面設定
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v111.25 (Bolin Fix)")

import pandas as pd
import math
import io
import os
import shutil
import tempfile
import subprocess
import re
from datetime import timedelta, datetime, date
from copy import copy

# =========================================================
# 2. Session State 初始化
# =========================================================
if "is_supervisor" not in st.session_state: st.session_state.is_supervisor = False
if "rad_share" not in st.session_state: st.session_state.rad_share = 100
if "fv_share" not in st.session_state: st.session_state.fv_share = 0
if "cf_share" not in st.session_state: st.session_state.cf_share = 0
if "cb_rad" not in st.session_state: st.session_state.cb_rad = True
if "cb_fv" not in st.session_state: st.session_state.cb_fv = False
if "cb_cf" not in st.session_state: st.session_state.cb_cf = False

# =========================================================
# 3. 全域常數
# =========================================================
GSHEET_SHARE_URL = "https://docs.google.com/spreadsheets/d/1bzmG-N8XFsj8m3LUPqA8K70AcIqaK4Qhq1VPWcK0w_s/edit?usp=sharing"
BOLIN_LOGO_URL = "https://docs.google.com/drawings/d/17Uqgp-7LJJj9E4bV7Azo7TwXESPKTTIsmTbf-9tU9eE/export/png"

FONT_MAIN = "微軟正黑體"
BS_THIN = 'thin'; BS_MEDIUM = 'medium'; BS_HAIR = 'hair'
FMT_MONEY = '"$"#,##0_);[Red]("$"#,##0)'; FMT_NUMBER = '#,##0'
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]
REGION_DISPLAY_MAP = {"北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投", "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東", "全省量販": "全省量販", "全省超市": "全省超市"}

# =========================================================
# 4. 基礎工具函式
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

def region_display(region):
    return REGION_DISPLAY_MAP.get(region, region)

def get_sec_factor(media_type, seconds, sec_factors):
    factors = sec_factors.get(media_type)
    if not factors:
        if media_type == "新鮮視": factors = sec_factors.get("全家新鮮視")
        elif media_type == "全家廣播": factors = sec_factors.get("全家廣播")
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
    d_str = sign_deadline.strftime("%Y/%m/%d (%a)") if sign_deadline else "____/__/__ (__)"
    p_str = payment_date.strftime("%Y/%m/%d") if payment_date else "____/__/__"
    return [
        f"1.請於 {d_str} 11:30前 回簽及進單，方可順利上檔。",
        "2.以上節目名稱如有異動，以上檔時節目名稱為主，如遇電台時段滿檔，上檔時間挪後或更換至同級時段。",
        "3.通路店鋪數與開機率開機率至少七成(以上)。每日因加盟數調整，或遇店舖年度季度改裝、設備維護升級及保修等狀況，會有一定幅度增減。",
        "4.託播方需於上檔前 5 個工作天，提供廣告帶(mp3)、影片/影像 1920x1080 (mp4)。",
        f"5.雙方同意費用請款月份 : {billing_month}，如有修正必要，將另行E-Mail告知，並視為正式合約之一部分。",
        f"6.付款兌現日期：{p_str}"
    ]

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

@st.cache_data(show_spinner="正在下載 Logo...", ttl=3600)
def get_cloud_logo_bytes():
    try:
        response = requests.get(BOLIN_LOGO_URL, timeout=10)
        if response.status_code == 200:
            return response.content
        return None
    except:
        return None

@st.cache_data(show_spinner="正在生成 PDF (LibreOffice)...", ttl=3600)
def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    soffice = find_soffice_path()
    if not soffice: 
        return None, "Fail", "伺服器未安裝 LibreOffice"
    try:
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, "cue.xlsx")
            with open(xlsx_path, "wb") as f: f.write(xlsx_bytes)
            subprocess.run([soffice, "--headless", "--nologo", "--convert-to", "pdf:calc_pdf_Export", "--outdir", tmp, xlsx_path], capture_output=True, timeout=60)
            pdf_path = os.path.join(tmp, "cue.pdf")
            if not os.path.exists(pdf_path):
                for fn in os.listdir(tmp):
                    if fn.endswith(".pdf"): pdf_path = os.path.join(tmp, fn); break
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f: return f.read(), "LibreOffice", ""
            return None, "Fail", "LibreOffice 未產出檔案"
    except subprocess.TimeoutExpired:
        return None, "Fail", "轉檔逾時"
    except Exception as e: return None, "Fail", str(e)
    finally:
        gc.collect()

def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks, total_list, grand_total, budget, prod):
    eff_days = days_cnt
    header_cls = "bg-dw-head" if format_type == "Dongwu" else "bg-sh-head"
    if format_type == "Bolin": header_cls = "bg-bolin-head"
    date_th1 = ""; date_th2 = ""; curr = start_dt; weekdays = ["一", "二", "三", "四", "五", "六", "日"]
    for i in range(eff_days):
        wd = curr.weekday(); bg = "bg-weekend" if wd >= 5 else ""
        date_th1 += f"<th class='{header_cls} col_day'>{curr.day}</th>"; date_th2 += f"<th class='{bg} col_day'>{weekdays[wd]}</th>"; curr += timedelta(days=1)
    cols_def = ["Station", "Location", "Program", "Day-part", "Size", "rate<br>(Net)", "Package-cost<br>(Net)"]
    if format_type == "Shenghuo": cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "秒數/規格", "單價", "金額"]
    elif format_type == "Bolin": cols_def = ["頻道", "播出地區", "播出店數", "播出時間", "規格", "單價", "金額"]
    th_fixed = "".join([f"<th rowspan='2' class='{header_cls}'>{c}</th>" for c in cols_def])
    unique_media = sorted(list(set([r['media'] for r in rows]))); medium_str = "/".join(unique_media) if format_type == "Dongwu" else "全家廣播/新鮮視/家樂福"
    tbody = ""; rows_sorted = sorted(rows, key=lambda x: ({"全家廣播":1,"新鮮視":2,"家樂福":3}.get(x["media"],9), x["seconds"]))
    for key, group in groupby(rows_sorted, lambda x: (x['media'], x['seconds'], x.get('nat_pkg_display', 0))):
        g_list = list(group); g_size = len(g_list); is_pkg = g_list[0]['is_pkg_member']
        for i, r in enumerate(g_list):
            tbody += "<tr>"; rate = f"${r['rate_display']:,}" if isinstance(r['rate_display'], (int, float)) else r['rate_display']
            pkg_val_str = ""
            if is_pkg:
                if i == 0: val = f"${r['nat_pkg_display']:,}"; pkg_val_str = f"<td class='right' rowspan='{g_size}'>{val}</td>"
            else:
                val = f"${r['pkg_display']:,}" if isinstance(r['pkg_display'], (int, float)) else r['pkg_display']; pkg_val_str = f"<td class='right'>{val}</td>"
            if format_type == "Shenghuo": sec_txt = f"{r['seconds']}秒"; tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{sec_txt}</td><td>{rate}</td>{pkg_val_str}"
            elif format_type == "Bolin": tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{r['seconds']}秒</td><td>{rate}</td>{pkg_val_str}"
            else: tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{r['seconds']}</td><td>{rate}</td>{pkg_val_str}"
            for d in r['schedule'][:eff_days]: tbody += f"<td>{d}</td>"
            tbody += "</tr>"
    remarks_html = "<br>".join([html_escape(x) for x in remarks])
    vat = int(round(budget * 0.05)); footer_html = f"<div style='margin-top:10px; font-weight:bold; text-align:right;'>製作費: ${prod:,}<br>5% VAT: ${vat:,}<br>Grand Total: ${grand_total:,}</div>"
    return f"<html><head><style>body {{ font-family: sans-serif; font-size: 10px; }} table {{ border-collapse: collapse; width: 100%; }} th, td {{ border: 0.5pt solid #000; padding: 4px; text-align: center; white-space: nowrap; }} .bg-dw-head {{ background-color: #4472C4; color: white; }} .bg-sh-head {{ background-color: white; color: black; font-weight: bold; border-bottom: 2px solid black; }} .bg-bolin-head {{ background-color: #F8CBAD; color: black; }} .bg-weekend {{ background-color: #FFFFCC; }}</style></head><body><div style='margin-bottom:10px;'><b>客戶名稱：</b>{html_escape(c_name)} &nbsp; <b>Product：</b>{html_escape(p_display)}<br><b>Period：</b>{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')} &nbsp; <b>Medium：</b>{html_escape(medium_str)}</div><div style='overflow-x:auto;'><table><thead><tr>{th_fixed}{date_th1}</tr><tr>{date_th2}</tr></thead><tbody>{tbody}</tbody></table></div>{footer_html}<div style='margin-top:10px; font-size:11px;'><b>Remarks：</b><br>{remarks_html}</div></body></html>"

# =========================================================
# 5. 資料運算
# =========================================================
@st.cache_data(ttl=300)
def load_config_from_cloud(share_url):
    try:
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", share_url)
        if not match: return None, None, None, None, "連結格式錯誤"
        file_id = match.group(1)
        def read_sheet(sheet_name):
            url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
            return pd.read_csv(url)
        df_store = read_sheet("Stores"); df_store.columns = [c.strip() for c in df_store.columns]
        store_counts = dict(zip(df_store['Key'], df_store['Display_Name'])); store_counts_num = dict(zip(df_store['Key'], df_store['Count']))
        df_fact = read_sheet("Factors"); df_fact.columns = [c.strip() for c in df_fact.columns]
        sec_factors = {}
        for _, row in df_fact.iterrows():
            if row['Media'] not in sec_factors: sec_factors[row['Media']] = {}
            sec_factors[row['Media']][int(row['Seconds'])] = float(row['Factor'])
        name_map = {"全家新鮮視": "新鮮視", "全家廣播": "全家廣播", "家樂福": "家樂福"}
        for k, v in name_map.items():
            if k in sec_factors and v not in sec_factors: sec_factors[v] = sec_factors[k]
        df_price = read_sheet("Pricing"); df_price.columns = [c.strip() for c in df_price.columns]
        pricing_db = {}
        for _, row in df_price.iterrows():
            m = row['Media']; r = row['Region']
            if m == "家樂福":
                if m not in pricing_db: pricing_db[m] = {}
                pricing_db[m][r] = {"List": int(row['List_Price']), "Net": int(row['Net_Price']), "Std_Spots": int(row['Std_Spots']), "Day_Part": row['Day_Part']}
            else:
                if m not in pricing_db: pricing_db[m] = {"Std_Spots": int(row['Std_Spots']), "Day_Part": row['Day_Part']}
                pricing_db[m][r] = [int(row['List_Price']), int(row['Net_Price'])]
        return store_counts, store_counts_num, pricing_db, sec_factors, None
    except Exception as e: return None, None, None, None, f"讀取失敗: {str(e)}"

def calculate_plan_data(config, total_budget, days_count, pricing_db, sec_factors, store_counts_num, regions_order):
    rows = []; total_list_accum = 0; debug_logs = []
    for m, cfg in config.items():
        m_budget_total = total_budget * (cfg["share"] / 100.0)
        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget_total * (sec_pct / 100.0)
            if s_budget <= 0: continue
            factor = get_sec_factor(m, sec, sec_factors)
            if m in ["全家廣播", "新鮮視"]:
                db = pricing_db[m]
                calc_regs = ["全省"] if cfg["is_national"] else cfg["regions"]
                display_regs = regions_order if cfg["is_national"] else cfg["regions"]
                unit_net_sum = 0
                for r in calc_regs: unit_net_sum += (db[r][1] / db["Std_Spots"]) * factor
                if unit_net_sum == 0: continue
                spots_init = math.ceil(s_budget / unit_net_sum); is_under_target = spots_init < db["Std_Spots"]
                calc_penalty = 1.1 if is_under_target else 1.0 
                if cfg["is_national"]: row_display_penalty = 1.0; total_display_penalty = 1.1 if is_under_target else 1.0
                else: row_display_penalty = 1.1 if is_under_target else 1.0; total_display_penalty = 1.0 
                spots_final = math.ceil(s_budget / (unit_net_sum * calc_penalty))
                if spots_final % 2 != 0: spots_final += 1
                if spots_final == 0: spots_final = 2
                sch = calculate_schedule(spots_final, days_count); nat_pkg_display = 0
                if cfg["is_national"]:
                    nat_list = db["全省"][0]; nat_unit_price = int((nat_list / db["Std_Spots"]) * factor * total_display_penalty)
                    nat_pkg_display = nat_unit_price * spots_final; total_list_accum += nat_pkg_display
                for i, r in enumerate(display_regs):
                    list_price_region = db[r][0]
                    unit_rate_display = int((list_price_region / db["Std_Spots"]) * factor * row_display_penalty)
                    total_rate_display = unit_rate_display * spots_final; row_pkg_display = total_rate_display
                    if not cfg["is_national"]: total_list_accum += row_pkg_display
                    rows.append({
                        "media": m, "region": r, "program_num": store_counts_num.get(f"新鮮視_{r}" if m=="新鮮視" else r, 0),
                        "daypart": db["Day_Part"], "seconds": sec, "spots": spots_final, "schedule": sch,
                        "rate_display": total_rate_display, "pkg_display": row_pkg_display, "is_pkg_member": cfg["is_national"], "nat_pkg_display": nat_pkg_display
                    })
            elif m == "家樂福":
                db = pricing_db["家樂福"]; base_std = db["量販_全省"]["Std_Spots"]
                unit_net = (db["量販_全省"]["Net"] / base_std) * factor
                spots_init = math.ceil(s_budget / unit_net); penalty = 1.1 if spots_init < base_std else 1.0
                spots_final = math.ceil(s_budget / (unit_net * penalty))
                if spots_final % 2 != 0: spots_final += 1
                sch_h = calculate_schedule(spots_final, days_count)
                base_list = db["量販_全省"]["List"]; unit_rate_h = int((base_list / base_std) * factor * penalty)
                total_rate_h = unit_rate_h * spots_final; total_list_accum += total_rate_h
                rows.append({"media": m, "region": "全省量販", "program_num": store_counts_num["家樂福_量販"], "daypart": db["量販_全省"]["Day_Part"], "seconds": sec, "spots": spots_final, "schedule": sch_h, "rate_display": total_rate_h, "pkg_display": total_rate_h, "is_pkg_member": False})
                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std)); sch_s = calculate_schedule(spots_s, days_count)
                rows.append({"media": m, "region": "全省超市", "program_num": store_counts_num["家樂福_超市"], "daypart": db["超市_全省"]["Day_Part"], "seconds": sec, "spots": spots_s, "schedule": sch_s, "rate_display": "計量販", "pkg_display": "計量販", "is_pkg_member": False})
    return rows, total_list_accum, debug_logs

# =========================================================
# 7. Render Engines (Optimized with Object Pooling & Caching)
# =========================================================

@st.cache_data(show_spinner="正在生成 Excel 報表...", ttl=3600)
def generate_excel_from_scratch(format_type, start_dt, end_dt, client_name, product_name, rows, remarks_list, final_budget_val, prod_cost):
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
    from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker

    SIDE_THIN = Side(style=BS_THIN); SIDE_MEDIUM = Side(style=BS_MEDIUM); SIDE_HAIR = Side(style=BS_HAIR)
    BORDER_ALL_THIN = Border(top=SIDE_THIN, bottom=SIDE_THIN, left=SIDE_THIN, right=SIDE_THIN)
    BORDER_ALL_MEDIUM = Border(top=SIDE_MEDIUM, bottom=SIDE_MEDIUM, left=SIDE_MEDIUM, right=SIDE_MEDIUM)
    ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ALIGN_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ALIGN_RIGHT = Alignment(horizontal='right', vertical='center', wrap_text=True)
    FONT_STD = Font(name=FONT_MAIN, size=12)
    FONT_BOLD = Font(name=FONT_MAIN, size=14, bold=True)
    FONT_TITLE = Font(name=FONT_MAIN, size=48, bold=True)
    FILL_WEEKEND = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
    FILL_HEADER_BOLIN = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")

    def set_border(cell, top=None, bottom=None, left=None, right=None):
        cur = cell.border
        new_top = Side(style=top) if top else cur.top
        new_bottom = Side(style=bottom) if bottom else cur.bottom
        new_left = Side(style=left) if left else cur.left
        new_right = Side(style=right) if right else cur.right
        cell.border = Border(top=new_top, bottom=new_bottom, left=new_left, right=new_right)

    def draw_outer_border_fast(ws, min_r, max_r, min_c, max_c):
        for c in range(min_c, max_c + 1):
            cell = ws.cell(min_r, c); cur = cell.border
            cell.border = Border(top=SIDE_MEDIUM, bottom=cur.bottom, left=cur.left, right=cur.right)
            cell = ws.cell(max_r, c); cur = cell.border
            cell.border = Border(top=cur.top, bottom=SIDE_MEDIUM, left=cur.left, right=cur.right)
        for r in range(min_r, max_r + 1):
            cell = ws.cell(r, min_c); cur = cell.border
            cell.border = Border(top=cur.top, bottom=cur.bottom, left=SIDE_MEDIUM, right=cur.right)
            cell = ws.cell(r, max_c); cur = cell.border
            cell.border = Border(top=cur.top, bottom=cur.bottom, left=cur.left, right=SIDE_MEDIUM)

    # -------------------------------------------------------------
    # Render Logic: Dongwu
    # -------------------------------------------------------------
    def render_dongwu_optimized(ws, start_dt, end_dt, rows, budget, prod):
        eff_days = (end_dt - start_dt).days + 1
        spots_col_idx = 7 + eff_days + 1
        total_cols = spots_col_idx

        COL_WIDTHS = {'A': 19.6, 'B': 22.8, 'C': 14.6, 'D': 20.0, 'E': 13.0, 'F': 19.6, 'G': 17.9}
        ROW_HEIGHTS = {1: 61.0, 2: 29.0, 3: 40.0, 4: 40.0, 5: 40.0, 6: 40.0, 7: 40.0, 8: 40.0}
        for k, v in COL_WIDTHS.items(): ws.column_dimensions[k].width = v
        for i in range(eff_days): ws.column_dimensions[get_column_letter(8+i)].width = 8.5
        ws.column_dimensions[get_column_letter(spots_col_idx)].width = 13.0
        for r, h in ROW_HEIGHTS.items(): ws.row_dimensions[r].height = h

        ws.merge_cells(f"A1:{get_column_letter(total_cols)}1"); c = ws['A1']; c.value = "Media Schedule"; c.font = FONT_TITLE; c.alignment = ALIGN_CENTER
        unique_media = sorted(list(set([r['media'] for r in rows]))); order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}; unique_media.sort(key=lambda x: order.get(x, 99)); medium_str = "/".join(unique_media)
        unique_secs = sorted(list(set([r['seconds'] for r in rows]))); p_str = f"{'、'.join([f'{s}秒' for s in unique_secs])} {product_name}"
        infos = [("A3", "客戶名稱：", client_name), ("A4", "Product：", p_str), 
                 ("A5", "Period :", f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"), ("A6", "Medium :", medium_str)]
        for pos, lbl, val in infos:
            c = ws[pos]; c.value = lbl; c.font = FONT_BOLD; c.alignment = Alignment(vertical='center')
            c2 = ws.cell(c.row, 2); c2.value = val; c2.font = FONT_BOLD; c2.alignment = Alignment(vertical='center')
        
        for c_idx in range(1, total_cols + 1): set_border(ws.cell(3, c_idx), top=BS_MEDIUM)

        ws['H6'] = f"{start_dt.month}月"; ws['H6'].font = Font(name=FONT_MAIN, size=16, bold=True); ws['H6'].alignment = ALIGN_CENTER
        
        headers = [("A","Station"), ("B","Location"), ("C","Program"), ("D","Day-part"), ("E","Size"), ("F","rate\n(Net)"), ("G","Package-cost\n(Net)")]
        for col, txt in headers:
            col_idx = column_index_from_string(col)
            ws.merge_cells(f"{col}7:{col}8"); c7 = ws.cell(7, col_idx); c7.value = txt; c8 = ws.cell(8, col_idx)
            c7.font = FONT_BOLD; c7.alignment = ALIGN_CENTER
            c7.border = BORDER_ALL_THIN; c8.border = BORDER_ALL_THIN
            set_border(c7, top=BS_MEDIUM); set_border(c8, bottom=BS_MEDIUM)

        curr = start_dt
        for i in range(eff_days):
            col_idx = 8 + i; c_d = ws.cell(7, col_idx); c_w = ws.cell(8, col_idx)
            c_d.value = curr; c_d.number_format = 'm/d'; c_w.value = ["一","二","三","四","五","六","日"][curr.weekday()]
            if curr.weekday() >= 5: c_w.fill = FILL_WEEKEND
            curr += timedelta(days=1)
            c_d.font = FONT_STD; c_w.font = FONT_STD; c_d.alignment = ALIGN_CENTER; c_w.alignment = ALIGN_CENTER
            c_d.border = BORDER_ALL_THIN; c_w.border = BORDER_ALL_THIN
            set_border(c_d, top=BS_MEDIUM); set_border(c_w, bottom=BS_MEDIUM)

        c_spots_7 = ws.cell(7, spots_col_idx); c_spots_7.value = "檔次"
        c_spots_8 = ws.cell(8, spots_col_idx)
        ws.merge_cells(start_row=7, start_column=spots_col_idx, end_row=8, end_column=spots_col_idx)
        c_spots_7.font = FONT_BOLD; c_spots_7.alignment = ALIGN_CENTER
        c_spots_7.border = BORDER_ALL_THIN; c_spots_8.border = BORDER_ALL_THIN
        set_border(c_spots_7, top=BS_MEDIUM, left=BS_MEDIUM); set_border(c_spots_8, bottom=BS_MEDIUM, left=BS_MEDIUM)
        set_border(ws['A7'], right=BS_MEDIUM); set_border(ws['A8'], right=BS_MEDIUM)

        curr_row = 9; grouped_data = {
            "全家廣播": sorted([r for r in rows if r["media"] == "全家廣播"], key=lambda x: x["seconds"]),
            "新鮮視": sorted([r for r in rows if r["media"] == "新鮮視"], key=lambda x: x["seconds"]),
            "家樂福": sorted([r for r in rows if r["media"] == "家樂福"], key=lambda x: x["seconds"]),
        }
        total_rate_sum = 0 

        for m_key, data in grouped_data.items():
            if not data: continue
            start_merge = curr_row
            display_name = f"全家便利商店\n{m_key if m_key!='家樂福' else ''}廣告"
            if m_key == "家樂福": display_name = "家樂福"
            elif m_key == "全家廣播": display_name = "全家便利商店\n通路廣播廣告"
            elif m_key == "新鮮視": display_name = "全家便利商店\n新鮮視廣告"

            for idx, r in enumerate(data):
                ws.row_dimensions[curr_row].height = 40
                ws.cell(curr_row, 1, display_name).alignment = ALIGN_CENTER
                ws.cell(curr_row, 2, r["region"]).alignment = ALIGN_CENTER
                ws.cell(curr_row, 3, r.get("program_num", 0)).alignment = ALIGN_CENTER
                ws.cell(curr_row, 4, r["daypart"]).alignment = ALIGN_CENTER
                ws.cell(curr_row, 5, f"{r['seconds']}秒").alignment = ALIGN_CENTER
                
                rate = r['rate_display']; pkg = r['pkg_display']
                if isinstance(rate, (int, float)): total_rate_sum += rate
                if r.get("is_pkg_member"): pkg = r['nat_pkg_display'] if idx == 0 else None
                c_rate = ws.cell(curr_row, 6); c_rate.value = rate; c_rate.number_format = FMT_MONEY; c_rate.alignment = ALIGN_CENTER
                if pkg is not None:
                    c_pkg = ws.cell(curr_row, 7); c_pkg.value = pkg; c_pkg.number_format = FMT_MONEY; c_pkg.alignment = ALIGN_CENTER

                row_sum = 0
                for d_idx in range(eff_days):
                    if d_idx < len(r["schedule"]):
                        val = r["schedule"][d_idx]; row_sum += val
                        c_s = ws.cell(curr_row, 8+d_idx); c_s.value = val; c_s.number_format = FMT_NUMBER; c_s.alignment = ALIGN_CENTER
                ws.cell(curr_row, spots_col_idx, row_sum).alignment = ALIGN_CENTER
                for c_idx in range(1, total_cols + 1):
                    cell = ws.cell(curr_row, c_idx); cell.font = FONT_STD; cell.border = BORDER_ALL_THIN
                curr_row += 1

            ws.merge_cells(start_row=start_merge, start_column=1, end_row=curr_row-1, end_column=1)
            if data[0].get("is_pkg_member"): ws.merge_cells(start_row=start_merge, start_column=7, end_row=curr_row-1, end_column=7)
            for col in [4, 5]:
                m_start = start_merge
                while m_start < curr_row:
                    m_end = m_start; val = ws.cell(m_start, col).value
                    while m_end + 1 < curr_row and ws.cell(m_end+1, col).value == val: m_end += 1
                    if m_end > m_start: ws.merge_cells(start_row=m_start, start_column=col, end_row=m_end, end_column=col)
                    m_start = m_end + 1
            draw_outer_border_fast(ws, start_merge, curr_row-1, 1, total_cols)
            for r in range(start_merge, curr_row):
                set_border(ws.cell(r, 1), right=BS_MEDIUM); set_border(ws.cell(r, spots_col_idx), left=BS_MEDIUM)

        ws.row_dimensions[curr_row].height = 30
        c_lbl = ws.cell(curr_row, 5, "Total"); c_lbl.alignment = ALIGN_CENTER; c_lbl.font = FONT_BOLD
        c_rate_sum = ws.cell(curr_row, 6, total_rate_sum); c_rate_sum.number_format = FMT_MONEY; c_rate_sum.alignment = ALIGN_CENTER; c_rate_sum.font = FONT_BOLD
        c_val = ws.cell(curr_row, 7, budget); c_val.number_format = FMT_MONEY; c_val.alignment = ALIGN_CENTER; c_val.font = FONT_BOLD
        total_spots_all = 0
        for d_idx in range(eff_days):
            daily_sum = sum([r['schedule'][d_idx] for r in rows if d_idx < len(r['schedule'])]); total_spots_all += daily_sum
            c = ws.cell(curr_row, 8+d_idx); c.value = daily_sum; c.alignment = ALIGN_CENTER; c.font = FONT_STD; c.number_format = FMT_NUMBER
        ws.cell(curr_row, spots_col_idx, total_spots_all).alignment = ALIGN_CENTER; ws.cell(curr_row, spots_col_idx).font = FONT_STD
        
        for c_idx in range(1, total_cols + 1): set_border(ws.cell(curr_row, c_idx), top=BS_MEDIUM, bottom=BS_MEDIUM, left=BS_THIN, right=BS_THIN)
        set_border(ws.cell(curr_row, 1), left=BS_MEDIUM, right=BS_MEDIUM); set_border(ws.cell(curr_row, spots_col_idx), left=BS_MEDIUM, right=BS_MEDIUM)
        curr_row += 1

        vat = int(budget * 0.05); grand_total = budget + vat
        footer_items = [("媒體", budget), ("製作", prod), ("5% VAT", vat), ("Grand Total", grand_total)]
        for label, val in footer_items:
            if label == "媒體": continue 
            ws.row_dimensions[curr_row].height = 30
            c_l = ws.cell(curr_row, 6); c_l.value = label; c_l.alignment = ALIGN_LEFT; c_l.font = FONT_STD
            c_v = ws.cell(curr_row, 7); c_v.value = val; c_v.number_format = FMT_MONEY; c_v.alignment = ALIGN_CENTER; c_v.font = FONT_STD
            set_border(c_l, left=BS_MEDIUM, top=BS_THIN, bottom=BS_THIN, right=BS_THIN)
            set_border(c_v, right=BS_MEDIUM, top=BS_THIN, bottom=BS_THIN, left=BS_THIN)
            if label == "Grand Total":
                for c_idx in range(1, total_cols + 1): set_border(ws.cell(curr_row, c_idx), top=BS_MEDIUM, bottom=BS_MEDIUM)
            curr_row += 1
        draw_outer_border_fast(ws, 7, curr_row-1, 1, total_cols)
        
        curr_row += 1; ws.cell(curr_row, 1, "Remarks:").font = Font(name=FONT_MAIN, size=16, bold=True, underline='single')
        for rm in remarks_list:
            curr_row += 1
            is_red = rm.strip().startswith("1.") or rm.strip().startswith("4.")
            c = ws.cell(curr_row, 1); c.value = rm; c.font = Font(name=FONT_MAIN, size=14, color="FF0000" if is_red else "000000")

        # --- v111.1: Signature Block (Polish) ---
        curr_row += 2 # Spacer
        sig_start = curr_row
        
        ws.merge_cells(start_row=sig_start, start_column=1, end_row=sig_start, end_column=7) 
        c_l1 = ws.cell(sig_start, 1); c_l1.value = "甲    方：東吳廣告股份有限公司"; c_l1.font = FONT_STD; c_l1.alignment = ALIGN_LEFT
        ws.merge_cells(start_row=sig_start+1, start_column=1, end_row=sig_start+1, end_column=7) 
        c_l2 = ws.cell(sig_start+1, 1); c_l2.value = "統一編號：20935458"; c_l2.font = FONT_STD; c_l2.alignment = ALIGN_LEFT
        
        right_start_col = 20 # Column T
        ws.merge_cells(start_row=sig_start, start_column=right_start_col, end_row=sig_start, end_column=right_start_col+7) 
        c_r1 = ws.cell(sig_start, right_start_col); c_r1.value = f"乙    方：{client_name}"; c_r1.font = FONT_STD; c_r1.alignment = ALIGN_LEFT
        ws.merge_cells(start_row=sig_start+1, start_column=right_start_col, end_row=sig_start+1, end_column=right_start_col+7)
        c_r2 = ws.cell(sig_start+1, right_start_col); c_r2.value = "統一編號："; c_r2.font = FONT_STD; c_r2.alignment = ALIGN_LEFT
        ws.merge_cells(start_row=sig_start+2, start_column=right_start_col, end_row=sig_start+2, end_column=right_start_col+7)
        c_r3 = ws.cell(sig_start+2, right_start_col); c_r3.value = "客戶簽章："; c_r3.font = FONT_STD; c_r3.alignment = ALIGN_LEFT

        for c_idx in range(1, total_cols + 1): set_border(ws.cell(sig_start, c_idx), top=BS_THIN)

        return curr_row + 3

    # -------------------------------------------------------------
    # Render Logic: Shenghuo (v111.13 Row/Font Fix)
    # -------------------------------------------------------------
    def render_shenghuo_optimized(ws, start_dt, end_dt, rows, budget, prod):
        eff_days = (end_dt - start_dt).days + 1
        end_c_start = 6 + eff_days
        total_cols = end_c_start + 2

        ws.column_dimensions['A'].width = 22.5; ws.column_dimensions['B'].width = 24.5; ws.column_dimensions['C'].width = 13.8; ws.column_dimensions['D'].width = 19.4; ws.column_dimensions['E'].width = 15.0
        for i in range(eff_days): ws.column_dimensions[get_column_letter(6 + i)].width = 8.1 
        ws.column_dimensions[get_column_letter(end_c_start)].width = 9.5 
        ws.column_dimensions[get_column_letter(end_c_start+1)].width = 58.0 
        ws.column_dimensions[get_column_letter(end_c_start+2)].width = 20.0 
        
        # (1) & (2) Row Height Fixes
        ROW_H_MAP = {1:30, 2:30, 3:46, 4:46, 5:40, 6:40, 7:35, 8:35} # Rows 3,4 -> 46; Rows 5,6 -> 40
        for r, h in ROW_H_MAP.items(): ws.row_dimensions[r].height = h
        
        ws.merge_cells(f"A1:{get_column_letter(total_cols)}1"); c1 = ws['A1']; c1.value = "聲活數位-媒體計劃排程表"; c1.font = Font(name=FONT_MAIN, size=24, bold=True); c1.alignment = ALIGN_CENTER
        ws.merge_cells(f"A2:{get_column_letter(total_cols)}2"); c2 = ws['A2']; c2.value = "Media Schedule"; c2.font = Font(name=FONT_MAIN, size=18, bold=True); c2.alignment = ALIGN_CENTER
        
        # (1) Font Size 16 for Rows 3, 4
        FONT_16 = Font(name=FONT_MAIN, size=16)
        ws.merge_cells(f"A3:{get_column_letter(total_cols)}3"); ws['A3'].value = "聲活數位科技股份有限公司 統編 28710100"; ws['A3'].font = FONT_16; ws['A3'].alignment = ALIGN_LEFT
        ws.merge_cells(f"A4:{get_column_letter(total_cols)}4"); ws['A4'].value = "蔡伊閔"; ws['A4'].font = FONT_16; ws['A4'].alignment = ALIGN_LEFT
        
        unique_secs = sorted(list(set([r['seconds'] for r in rows]))); sec_str = " ".join([f"{s}秒廣告" for s in unique_secs])
        space_gap = "　" * 10
        period_str = f"執行期間：{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')}"
        info_text = f"客戶名稱：{client_name}{space_gap}廣告規格：{sec_str}"
        
        # (2) Font Size 14 for Rows 5, 6
        FONT_14 = Font(name=FONT_MAIN, size=14)
        mid_split_col = end_c_start
        ws.merge_cells(f"A5:{get_column_letter(mid_split_col)}5")
        c5 = ws['A5']; c5.value = info_text; c5.font = FONT_14; c5.alignment = ALIGN_LEFT 
        
        ws.merge_cells(f"{get_column_letter(end_c_start+1)}5:{get_column_letter(total_cols)}5")
        c5_r = ws[f"{get_column_letter(end_c_start+1)}5"]; c5_r.value = period_str; c5_r.font = FONT_14; c5_r.alignment = ALIGN_LEFT 
        
        draw_outer_border_fast(ws, 5, 5, 1, total_cols)

        ws.merge_cells(f"A6:{get_column_letter(total_cols)}6")
        c6 = ws['A6']; c6.value = f"廣告名稱：{product_name}"; c6.font = FONT_14; c6.alignment = ALIGN_LEFT
        draw_outer_border_fast(ws, 6, 6, 1, total_cols)
        
        headers = ["頻道", "播出地區", "播出店數", "播出時間", "秒數\n規格"]
        for i, h in enumerate(headers):
            c_idx = i + 1
            ws.merge_cells(start_row=7, start_column=c_idx, end_row=8, end_column=c_idx)
            c = ws.cell(7, c_idx); c.value = h; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER
            t, b, l, r = BS_MEDIUM, BS_THIN, BS_THIN, BS_THIN
            if c_idx == 1: l = BS_MEDIUM
            c.border = Border(top=Side(style=t), bottom=Side(style=b), left=Side(style=l), right=Side(style=r))
            ws.cell(8, c_idx).border = Border(top=Side(style=BS_THIN), bottom=Side(style=BS_THIN), left=Side(style=l), right=Side(style=r))

        curr = start_dt
        month_groups = []
        for i in range(eff_days):
            d = start_dt + timedelta(days=i); m_key = (d.year, d.month)
            if not month_groups or month_groups[-1][0] != m_key: month_groups.append([m_key, i, i]) 
            else: month_groups[-1][2] = i
        
        for m_key, s_idx, e_idx in month_groups:
            start_col = 6 + s_idx; end_col = 6 + e_idx
            ws.merge_cells(start_row=7, start_column=start_col, end_row=7, end_column=end_col)
            c = ws.cell(7, start_col); c.value = f"{m_key[1]}月"; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER; c.border = BORDER_ALL_MEDIUM
            for c_i in range(start_col, end_col+1): set_border(ws.cell(7, c_i), top=BS_MEDIUM, bottom=BS_MEDIUM, left=BS_THIN, right=BS_THIN)
            set_border(ws.cell(7, start_col), left=BS_MEDIUM); set_border(ws.cell(7, end_col), right=BS_MEDIUM)

        curr = start_dt
        for i in range(eff_days):
            col_idx = 6 + i
            c = ws.cell(8, col_idx); c.value = curr.day; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER; c.border = BORDER_ALL_MEDIUM
            if curr.weekday() >= 5: c.fill = FILL_WEEKEND
            curr += timedelta(days=1)

        end_headers = ["檔次", "定價", "專案價"]
        for i, h in enumerate(end_headers):
            c_idx = end_c_start + i
            ws.merge_cells(start_row=7, start_column=c_idx, end_row=8, end_column=c_idx)
            c = ws.cell(7, c_idx); c.value = h; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER
            t, b, l, r = BS_MEDIUM, BS_THIN, BS_THIN, BS_THIN
            if c_idx == total_cols: r = BS_MEDIUM
            c.border = Border(top=Side(style=t), bottom=Side(style=b), left=Side(style=l), right=Side(style=r))
            ws.cell(8, c_idx).border = Border(top=Side(style=BS_THIN), bottom=Side(style=BS_THIN), left=Side(style=l), right=Side(style=r))

        date_start_col = 6
        for c_idx in range(date_start_col, total_cols + 1):
            c7 = ws.cell(7, c_idx)
            c7.border = Border(top=Side(style=BS_MEDIUM), bottom=Side(style=BS_THIN), left=Side(style=BS_THIN), right=Side(style=BS_THIN))
            if c_idx == date_start_col: set_border(c7, left=BS_MEDIUM)
            if c_idx == total_cols: set_border(c7, right=BS_MEDIUM)
            c8 = ws.cell(8, c_idx)
            c8.border = Border(top=Side(style=BS_THIN), bottom=Side(style=BS_THIN), left=Side(style=BS_THIN), right=Side(style=BS_THIN))
            if c_idx == date_start_col: set_border(c8, left=BS_MEDIUM)
            if c_idx == total_cols: set_border(c8, right=BS_MEDIUM)

        curr_row = 9
        grouped_data = {"全家廣播": sorted([r for r in rows if r["media"]=="全家廣播"], key=lambda x:x['seconds']),
                        "新鮮視": sorted([r for r in rows if r["media"]=="新鮮視"], key=lambda x:x['seconds']),
                        "家樂福": sorted([r for r in rows if r["media"]=="家樂福"], key=lambda x:x['seconds'])}
        
        total_store_count = 0; total_list_sum = 0
        for m_key, data in grouped_data.items():
            if not data: continue
            start_merge = curr_row
            d_name = f"全家便利商店\n{m_key}廣告" if m_key != "家樂福" else "家樂福"
            for idx, r in enumerate(data):
                ws.row_dimensions[curr_row].height = 40
                ws.cell(curr_row, 1, d_name).alignment = ALIGN_CENTER
                ws.cell(curr_row, 2, r['region']).alignment = ALIGN_CENTER
                p_num = int(r.get('program_num', 0)); total_store_count += p_num 
                suffix = "面" if m_key == "新鮮視" else "店"
                ws.cell(curr_row, 3, f"{p_num:,}{suffix}").alignment = ALIGN_CENTER
                ws.cell(curr_row, 4, r['daypart']).alignment = ALIGN_CENTER
                sec = r['seconds']
                if m_key == "新鮮視": sec_txt = f"{sec}秒\n影片/影像 1920x1080 (mp4)"
                else: sec_txt = f"{sec}秒廣告"
                c_spec = ws.cell(curr_row, 5, sec_txt); c_spec.alignment = ALIGN_CENTER; c_spec.font = Font(name=FONT_MAIN, size=10)
                row_sum = 0
                for d_idx in range(eff_days):
                    if d_idx < len(r['schedule']):
                        val = r['schedule'][d_idx]; row_sum += val
                        c = ws.cell(curr_row, 6+d_idx); c.value = val; c.alignment = ALIGN_CENTER; c.font = FONT_STD; c.border = BORDER_ALL_THIN
                ws.cell(curr_row, end_c_start, row_sum).alignment = ALIGN_CENTER
                rate_val = r['rate_display']
                if isinstance(rate_val, (int, float)): total_list_sum += rate_val
                ws.cell(curr_row, end_c_start+1, rate_val).number_format = FMT_MONEY; ws.cell(curr_row, end_c_start+1).alignment = ALIGN_CENTER 
                pkg = r['pkg_display']
                if r.get('is_pkg_member'): pkg = r['nat_pkg_display'] if idx == 0 else None
                ws.cell(curr_row, end_c_start+2, pkg).number_format = FMT_MONEY; ws.cell(curr_row, end_c_start+2).alignment = ALIGN_CENTER
                for c_idx in range(1, total_cols + 1):
                    c = ws.cell(curr_row, c_idx); c.font = FONT_STD; c.border = BORDER_ALL_THIN
                set_border(ws.cell(curr_row, 5), right=BS_MEDIUM)
                curr_row += 1
            ws.merge_cells(start_row=start_merge, start_column=1, end_row=curr_row-1, end_column=1)
            if data[0].get('is_pkg_member'): ws.merge_cells(start_row=start_merge, start_column=end_c_start+2, end_row=curr_row-1, end_column=end_c_start+2)
            draw_outer_border_fast(ws, start_merge, curr_row-1, 1, total_cols)

        # Total Row
        ws.row_dimensions[curr_row].height = 40
        ws.cell(curr_row, 3, total_store_count).number_format = FMT_NUMBER; ws.cell(curr_row, 3).alignment = ALIGN_CENTER; ws.cell(curr_row, 3).font = FONT_BOLD
        ws.cell(curr_row, 5, "Total").alignment = ALIGN_CENTER; ws.cell(curr_row, 5).font = FONT_BOLD
        for d_idx in range(eff_days):
            daily_sum = sum([r['schedule'][d_idx] for r in rows if d_idx < len(r['schedule'])])
            c = ws.cell(curr_row, 6+d_idx); c.value = daily_sum; c.alignment = ALIGN_CENTER; c.font = FONT_BOLD
        ws.cell(curr_row, end_c_start, sum([sum(r['schedule']) for r in rows])).alignment = ALIGN_CENTER; ws.cell(curr_row, end_c_start).font = FONT_BOLD
        ws.cell(curr_row, end_c_start+1, total_list_sum).number_format = FMT_MONEY; ws.cell(curr_row, end_c_start+1).font = FONT_BOLD; ws.cell(curr_row, end_c_start+1).alignment = ALIGN_CENTER
        ws.cell(curr_row, end_c_start+2, budget).number_format = FMT_MONEY; ws.cell(curr_row, end_c_start+2).font = FONT_BOLD; ws.cell(curr_row, end_c_start+2).alignment = ALIGN_CENTER
        for c_idx in range(1, total_cols+1): ws.cell(curr_row, c_idx).border = BORDER_ALL_THIN
        draw_outer_border_fast(ws, curr_row, curr_row, 1, total_cols)
        for c_idx in range(1, total_cols+1): set_border(ws.cell(curr_row, c_idx), bottom=BS_MEDIUM)
        set_border(ws.cell(curr_row, 5), right=BS_MEDIUM)
        curr_row += 1

        # Footer
        vat = int(budget * 0.05); grand_total = budget + vat
        footer_stack = [("製作", prod), ("5% VAT", vat), ("Grand Total", grand_total)]
        for lbl, val in footer_stack:
            ws.row_dimensions[curr_row].height = 30
            c_l = ws.cell(curr_row, end_c_start+1); c_l.value = lbl; c_l.alignment = ALIGN_RIGHT; c_l.font = FONT_STD
            c_v = ws.cell(curr_row, end_c_start+2); c_v.value = val; c_v.number_format = FMT_MONEY; c_v.alignment = ALIGN_CENTER; c_v.font = FONT_BOLD 
            t, b, l, r = BS_THIN, BS_THIN, BS_MEDIUM, BS_THIN
            if lbl == "Grand Total": b = BS_MEDIUM 
            c_l.border = Border(top=Side(style=t), bottom=Side(style=b), left=Side(style=l), right=Side(style=r))
            t, b, l, r = BS_THIN, BS_THIN, BS_THIN, BS_MEDIUM
            if lbl == "Grand Total": b = BS_MEDIUM 
            c_v.border = Border(top=Side(style=t), bottom=Side(style=b), left=Side(style=l), right=Side(style=r))
            if lbl == "Grand Total":
                for c_idx in range(1, total_cols + 1): set_border(ws.cell(curr_row, c_idx), bottom=BS_MEDIUM)
            curr_row += 1
        
        curr_row += 1
        ws.cell(curr_row, 1, "Remarks:").font = Font(name=FONT_MAIN, size=16, bold=True)
        for rm in remarks_list:
            curr_row += 1
            is_red = rm.strip().startswith("1.") or rm.strip().startswith("4.")
            is_blue = rm.strip().startswith("6.")
            color = "000000"
            if is_red: color = "FF0000"
            if is_blue: color = "0000FF"
            c = ws.cell(curr_row, 1); c.value = rm; c.font = Font(name=FONT_MAIN, size=16, color=color)

        # (5) Signature - Parallel to Remarks
        # Shift down 1 row relative to Remarks start
        sig_start = curr_row - len(remarks_list)
        
        sig_col_start = max(1, total_cols - 8)
        
        ws.cell(sig_start, sig_col_start).value = "乙      方："
        ws.cell(sig_start, sig_col_start).font = Font(name=FONT_MAIN, size=16) 
        
        # (2) Party B is Client (v111.23 Fix)
        ws.cell(sig_start+1, sig_col_start+1).value = client_name 
        ws.cell(sig_start+1, sig_col_start+1).font = Font(name=FONT_MAIN, size=16)
        
        ws.cell(sig_start+2, sig_col_start).value = "統一編號："
        ws.cell(sig_start+2, sig_col_start).font = Font(name=FONT_MAIN, size=16)
        # (2) Tax ID Blank (v111.23 Fix)
        ws.cell(sig_start+2, sig_col_start+2).value = "" 
        ws.cell(sig_start+2, sig_col_start+2).font = Font(name=FONT_MAIN, size=16)
        
        ws.cell(sig_start+3, sig_col_start).value = "客戶簽章："
        ws.cell(sig_start+3, sig_col_start).font = Font(name=FONT_MAIN, size=16)

        # (3) Double Border below Remark 6 + 2 rows
        target_border_row = curr_row + 2
        for c_idx in range(1, total_cols + 1):
            ws.cell(target_border_row, c_idx).border = Border(bottom=SIDE_DOUBLE)

        return target_border_row

    # -------------------------------------------------------------
    # Render Logic: Bolin (v111.25 Explicit Args + Tweaks)
    # -------------------------------------------------------------
    # NOTE: Added explicit arguments here to solve NameError
    def render_bolin_optimized(ws, start_dt, end_dt, rows, budget, prod, client_name, product_name, remarks_list, logo_bytes=None):
        SIDE_DOUBLE = Side(style='double')
        if logo_bytes is None:
            logo_bytes = get_cloud_logo_bytes() # Auto fetch
        
        eff_days = (end_dt - start_dt).days + 1
        end_c_start = 6 + eff_days
        total_cols = end_c_start + 2

        ws.column_dimensions['A'].width = 21.0
        ws.column_dimensions['B'].width = 21.0
        ws.column_dimensions['C'].width = 13.8; ws.column_dimensions['D'].width = 19.4; ws.column_dimensions['E'].width = 15.0
        for i in range(eff_days): ws.column_dimensions[get_column_letter(6 + i)].width = 8.1
        ws.column_dimensions[get_column_letter(end_c_start)].width = 9.5
        ws.column_dimensions[get_column_letter(end_c_start+1)].width = 36.0 # [TWEAK 2]: Changed 58.0 to 36.0
        ws.column_dimensions[get_column_letter(end_c_start+2)].width = 20.0
        
        ROW_H_MAP = {1:70, 2:33.5, 3:33.5, 4:46, 5:40, 6:35, 7:35}
        for r, h in ROW_H_MAP.items(): ws.row_dimensions[r].height = h
        
        ws.merge_cells(f"A1:{get_column_letter(total_cols)}1"); c1 = ws['A1']
        c1.value = "鉑霖行動行銷-媒體計劃排程表 Mobi Media Schedule"; c1.font = Font(name=FONT_MAIN, size=28, bold=True); c1.alignment = ALIGN_LEFT 
        
        # (1) Logo (v111.23)
        if logo_bytes:
            try:
                img = openpyxl.drawing.image.Image(io.BytesIO(logo_bytes))
                scale = 130 / img.height
                img.height = 130
                img.width = int(img.width * scale)
                
                col_letter = get_column_letter(total_cols - 4) # [TWEAK 3]: Move left ~4 columns (approx)
                img.anchor = f"{col_letter}1" 
                ws.add_image(img)
            except Exception: pass

        c2a = ws['A2']; c2a.value = "TO："; c2a.font = Font(name=FONT_MAIN, size=20, bold=True, color="FF0000"); c2a.alignment = ALIGN_LEFT
        ws.merge_cells(f"B2:{get_column_letter(total_cols)}2"); c2b = ws['B2']; c2b.value = client_name; c2b.font = Font(name=FONT_MAIN, size=20, bold=True, color="FF0000"); c2b.alignment = ALIGN_LEFT
        
        c3a = ws['A3']; c3a.value = "FROM："; c3a.font = Font(name=FONT_MAIN, size=20, bold=True); c3a.alignment = ALIGN_LEFT
        ws.merge_cells(f"B3:{get_column_letter(total_cols)}3"); c3b = ws['B3']; c3b.value = "鉑霖行動行銷 許雅婷 TINA"; c3b.font = Font(name=FONT_MAIN, size=20, bold=True); c3b.alignment = ALIGN_LEFT

        unique_secs = sorted(list(set([r['seconds'] for r in rows]))); sec_str = " ".join([f"{s}秒廣告" for s in unique_secs])
        period_str = f"執行期間：{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')}"
        
        c4a = ws['A4']; c4a.value = "客戶名稱："; c4a.font = Font(name=FONT_MAIN, size=14, bold=True); c4a.alignment = ALIGN_LEFT
        ws.merge_cells("B4:E4"); c4b = ws['B4']; c4b.value = client_name; c4b.font = Font(name=FONT_MAIN, size=14, bold=True); c4b.alignment = ALIGN_LEFT
        spec_merge_start = "F4"; spec_merge_end = f"{get_column_letter(end_c_start)}4"
        ws.merge_cells(f"{spec_merge_start}:{spec_merge_end}")
        c4f = ws['F4']; c4f.value = f"廣告規格：{sec_str}"; c4f.font = Font(name=FONT_MAIN, size=14, bold=True); c4f.alignment = ALIGN_LEFT
        ws.merge_cells(f"{get_column_letter(end_c_start+1)}4:{get_column_letter(total_cols)}4")
        c4_r = ws[f"{get_column_letter(end_c_start+1)}4"]; c4_r.value = period_str; c4_r.font = Font(name=FONT_MAIN, size=14, bold=True); c4_r.alignment = ALIGN_LEFT
        draw_outer_border_fast(ws, 4, 4, 1, total_cols)

        c5a = ws['A5']; c5a.value = "廣告名稱："; c5a.font = Font(name=FONT_MAIN, size=14, bold=True); c5a.alignment = ALIGN_LEFT
        ws.merge_cells("B5:E5"); c5b = ws['B5']; c5b.value = product_name; c5b.font = Font(name=FONT_MAIN, size=14, bold=True); c5b.alignment = ALIGN_LEFT

        month_groups = []
        for i in range(eff_days):
            d = start_dt + timedelta(days=i); m_key = (d.year, d.month)
            if not month_groups or month_groups[-1][0] != m_key: month_groups.append([m_key, i, i]) 
            else: month_groups[-1][2] = i
        
        for m_key, s_idx, e_idx in month_groups:
            start_col = 6 + s_idx; end_col = 6 + e_idx
            ws.merge_cells(start_row=5, start_column=start_col, end_row=5, end_column=end_col)
            c = ws.cell(5, start_col); c.value = f"{m_key[1]}月"; c.font = FONT_BOLD; c.alignment = ALIGN_LEFT 

        for c_idx in range(1, total_cols + 1):
            c = ws.cell(5, c_idx)
            t, b, l, r = BS_MEDIUM, BS_MEDIUM, None, None
            if c_idx == 1: l = BS_MEDIUM 
            if c_idx == total_cols: r = BS_MEDIUM 
            if c_idx == 6: l = None 
            
            # [TWEAK 1]: Cancel Right Border for E5 (Col 5)
            if c_idx == 5: r = None

            c.border = Border(top=Side(style=t), bottom=Side(style=b), left=Side(style=l) if l else None, right=Side(style=r) if r else None)

        draw_outer_border_fast(ws, 5, 5, 1, 5) 

        header_start_row = 6
        headers = ["頻道", "播出地區", "播出店數", "播出時間", "秒數\n規格"]
        for i, h in enumerate(headers):
            c_idx = i + 1
            ws.merge_cells(start_row=header_start_row, start_column=c_idx, end_row=header_start_row+1, end_column=c_idx)
            c = ws.cell(header_start_row, c_idx); c.value = h; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER
            t, b, l, r = BS_MEDIUM, BS_THIN, BS_THIN, BS_THIN
            if c_idx == 1: l = BS_MEDIUM
            c.border = Border(top=Side(style=t), bottom=Side(style=b), left=Side(style=l), right=Side(style=r))
            ws.cell(header_start_row+1, c_idx).border = Border(top=Side(style=BS_THIN), bottom=Side(style=BS_THIN), left=Side(style=l), right=Side(style=r))

        curr = start_dt
        for i in range(eff_days):
            col_idx = 6 + i
            c6 = ws.cell(header_start_row, col_idx); c6.value = curr.day; c6.font = FONT_BOLD; c6.alignment = ALIGN_CENTER; c6.border = BORDER_ALL_MEDIUM
            c6.border = Border(top=Side(style=BS_MEDIUM), bottom=Side(style=BS_THIN), left=Side(style=BS_THIN), right=Side(style=BS_THIN))
            c7 = ws.cell(header_start_row+1, col_idx); c7.value = ["日","一","二","三","四","五","六"][(curr.weekday()+1)%7]
            c7.font = FONT_BOLD; c7.alignment = ALIGN_CENTER
            c7.border = Border(top=Side(style=BS_THIN), bottom=Side(style=BS_THIN), left=Side(style=BS_THIN), right=Side(style=BS_THIN))
            if curr.weekday() >= 5: c7.fill = FILL_WEEKEND
            curr += timedelta(days=1)

        end_headers = ["檔次", "定價", "專案價"]
        for i, h in enumerate(end_headers):
            c_idx = end_c_start + i
            ws.merge_cells(start_row=header_start_row, start_column=c_idx, end_row=header_start_row+1, end_column=c_idx)
            c = ws.cell(header_start_row, c_idx); c.value = h; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER
            t, b, l, r = BS_MEDIUM, BS_THIN, BS_THIN, BS_THIN
            if c_idx == total_cols: r = BS_MEDIUM
            c.border = Border(top=Side(style=t), bottom=Side(style=b), left=Side(style=l), right=Side(style=r))
            ws.cell(header_start_row+1, c_idx).border = Border(top=Side(style=BS_THIN), bottom=Side(style=BS_THIN), left=Side(style=l), right=Side(style=r))

        date_start_col = 6
        for c_idx in range(date_start_col, total_cols + 1):
            c7 = ws.cell(header_start_row, c_idx)
            c7.border = Border(top=Side(style=BS_MEDIUM), bottom=Side(style=BS_THIN), left=Side(style=BS_THIN), right=Side(style=BS_THIN))
            if c_idx == date_start_col: set_border(c7, left=BS_MEDIUM)
            if c_idx == total_cols: set_border(c7, right=BS_MEDIUM)
            c8 = ws.cell(header_start_row+1, c_idx)
            c8.border = Border(top=Side(style=BS_THIN), bottom=Side(style=BS_THIN), left=Side(style=BS_THIN), right=Side(style=BS_THIN))
            if c_idx == date_start_col: set_border(c8, left=BS_MEDIUM)
            if c_idx == total_cols: set_border(c8, right=BS_MEDIUM)

        # 4. Data Rows
        curr_row = header_start_row + 2
        grouped_data = {"全家廣播": sorted([r for r in rows if r["media"]=="全家廣播"], key=lambda x:x['seconds']),
                        "新鮮視": sorted([r for r in rows if r["media"]=="新鮮視"], key=lambda x:x['seconds']),
                        "家樂福": sorted([r for r in rows if r["media"]=="家樂福"], key=lambda x:x['seconds'])}
        
        total_store_count = 0; total_list_sum = 0
        for m_key, data in grouped_data.items():
            if not data: continue
            start_merge = curr_row
            d_name = f"全家便利商店\n{m_key}廣告" if m_key != "家樂福" else "家樂福"
            for idx, r in enumerate(data):
                ws.row_dimensions[curr_row].height = 40
                ws.cell(curr_row, 1, d_name).alignment = ALIGN_CENTER
                ws.cell(curr_row, 2, r['region']).alignment = ALIGN_CENTER
                p_num = int(r.get('program_num', 0)); total_store_count += p_num 
                suffix = "面" if m_key == "新鮮視" else "店"
                ws.cell(curr_row, 3, f"{p_num:,}{suffix}").alignment = ALIGN_CENTER
                ws.cell(curr_row, 4, r['daypart']).alignment = ALIGN_CENTER
                sec = r['seconds']
                if m_key == "新鮮視": sec_txt = f"{sec}秒\n影片/影像 1920x1080 (mp4)"
                else: sec_txt = f"{sec}秒廣告"
                c_spec = ws.cell(curr_row, 5, sec_txt); c_spec.alignment = ALIGN_CENTER; c_spec.font = Font(name=FONT_MAIN, size=10)
                row_sum = 0
                for d_idx in range(eff_days):
                    if d_idx < len(r['schedule']):
                        val = r['schedule'][d_idx]; row_sum += val
                        c = ws.cell(curr_row, 6+d_idx); c.value = val; c.alignment = ALIGN_CENTER; c.font = FONT_STD; c.border = BORDER_ALL_THIN
                ws.cell(curr_row, end_c_start, row_sum).alignment = ALIGN_CENTER
                rate_val = r['rate_display']
                if isinstance(rate_val, (int, float)): total_list_sum += rate_val
                ws.cell(curr_row, end_c_start+1, rate_val).number_format = FMT_MONEY; ws.cell(curr_row, end_c_start+1).alignment = ALIGN_CENTER 
                pkg = r['pkg_display']
                if r.get('is_pkg_member'): pkg = r['nat_pkg_display'] if idx == 0 else None
                ws.cell(curr_row, end_c_start+2, pkg).number_format = FMT_MONEY; ws.cell(curr_row, end_c_start+2).alignment = ALIGN_CENTER
                for c_idx in range(1, total_cols + 1):
                    c = ws.cell(curr_row, c_idx); c.font = FONT_STD; c.border = BORDER_ALL_THIN
                set_border(ws.cell(curr_row, 5), right=BS_MEDIUM)
                curr_row += 1
            ws.merge_cells(start_row=start_merge, start_column=1, end_row=curr_row-1, end_column=1)
            if data[0].get('is_pkg_member'): ws.merge_cells(start_row=start_merge, start_column=end_c_start+2, end_row=curr_row-1, end_column=end_c_start+2)
            draw_outer_border_fast(ws, start_merge, curr_row-1, 1, total_cols)

        # Total Row
        ws.row_dimensions[curr_row].height = 40
        ws.cell(curr_row, 3, total_store_count).number_format = FMT_NUMBER; ws.cell(curr_row, 3).alignment = ALIGN_CENTER; ws.cell(curr_row, 3).font = FONT_BOLD
        ws.cell(curr_row, 5, "Total").alignment = ALIGN_CENTER; ws.cell(curr_row, 5).font = FONT_BOLD
        for d_idx in range(eff_days):
            daily_sum = sum([r['schedule'][d_idx] for r in rows if d_idx < len(r['schedule'])])
            c = ws.cell(curr_row, 6+d_idx); c.value = daily_sum; c.alignment = ALIGN_CENTER; c.font = FONT_BOLD
        ws.cell(curr_row, end_c_start, sum([sum(r['schedule']) for r in rows])).alignment = ALIGN_CENTER; ws.cell(curr_row, end_c_start).font = FONT_BOLD
        ws.cell(curr_row, end_c_start+1, total_list_sum).number_format = FMT_MONEY; ws.cell(curr_row, end_c_start+1).font = FONT_BOLD; ws.cell(curr_row, end_c_start+1).alignment = ALIGN_CENTER
        ws.cell(curr_row, end_c_start+2, budget).number_format = FMT_MONEY; ws.cell(curr_row, end_c_start+2).font = FONT_BOLD; ws.cell(curr_row, end_c_start+2).alignment = ALIGN_CENTER
        for c_idx in range(1, total_cols+1): ws.cell(curr_row, c_idx).border = BORDER_ALL_THIN
        draw_outer_border_fast(ws, curr_row, curr_row, 1, total_cols)
        for c_idx in range(1, total_cols+1): set_border(ws.cell(curr_row, c_idx), bottom=BS_MEDIUM)
        set_border(ws.cell(curr_row, 5), right=BS_MEDIUM)
        curr_row += 1

        # Footer
        vat = int(budget * 0.05); grand_total = budget + vat
        footer_stack = [("製作", prod), ("5% VAT", vat), ("Grand Total", grand_total)]
        for lbl, val in footer_stack:
            ws.row_dimensions[curr_row].height = 30
            c_l = ws.cell(curr_row, end_c_start+1); c_l.value = lbl; c_l.alignment = ALIGN_RIGHT; c_l.font = FONT_STD
            c_v = ws.cell(curr_row, end_c_start+2); c_v.value = val; c_v.number_format = FMT_MONEY; c_v.alignment = ALIGN_CENTER; c_v.font = FONT_BOLD 
            t, b, l, r = BS_THIN, BS_THIN, BS_MEDIUM, BS_THIN
            if lbl == "Grand Total": b = BS_MEDIUM 
            c_l.border = Border(top=Side(style=t), bottom=Side(style=b), left=Side(style=l), right=Side(style=r))
            t, b, l, r = BS_THIN, BS_THIN, BS_THIN, BS_MEDIUM
            if lbl == "Grand Total": b = BS_MEDIUM 
            c_v.border = Border(top=Side(style=t), bottom=Side(style=b), left=Side(style=l), right=Side(style=r))
            if lbl == "Grand Total":
                for c_idx in range(1, total_cols + 1): set_border(ws.cell(curr_row, c_idx), bottom=BS_MEDIUM)
            curr_row += 1
        
        curr_row += 1
        start_footer = curr_row
        
        r_col_start = 6 
        ws.cell(start_footer, r_col_start).value = "Remarks："
        ws.cell(start_footer, r_col_start).font = Font(name=FONT_MAIN, size=16, bold=True)
        r_row = start_footer
        for rm in remarks_list:
            r_row += 1
            color = "000000"
            if rm.strip().startswith("1.") or rm.strip().startswith("4."): color = "FF0000"
            if rm.strip().startswith("6."): color = "0000FF"
            c = ws.cell(r_row, r_col_start); c.value = rm; c.font = Font(name=FONT_MAIN, size=16, color=color)

        sig_col_start = 1
        ws.cell(start_footer, sig_col_start).value = "乙      方："
        ws.cell(start_footer, sig_col_start).font = Font(name=FONT_MAIN, size=16)
        
        # (2) Party B is Client (v111.23 Fix)
        ws.cell(start_footer+1, sig_col_start+1).value = client_name 
        ws.cell(start_footer+1, sig_col_start+1).font = Font(name=FONT_MAIN, size=16)
        
        ws.cell(start_footer+2, sig_col_start).value = "統一編號："
        ws.cell(start_footer+2, sig_col_start).font = Font(name=FONT_MAIN, size=16)
        # (2) Tax ID Blank (v111.23 Fix)
        ws.cell(start_footer+2, sig_col_start+2).value = "" 
        ws.cell(start_footer+2, sig_col_start+2).font = Font(name=FONT_MAIN, size=16)
        
        ws.cell(start_footer+3, sig_col_start).value = "客戶簽章："
        ws.cell(start_footer+3, sig_col_start).font = Font(name=FONT_MAIN, size=16)

        # (3) Double Border below Remark 6 + 2 rows
        target_border_row = r_row + 2
        for c_idx in range(1, total_cols + 1):
            ws.cell(target_border_row, c_idx).border = Border(bottom=SIDE_DOUBLE)

        return target_border_row

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Schedule"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE; ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.fitToPage = True
    
    if format_type == "Dongwu": render_dongwu_optimized(ws, start_dt, end_dt, rows, final_budget_val, prod_cost)
    elif format_type == "Shenghuo": render_shenghuo_optimized(ws, start_dt, end_dt, rows, final_budget_val, prod_cost)
    else: render_bolin_optimized(ws, start_dt, end_dt, rows, final_budget_val, prod_cost, client_name, product_name, remarks_list) # [MODIFIED]: Added explicit arguments

    out = io.BytesIO(); wb.save(out); return out.getvalue()

if __name__ == "__main__":
    main()
