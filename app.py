import streamlit as st
import traceback
import time
import gc
from itertools import groupby

# =========================================================
# 1. 頁面設定
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v109.0 (Direct)")

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
        "2.以上節目名稱如有異動，以上檔時節目名稱為主，如遇時段滿檔，上檔時間挪後或更換至同級時段。",
        "3.通路店鋪數與開機率至少七成(以上)。每日因加盟數調整，或遇店舖年度季度改裝、設備維護升級及保修等狀況，會有一定幅度增減。",
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

# [快取優化] 加入 Cache，避免重複運算
@st.cache_data(show_spinner="正在生成 PDF (LibreOffice)...", ttl=3600)
def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    soffice = find_soffice_path()
    if not soffice: 
        return None, "Fail", "伺服器未安裝 LibreOffice"
    try:
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, "cue.xlsx")
            with open(xlsx_path, "wb") as f: f.write(xlsx_bytes)
            
            subprocess.run(
                [soffice, "--headless", "--nologo", "--convert-to", "pdf:calc_pdf_Export", "--outdir", tmp, xlsx_path], 
                capture_output=True, 
                timeout=60
            )
            
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
        g_list = list(group)
        g_size = len(g_list)
        is_pkg = g_list[0]['is_pkg_member']
        
        for i, r in enumerate(g_list):
            tbody += "<tr>"
            rate = f"${r['rate_display']:,}" if isinstance(r['rate_display'], (int, float)) else r['rate_display']
            pkg_val_str = ""
            if is_pkg:
                if i == 0:
                    val = f"${r['nat_pkg_display']:,}"; pkg_val_str = f"<td class='right' rowspan='{g_size}'>{val}</td>"
            else:
                val = f"${r['pkg_display']:,}" if isinstance(r['pkg_display'], (int, float)) else r['pkg_display']; pkg_val_str = f"<td class='right'>{val}</td>"

            if format_type == "Shenghuo": 
                sec_txt = f"{r['seconds']}秒"; tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{sec_txt}</td><td>{rate}</td>{pkg_val_str}"
            elif format_type == "Bolin": 
                tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{r['seconds']}秒</td><td>{rate}</td>{pkg_val_str}"
            else: 
                tbody += f"<td>{r['media']}</td><td>{r['region']}</td><td>{r.get('program_num','')}</td><td>{r['daypart']}</td><td>{r['seconds']}</td><td>{rate}</td>{pkg_val_str}"
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

# [快取優化] 加入 Cache，提升重複點擊時的速度
@st.cache_data(show_spinner="正在生成 Excel 報表...", ttl=3600)
def generate_excel_from_scratch(format_type, start_dt, end_dt, client_name, product_name, rows, remarks_list, final_budget_val, prod_cost):
    import openpyxl
    from openpyxl.utils import get_column_letter, column_index_from_string
    from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

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
        COL_WIDTHS = {'A': 19.6, 'B': 22.8, 'C': 14.6, 'D': 20.0, 'E': 13.0, 'F': 19.6, 'G': 17.9}
        ROW_HEIGHTS = {1: 61.0, 2: 29.0, 3: 40.0, 4: 40.0, 5: 40.0, 6: 40.0, 7: 40.0, 8: 40.0}
        for k, v in COL_WIDTHS.items(): ws.column_dimensions[k].width = v
        for i in range(8, 40): ws.column_dimensions[get_column_letter(i)].width = 8.5
        ws.column_dimensions['AM'].width = 13.0
        for r, h in ROW_HEIGHTS.items(): ws.row_dimensions[r].height = h

        ws.merge_cells("A1:AM1"); c = ws['A1']; c.value = "Media Schedule"; c.font = FONT_TITLE; c.alignment = ALIGN_CENTER
        unique_media = sorted(list(set([r['media'] for r in rows]))); order = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}; unique_media.sort(key=lambda x: order.get(x, 99)); medium_str = "/".join(unique_media)
        unique_secs = sorted(list(set([r['seconds'] for r in rows]))); p_str = f"{'、'.join([f'{s}秒' for s in unique_secs])} {product_name}"
        infos = [("A3", "客戶名稱：", client_name), ("A4", "Product：", p_str), 
                 ("A5", "Period :", f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"), ("A6", "Medium :", medium_str)]
        for pos, lbl, val in infos:
            c = ws[pos]; c.value = lbl; c.font = FONT_BOLD; c.alignment = Alignment(vertical='center')
            c2 = ws.cell(c.row, 2); c2.value = val; c2.font = FONT_BOLD; c2.alignment = Alignment(vertical='center')

        ws['H6'] = f"{start_dt.month}月"; ws['H6'].font = Font(name=FONT_MAIN, size=16, bold=True); ws['H6'].alignment = ALIGN_CENTER
        headers = [("A","Station"), ("B","Location"), ("C","Program"), ("D","Day-part"), ("E","Size"), ("F","rate\n(Net)"), ("G","Package-cost\n(Net)")]
        for col, txt in headers:
            ws[f"{col}7"] = txt; ws.merge_cells(f"{col}7:{col}8"); c = ws[f"{col}7"]; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER; c.border = BORDER_ALL_MEDIUM

        eff_days = (end_dt - start_dt).days + 1; curr = start_dt
        for i in range(31):
            col_idx = 8 + i; c_d = ws.cell(7, col_idx); c_w = ws.cell(8, col_idx)
            if i < eff_days:
                c_d.value = curr; c_d.number_format = 'm/d'; c_w.value = ["一","二","三","四","五","六","日"][curr.weekday()]
                if curr.weekday() >= 5: c_w.fill = FILL_WEEKEND
                curr += timedelta(days=1)
            c_d.font = FONT_STD; c_w.font = FONT_STD; c_d.alignment = ALIGN_CENTER; c_w.alignment = ALIGN_CENTER; c_d.border = BORDER_ALL_THIN; c_w.border = BORDER_ALL_THIN

        ws['AM7'] = "檔次"; ws.merge_cells("AM7:AM8"); ws['AM7'].font = FONT_BOLD; ws['AM7'].alignment = ALIGN_CENTER; ws['AM7'].border = BORDER_ALL_MEDIUM

        curr_row = 9; grouped_data = {
            "全家廣播": sorted([r for r in rows if r["media"] == "全家廣播"], key=lambda x: x["seconds"]),
            "新鮮視": sorted([r for r in rows if r["media"] == "新鮮視"], key=lambda x: x["seconds"]),
            "家樂福": sorted([r for r in rows if r["media"] == "家樂福"], key=lambda x: x["seconds"]),
        }

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
                if r.get("is_pkg_member"): pkg = r['nat_pkg_display'] if idx == 0 else None
                c_rate = ws.cell(curr_row, 6); c_rate.value = rate; c_rate.number_format = FMT_MONEY; c_rate.alignment = ALIGN_CENTER
                if pkg is not None:
                    c_pkg = ws.cell(curr_row, 7); c_pkg.value = pkg; c_pkg.number_format = FMT_MONEY; c_pkg.alignment = ALIGN_CENTER

                row_sum = 0
                for d_idx in range(eff_days):
                    if d_idx < len(r["schedule"]):
                        val = r["schedule"][d_idx]; row_sum += val
                        c_s = ws.cell(curr_row, 8+d_idx); c_s.value = val; c_s.number_format = FMT_NUMBER; c_s.alignment = ALIGN_CENTER
                
                ws.cell(curr_row, 39, row_sum).alignment = ALIGN_CENTER
                for c_idx in range(1, 40):
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
            draw_outer_border_fast(ws, start_merge, curr_row-1, 1, 39)

        ws.row_dimensions[curr_row].height = 30
        c_lbl = ws.cell(curr_row, 6, "Grand Total"); c_lbl.alignment = ALIGN_RIGHT; c_lbl.font = FONT_STD
        vat = int(budget * 0.05)
        c_val = ws.cell(curr_row, 7, budget + vat); c_val.number_format = FMT_MONEY; c_val.alignment = ALIGN_CENTER; c_val.font = FONT_STD
        total_spots = sum([sum(r['schedule']) for r in rows]); ws.cell(curr_row, 39, total_spots).alignment = ALIGN_CENTER
        draw_outer_border_fast(ws, curr_row, curr_row, 1, 39)
        
        curr_row += 2
        ws.cell(curr_row, 1, "Remarks:").font = Font(name=FONT_MAIN, size=16, bold=True, underline='single')
        for rm in remarks_list:
            curr_row += 1
            c = ws.cell(curr_row, 1); c.value = rm; c.font = Font(name=FONT_MAIN, size=14, color="FF0000" if rm.startswith("1") else "000000")
        return curr_row

    # -------------------------------------------------------------
    # Render Logic: Shenghuo
    # -------------------------------------------------------------
    def render_shenghuo_optimized(ws, start_dt, end_dt, rows, budget, prod):
        days_n = (end_dt - start_dt).days + 1
        ws.column_dimensions['A'].width = 22.5; ws.column_dimensions['B'].width = 24.5; ws.column_dimensions['C'].width = 13.8; ws.column_dimensions['D'].width = 19.4; ws.column_dimensions['E'].width = 13.0
        for i in range(days_n): ws.column_dimensions[get_column_letter(6 + i)].width = 13.0
        end_c_start = 6 + days_n; ws.column_dimensions[get_column_letter(end_c_start)].width = 13.0; ws.column_dimensions[get_column_letter(end_c_start+1)].width = 59.0; ws.column_dimensions[get_column_letter(end_c_start+2)].width = 13.2 
        total_cols = 5 + days_n + 3
        
        ROW_H_MAP = {1:46, 2:46, 3:46, 4:46.5, 5:40, 6:40, 7:40, 8:40}
        for r, h in ROW_H_MAP.items(): ws.row_dimensions[r].height = h
        
        ws['A3'] = "聲活數位科技股份有限公司 統編 28710100"; ws['A3'].font = Font(name=FONT_MAIN, size=20); ws['A3'].alignment = Alignment(vertical='center')
        ws['A4'] = "蔡伊閔"; ws['A4'].font = Font(name=FONT_MAIN, size=16); ws['A4'].alignment = Alignment(vertical='center')
        unique_secs = sorted(list(set([r['seconds'] for r in rows]))); sec_str = " ".join([f"{s}秒廣告" for s in unique_secs])
        ws['A5'] = "客戶名稱："; ws['B5'] = client_name; ws['F5'] = "廣告規格："; ws['H5'] = sec_str
        ws['A6'] = "廣告名稱："; ws['B6'] = product_name
        
        headers = ["頻道", "播出地區", "播出店數", "播出時間", "秒數\n規格"]
        for i, h in enumerate(headers):
            ws.merge_cells(start_row=7, start_column=i+1, end_row=8, end_column=i+1); c = ws.cell(7, i+1); c.value = h; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER; c.border = BORDER_ALL_MEDIUM

        curr = start_dt
        for i in range(days_n):
            c = 6 + i
            c7 = ws.cell(7, c); c7.value = curr; c7.number_format = 'd'; c7.font = FONT_BOLD; c7.alignment = ALIGN_CENTER; c7.border = BORDER_ALL_MEDIUM
            c8 = ws.cell(8, c); c8.value = ["日","一","二","三","四","五","六"][(curr.weekday()+1)%7]; c8.font = FONT_BOLD; c8.alignment = ALIGN_CENTER; c8.border = BORDER_ALL_MEDIUM
            if curr.weekday() >= 5: c8.fill = FILL_WEEKEND
            curr += timedelta(days=1)

        end_headers = ["檔次", "定價", "專案價"]
        for i, h in enumerate(end_headers):
            c = end_c_start + i; ws.merge_cells(start_row=7, start_column=c, end_row=8, end_column=c); cell = ws.cell(7, c); cell.value = h; cell.font = FONT_BOLD; cell.alignment = ALIGN_CENTER; cell.border = BORDER_ALL_MEDIUM

        curr_row = 9
        grouped_data = {"全家廣播": sorted([r for r in rows if r["media"]=="全家廣播"], key=lambda x:x['seconds']),
                        "新鮮視": sorted([r for r in rows if r["media"]=="新鮮視"], key=lambda x:x['seconds']),
                        "家樂福": sorted([r for r in rows if r["media"]=="家樂福"], key=lambda x:x['seconds'])}
        
        for m_key, data in grouped_data.items():
            if not data: continue
            start_merge = curr_row
            d_name = f"全家便利商店\n{m_key}廣告" if m_key != "家樂福" else "家樂福"
            
            for idx, r in enumerate(data):
                ws.row_dimensions[curr_row].height = 40
                ws.cell(curr_row, 1, d_name).alignment = ALIGN_CENTER
                ws.cell(curr_row, 2, r['region']).alignment = ALIGN_CENTER
                ws.cell(curr_row, 3, r.get('program_num',0)).alignment = ALIGN_CENTER
                ws.cell(curr_row, 4, r['daypart']).alignment = ALIGN_CENTER
                ws.cell(curr_row, 5, f"{r['seconds']}秒").alignment = ALIGN_CENTER
                row_sum = 0
                for d_idx in range(days_n):
                    if d_idx < len(r['schedule']):
                        val = r['schedule'][d_idx]; row_sum += val
                        c = ws.cell(curr_row, 6+d_idx); c.value = val; c.alignment = ALIGN_CENTER
                ws.cell(curr_row, end_c_start, row_sum).alignment = ALIGN_CENTER
                rate = r['rate_display']; pkg = r['pkg_display']
                if r.get('is_pkg_member'): pkg = r['nat_pkg_display'] if idx == 0 else None
                ws.cell(curr_row, end_c_start+1, rate).number_format = FMT_MONEY
                if pkg is not None: ws.cell(curr_row, end_c_start+2, pkg).number_format = FMT_MONEY
                for c_idx in range(1, total_cols + 1):
                    c = ws.cell(curr_row, c_idx); c.font = FONT_STD; c.border = BORDER_ALL_THIN
                curr_row += 1
            ws.merge_cells(start_row=start_merge, start_column=1, end_row=curr_row-1, end_column=1)
            if data[0].get('is_pkg_member'): ws.merge_cells(start_row=start_merge, start_column=end_c_start+2, end_row=curr_row-1, end_column=end_c_start+2)
            draw_outer_border_fast(ws, start_merge, curr_row-1, 1, total_cols)

        ws.row_dimensions[curr_row].height = 40
        ws.cell(curr_row, end_c_start+1, "Total").alignment = ALIGN_RIGHT
        ws.cell(curr_row, end_c_start+2, budget + int(budget*0.05)).number_format = FMT_MONEY
        draw_outer_border_fast(ws, curr_row, curr_row, 1, total_cols)
        curr_row += 2
        ws.cell(curr_row, 1, "Remarks:").font = FONT_BOLD
        for rm in remarks_list:
            curr_row += 1; ws.cell(curr_row, 1, rm).font = FONT_STD

    # -------------------------------------------------------------
    # Render Logic: Bolin
    # -------------------------------------------------------------
    def render_bolin_optimized(ws, start_dt, end_dt, rows, budget, prod):
        days_n = (end_dt - start_dt).days + 1; total_cols = 1 + 5 + days_n + 3
        ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 20
        for i in range(days_n): ws.column_dimensions[get_column_letter(7+i)].width = 5
        ws['B2']="TO："; ws['C2']=client_name; ws['B3']="FROM："; ws['C3']="鉑霖行動行銷 許雅婷 TINA"
        unique_secs = sorted(list(set([r['seconds'] for r in rows]))); sec_str = " ".join([f"{s}秒廣告" for s in unique_secs])
        ws['G4']="廣告規格："; ws['H4']=sec_str
        
        header_fill = FILL_HEADER_BOLIN
        headers = ["頻道", "播出地區", "播出店數", "播出時間", "規格"]
        for i, h in enumerate(headers):
            c = ws.cell(7, 2+i); c.value = h; c.fill = header_fill; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER; c.border = BORDER_ALL_MEDIUM
        curr = start_dt
        for i in range(days_n):
            c = ws.cell(7, 7+i); c.value = curr; c.number_format = 'm/d'; c.fill = header_fill; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER; c.border = BORDER_ALL_MEDIUM
            if curr.weekday() >= 5: c.fill = FILL_WEEKEND
            curr += timedelta(days=1)
        end_h = ["總檔次", "單價", "金額"]; end_c_start = 7 + days_n
        for i, h in enumerate(end_h):
            c = ws.cell(7, end_c_start+i); c.value = h; c.fill = header_fill; c.font = FONT_BOLD; c.alignment = ALIGN_CENTER; c.border = BORDER_ALL_MEDIUM

        curr_row = 8
        grouped_data = {"全家廣播": sorted([r for r in rows if r["media"]=="全家廣播"], key=lambda x:x['seconds']),
                        "新鮮視": sorted([r for r in rows if r["media"]=="新鮮視"], key=lambda x:x['seconds']),
                        "家樂福": sorted([r for r in rows if r["media"]=="家樂福"], key=lambda x:x['seconds'])}
        
        for m_key, data in grouped_data.items():
            if not data: continue
            start_merge = curr_row
            d_name = f"全家便利商店\n{m_key}" if m_key != "家樂福" else "家樂福"
            for idx, r in enumerate(data):
                ws.row_dimensions[curr_row].height = 25
                ws.cell(curr_row, 2, d_name).alignment = ALIGN_CENTER
                ws.cell(curr_row, 3, r['region']).alignment = ALIGN_CENTER
                ws.cell(curr_row, 4, r.get('program_num',0)).alignment = ALIGN_CENTER
                ws.cell(curr_row, 5, r['daypart']).alignment = ALIGN_CENTER
                ws.cell(curr_row, 6, f"{r['seconds']}秒").alignment = ALIGN_CENTER
                row_sum = 0
                for d_idx in range(days_n):
                    if d_idx < len(r['schedule']):
                        val = r['schedule'][d_idx]; row_sum += val
                        ws.cell(curr_row, 7+d_idx, val).alignment = ALIGN_CENTER
                ws.cell(curr_row, end_c_start, row_sum).alignment = ALIGN_CENTER
                rate = r['rate_display']; pkg = r['pkg_display']
                if r.get('is_pkg_member'): pkg = r['nat_pkg_display'] if idx == 0 else None
                ws.cell(curr_row, end_c_start+1, rate).number_format = FMT_MONEY
                if pkg is not None: ws.cell(curr_row, end_c_start+2, pkg).number_format = FMT_MONEY
                for c_idx in range(2, total_cols+1):
                    c = ws.cell(curr_row, c_idx); c.font = FONT_STD; c.border = BORDER_ALL_THIN
                curr_row += 1
            ws.merge_cells(start_row=start_merge, start_column=2, end_row=curr_row-1, end_column=2)
            if data[0].get('is_pkg_member'): ws.merge_cells(start_row=start_merge, start_column=end_c_start+2, end_row=curr_row-1, end_column=end_c_start+2)
            draw_outer_border_fast(ws, start_merge, curr_row-1, 2, total_cols)

        ws.row_dimensions[curr_row].height = 30
        ws.cell(curr_row, end_c_start+1, "Total").alignment = ALIGN_RIGHT
        ws.cell(curr_row, end_c_start+2, budget + int(budget*0.05)).number_format = FMT_MONEY
        draw_outer_border_fast(ws, curr_row, curr_row, 2, total_cols)
        curr_row += 2
        ws.cell(curr_row, 9, "Remarks:").font = FONT_BOLD
        for rm in remarks_list:
            curr_row += 1; ws.cell(curr_row, 9, rm).font = FONT_STD

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Schedule"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE; ws.page_setup.paperSize = ws.PAPERSIZE_A4; ws.page_setup.fitToPage = True
    
    if format_type == "Dongwu": render_dongwu_optimized(ws, start_dt, end_dt, rows, final_budget_val, prod_cost)
    elif format_type == "Shenghuo": render_shenghuo_optimized(ws, start_dt, end_dt, rows, final_budget_val, prod_cost)
    else: render_bolin_optimized(ws, start_dt, end_dt, rows, final_budget_val, prod_cost)

    out = io.BytesIO(); wb.save(out); return out.getvalue()

# =========================================================
# 10. Main Execution Block
# =========================================================
def main():
    try:
        with st.spinner("正在讀取 Google 試算表設定檔..."):
            STORE_COUNTS, STORE_COUNTS_NUM, PRICING_DB, SEC_FACTORS, err_msg = load_config_from_cloud(GSHEET_SHARE_URL)
        if err_msg:
            st.error(f"❌ 設定檔載入失敗: {err_msg}")
            st.stop()
        
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
            st.markdown("---")
            if st.button("🧹 清除快取"): st.cache_data.clear(); st.rerun()

        st.title("📺 媒體 Cue 表生成器 (v109.0 Direct)")
        format_type = st.radio("選擇格式", ["Dongwu", "Shenghuo", "Bolin"], horizontal=True)

        c1, c2, c3, c4, c5_sales = st.columns(5)
        with c1: client_name = st.text_input("客戶名稱", "萬國通路")
        with c2: product_name = st.text_input("產品名稱", "統一布丁")
        with c3: total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)
        with c4: prod_cost_input = st.number_input("製作費 (未稅)", value=0, step=1000)
        with c5_sales: sales_person = st.text_input("業務名稱", "")

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
        col_cb1, col_cb2, col_cb3 = st.columns(3)
        
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

        is_rad = col_cb1.checkbox("全家廣播", key="cb_rad", on_change=on_media_change)
        is_fv = col_cb2.checkbox("新鮮視", key="cb_fv", on_change=on_media_change)
        is_cf = col_cb3.checkbox("家樂福", key="cb_cf", on_change=on_media_change)

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
            rows, total_list_accum, logs = calculate_plan_data(config, total_budget_input, days_count, PRICING_DB, SEC_FACTORS, STORE_COUNTS_NUM, REGIONS_ORDER)
            prod_cost = prod_cost_input 
            vat = int(round(final_budget_val * 0.05))
            grand_total = final_budget_val + vat
            p_str = f"{'、'.join([f'{s}秒' for s in sorted(list(set(r['seconds'] for r in rows)))])} {product_name}"
            rem = get_remarks_text(sign_deadline, billing_month, payment_date)
            html_preview = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_str, format_type, rem, total_list_accum, grand_total, final_budget_val, prod_cost)
            
            st.components.v1.html(html_preview, height=700, scrolling=True)
            
            st.markdown("---")
            st.subheader("📥 檔案下載區")
            
            # --- 直覺式下載 (Logic v109) ---
            
            # 1. 預先準備 (Cache Hit = Instant)
            xlsx_temp = generate_excel_from_scratch(format_type, start_date, end_date, client_name, product_name, rows, rem, final_budget_val, prod_cost)
            
            col_dl1, col_dl2 = st.columns(2)
            
            # 2. PDF Download (Only calc when user clicks, or pre-calc if fast)
            # 為了避免每次小改動都要跑 LibreOffice (慢)，我們這裡用 "Lazy" 邏輯：
            # 下載按鈕直接連到 Cache 函數。只有第一次按 (或參數變更後第一次按) 會轉圈圈。
            with col_dl2:
                # 呼叫 Cache Function
                pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx_temp)
                if pdf_bytes:
                    st.download_button(
                        f"📥 下載 PDF", 
                        pdf_bytes, 
                        f"Cue_{safe_filename(client_name)}.pdf", 
                        key="pdf_dl_btn",
                        mime="application/pdf"
                    )
                else:
                    st.warning(f"PDF 生成失敗: {err}")

            # 3. Excel Download (Fast)
            with col_dl1:
                if st.session_state.is_supervisor:
                    st.download_button(
                        "📥 下載 Excel (主管權限)", 
                        xlsx_temp, 
                        f"Cue_{safe_filename(client_name)}.xlsx", 
                        key="xlsx_dl_btn",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("🔒 Excel 下載功能僅限主管使用")

    except Exception as e:
        st.error("程式執行發生錯誤，請聯絡開發者。")
        st.error(traceback.format_exc())

if __name__ == "__main__":
    main()
