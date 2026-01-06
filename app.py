import streamlit as st
import traceback
import time
import gc
from itertools import groupby
import pandas as pd
import math
import io
import os
import shutil
import tempfile
import subprocess
import re
import requests
from datetime import timedelta, datetime, date
from copy import copy

# Excel 處理相關庫
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.drawing.image import Image as OpenpyxlImage

# =========================================================
# 1. 頁面設定
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro v112.1 (Ragic Connected)")

# =========================================================
# 2. Session State 初始化 (含 Ragic 預設值)
# =========================================================
# 您提供的 Ragic 資訊已預設在此
DEFAULT_RAGIC_URL = "https://ap15.ragic.com/liuskyo/cue/2" # 已移除 ?PAGEID 參數，API 不需要
DEFAULT_RAGIC_KEY = "L04zZGhrVmtTV3pqN1VLbUpnOFZMa01NTHh3OUw3RUVlb0ovNXUrTXJsaGJhMWpKOUxHanFUODREMmN1dEZvcw=="

DEFAULT_STATES = {
    "is_supervisor": False,
    "rad_share": 100, "fv_share": 0, "cf_share": 0,
    "cb_rad": True, "cb_fv": False, "cb_cf": False,
    "ragic_url": DEFAULT_RAGIC_URL,
    "ragic_key": DEFAULT_RAGIC_KEY,
    "ragic_confirm": False
}

for key, default_val in DEFAULT_STATES.items():
    if key not in st.session_state: st.session_state[key] = default_val

# =========================================================
# 3. 全域常數設定
# =========================================================
GSHEET_SHARE_URL = "https://docs.google.com/spreadsheets/d/1bzmG-N8XFsj8m3LUPqA8K70AcIqaK4Qhq1VPWcK0w_s/edit?usp=sharing"
BOLIN_LOGO_URL = "https://docs.google.com/drawings/d/17Uqgp-7LJJj9E4bV7Azo7TwXESPKTTIsmTbf-9tU9eE/export/png"
FONT_MAIN = "微軟正黑體"
BS_THIN, BS_MEDIUM, BS_HAIR = 'thin', 'medium', 'hair'
FMT_MONEY = '"$"#,##0_);[Red]("$"#,##0)'
FMT_NUMBER = '#,##0'
REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]
REGION_DISPLAY_MAP = {
    "北區": "北區-北北基", "桃竹苗": "桃區-桃竹苗", "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南", "高屏": "高屏區-高屏", "東區": "東區-宜花東",
    "全省量販": "全省量販", "全省超市": "全省超市"
}

# =========================================================
# 4. 基礎工具函式
# =========================================================
def parse_count_to_int(x):
    if x is None: return 0
    if isinstance(x, (int, float)): return int(x)
    m = re.findall(r"[\d,]+", str(x))
    return int(m[0].replace(",", "")) if m else 0

def safe_filename(name: str) -> str: return re.sub(r'[\\/*?:"<>|]', "_", name).strip()
def html_escape(s): return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") if s else ""
def region_display(region): return REGION_DISPLAY_MAP.get(region, region)

def get_sec_factor(media_type, seconds, sec_factors):
    factors = sec_factors.get(media_type) or sec_factors.get("全家新鮮視" if media_type=="新鮮視" else "全家廣播")
    if not factors: return 1.0
    if seconds in factors: return factors[seconds]
    for base in [10, 20, 15, 30]:
        if base in factors: return (seconds / base) * factors[base]
    return 1.0

def calculate_schedule(total_spots, days):
    if days <= 0: return []
    if total_spots % 2 != 0: total_spots += 1
    base, rem = divmod(total_spots // 2, days)
    return [(base + (1 if i < rem else 0)) * 2 for i in range(days)]

def get_remarks_text(sign_deadline, billing_month, payment_date):
    d_str = sign_deadline.strftime("%Y/%m/%d (%a)") if sign_deadline else "____/__/__ (__)"
    p_str = payment_date.strftime("%Y/%m/%d") if payment_date else "____/__/__"
    return [
        f"1.請於 {d_str} 11:30前 回簽及進單，方可順利上檔。",
        "2.以上節目名稱如有異動，以上檔時節目名稱為主，如遇電台時段滿檔，上檔時間挪後或更換至同級時段。",
        "3.通路店鋪數與開機率至少七成。每日因加盟數調整，會有一定幅度增減。",
        "4.託播方需於上檔前 5 個工作天，提供廣告帶(mp3)、影片/影像 1920x1080 (mp4)。",
        f"5.雙方同意費用請款月份 : {billing_month}，如有修正必要，將另行E-Mail告知。",
        f"6.付款兌現日期：{p_str}"
    ]

# Ragic 參數格式化工具
def format_campaign_details(config):
    details = []
    for media, settings in config.items():
        sec_str = ", ".join([f"{s}秒({p}%)" for s, p in settings.get("sec_shares", {}).items()])
        if settings.get("is_national"): reg_str = "全省聯播"
        else: reg_str = "/".join(settings.get("regions", []))
        info = f"【{media}】 預算佔比: {settings.get('share')}% | 秒數分配: {sec_str} | 區域: {reg_str}"
        details.append(info)
    return "\n".join(details)

def find_soffice_path():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice: return soffice
    if os.name == "nt": 
        for p in [r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
            if os.path.exists(p): return p
    return None

@st.cache_data(show_spinner="下載 Logo...", ttl=3600)
def get_cloud_logo_bytes():
    try:
        r = requests.get(BOLIN_LOGO_URL, timeout=5)
        return r.content if r.status_code == 200 else None
    except: return None

@st.cache_data(show_spinner="生成 PDF...", ttl=3600)
def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    soffice = find_soffice_path()
    if not soffice: return None, "Fail", "伺服器未安裝 LibreOffice"
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
            return None, "Fail", "未產出檔案"
    except Exception as e: return None, "Fail", str(e)
    finally: gc.collect()

# Ragic 上傳核心函式
def upload_to_ragic(api_url, api_key, data_dict, files_dict=None):
    if not api_url or not api_key: return False, "API URL 或 Key 未設定"
    # 確保 URL 包含 ?api 參數
    target_url = api_url if api_url.endswith("?api") else f"{api_url}?api"
    
    try:
        # Ragic 使用 Basic Auth (API Key 為帳號)
        resp = requests.post(target_url, auth=(api_key, ''), data=data_dict, files=files_dict, timeout=60)
        
        if resp.status_code == 200:
            rjson = resp.json()
            if rjson.get('status') == 'SUCCESS': 
                return True, f"✅ 上傳成功! Ragic ID: {rjson.get('ragicId')}"
            else: 
                return False, f"❌ Ragic 錯誤: {rjson.get('msg')}"
        return False, f"❌ HTTP 錯誤: {resp.status_code} - {resp.text}"
    except Exception as e: return False, f"❌ 連線異常: {str(e)}"

# =========================================================
# HTML 預覽 (簡化版，請用原版替換內容以求美觀)
# =========================================================
def generate_html_preview(rows, days_cnt, start_dt, end_dt, c_name, p_display, format_type, remarks, total_list, grand_total, budget, prod):
    # 這裡僅回傳簡單字串證明流程通順，實際專案請貼回您原本漂亮的 generate_html_preview 函式
    return f"""
    <html><body>
    <h3>預覽產生成功 (Preview Generated)</h3>
    <p><b>客戶:</b> {c_name} | <b>產品:</b> {p_display}</p>
    <p><b>總金額 (含稅):</b> ${grand_total:,}</p>
    </body></html>
    """

# =========================================================
# 6. Excel 渲染 (簡化版，請用原版替換)
# =========================================================
@st.cache_data(show_spinner="生成 Excel...", ttl=3600)
def generate_excel_from_scratch(format_type, start_dt, end_dt, client_name, product_name, rows, remarks_list, final_budget_val, prod_cost, sales_person):
    # 這裡僅生成有資料的 Excel 供上傳，實際請貼回您原本的 Excel 生成邏輯
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Schedule"
    ws['A1'] = f"Client: {client_name}"
    ws['A2'] = f"Product: {product_name}"
    ws['A3'] = f"Budget: {final_budget_val}"
    ws['A4'] = f"Sales: {sales_person}"
    
    # 填入一些資料列
    for i, r in enumerate(rows):
        ws.cell(6+i, 1, r['media'])
        ws.cell(6+i, 2, r['region'])
        ws.cell(6+i, 3, r['rate_display'])
    
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 5. 資料讀取與運算
# =========================================================
@st.cache_data(ttl=300)
def load_config_from_cloud(share_url):
    try:
        match = re.search(r"/d/([a-zA-Z0-9-_]+)", share_url)
        if not match: return None, None, None, None, "連結錯誤"
        file_id = match.group(1)
        def read_sheet(sheet_name):
            url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
            return pd.read_csv(url)
        
        df_store = read_sheet("Stores")
        df_store.columns = [c.strip() for c in df_store.columns]
        store_counts_num = dict(zip(df_store['Key'], df_store['Count']))
        
        df_fact = read_sheet("Factors")
        df_fact.columns = [c.strip() for c in df_fact.columns]
        sec_factors = {}
        for _, row in df_fact.iterrows():
            if row['Media'] not in sec_factors: sec_factors[row['Media']] = {}
            sec_factors[row['Media']][int(row['Seconds'])] = float(row['Factor'])
            
        name_map = {"全家新鮮視": "新鮮視", "全家廣播": "全家廣播", "家樂福": "家樂福"}
        for k, v in name_map.items():
            if k in sec_factors and v not in sec_factors: sec_factors[v] = sec_factors[k]
            
        df_price = read_sheet("Pricing")
        df_price.columns = [c.strip() for c in df_price.columns]
        pricing_db = {}
        for _, row in df_price.iterrows():
            m, r = row['Media'], row['Region']
            if m == "家樂福":
                if m not in pricing_db: pricing_db[m] = {}
                pricing_db[m][r] = {"List": int(row['List_Price']), "Net": int(row['Net_Price']), "Std_Spots": int(row['Std_Spots']), "Day_Part": row['Day_Part']}
            else:
                if m not in pricing_db: pricing_db[m] = {"Std_Spots": int(row['Std_Spots']), "Day_Part": row['Day_Part']}
                pricing_db[m][r] = [int(row['List_Price']), int(row['Net_Price'])]
        return None, store_counts_num, pricing_db, sec_factors, None
    except Exception as e: return None, None, None, None, str(e)

def calculate_plan_data(config, total_budget, days_count, pricing_db, sec_factors, store_counts_num, regions_order):
    rows, total_list_accum = [], 0
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
                unit_net_sum = sum([(db[r][1] / db["Std_Spots"]) * factor for r in calc_regs])
                if unit_net_sum == 0: continue
                
                spots_init = math.ceil(s_budget / unit_net_sum)
                is_under_target = spots_init < db["Std_Spots"]
                calc_penalty = 1.1 if is_under_target else 1.0 
                if cfg["is_national"]: row_display_penalty, total_display_penalty = 1.0, (1.1 if is_under_target else 1.0)
                else: row_display_penalty, total_display_penalty = (1.1 if is_under_target else 1.0), 1.0 
                
                spots_final = math.ceil(s_budget / (unit_net_sum * calc_penalty))
                if spots_final % 2 != 0: spots_final += 1
                if spots_final == 0: spots_final = 2
                
                sch = calculate_schedule(spots_final, days_count)
                nat_pkg_display = 0
                if cfg["is_national"]:
                    nat_list = db["全省"][0]
                    nat_pkg_display = int((nat_list / db["Std_Spots"]) * factor * total_display_penalty) * spots_final
                    total_list_accum += nat_pkg_display
                
                for i, r in enumerate(display_regs):
                    list_price_region = db[r][0]
                    total_rate_display = int((list_price_region / db["Std_Spots"]) * factor * row_display_penalty) * spots_final
                    if not cfg["is_national"]: total_list_accum += total_rate_display
                    
                    rows.append({
                        "media": m, "region": r, "program_num": store_counts_num.get(f"新鮮視_{r}" if m=="新鮮視" else r, 0),
                        "daypart": db["Day_Part"], "seconds": sec, "spots": spots_final, "schedule": sch,
                        "rate_display": total_rate_display, "pkg_display": total_rate_display,
                        "is_pkg_member": cfg["is_national"], "nat_pkg_display": nat_pkg_display
                    })
            elif m == "家樂福":
                db = pricing_db["家樂福"]
                base_std = db["量販_全省"]["Std_Spots"]
                unit_net = (db["量販_全省"]["Net"] / base_std) * factor
                spots_init = math.ceil(s_budget / unit_net)
                penalty = 1.1 if spots_init < base_std else 1.0
                spots_final = math.ceil(s_budget / (unit_net * penalty))
                if spots_final % 2 != 0: spots_final += 1
                
                sch_h = calculate_schedule(spots_final, days_count)
                total_rate_h = int((db["量販_全省"]["List"] / base_std) * factor * penalty) * spots_final
                total_list_accum += total_rate_h
                
                rows.append({
                    "media": m, "region": "全省量販", "program_num": store_counts_num["家樂福_量販"],
                    "daypart": db["量販_全省"]["Day_Part"], "seconds": sec, "spots": spots_final, "schedule": sch_h,
                    "rate_display": total_rate_h, "pkg_display": total_rate_h, "is_pkg_member": False
                })
                spots_s = int(spots_final * (db["超市_全省"]["Std_Spots"] / base_std))
                sch_s = calculate_schedule(spots_s, days_count)
                rows.append({
                    "media": m, "region": "全省超市", "program_num": store_counts_num["家樂福_超市"],
                    "daypart": db["超市_全省"]["Day_Part"], "seconds": sec, "spots": spots_s, "schedule": sch_s,
                    "rate_display": "計量販", "pkg_display": "計量販", "is_pkg_member": False
                })
    return rows, total_list_accum

# =========================================================
# 7. 主程式邏輯
# =========================================================
def main():
    try:
        with st.spinner("讀取設定檔..."):
            _, STORE_COUNTS_NUM, PRICING_DB, SEC_FACTORS, err_msg = load_config_from_cloud(GSHEET_SHARE_URL)
        if err_msg: st.error(err_msg); st.stop()
        
        # --- Sidebar ---
        with st.sidebar:
            st.header("🕵️ 主管登入")
            if not st.session_state.is_supervisor:
                if st.button("登入") or st.session_state.get('pwd_input') == "1234":
                    pwd = st.text_input("密碼", type="password", key="pwd_input")
                    if pwd == "1234": st.session_state.is_supervisor = True; st.rerun()
            else:
                st.success("✅ 主管模式")
                if st.button("登出"): st.session_state.is_supervisor = False; st.rerun()
            
            st.markdown("---")
            st.subheader("☁️ Ragic 連線設定")
            # 這裡會直接讀取最上方 DEFAULT_STATES 中的預設值
            if st.session_state.is_supervisor:
                st.session_state.ragic_url = st.text_input("URL", st.session_state.ragic_url)
                st.session_state.ragic_key = st.text_input("Key", st.session_state.ragic_key, type="password")
            else:
                st.text_input("URL", st.session_state.ragic_url, disabled=True)
                
            if st.button("Clear Cache"): st.cache_data.clear(); st.rerun()

        # --- Main UI ---
        st.title("📺 Cue Sheet Pro (Ragic Integrated)")
        format_type = st.radio("格式", ["Dongwu", "Shenghuo", "Bolin"], horizontal=True)
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: client_name = st.text_input("客戶", "測試客戶")
        with c2: product_name = st.text_input("產品", "測試產品")
        with c3: total_budget_input = st.number_input("預算 (Net)", value=1000000, step=10000)
        with c4: prod_cost_input = st.number_input("製作費", value=0, step=1000)
        with c5: sales_person = st.text_input("業務", "王小明")
        
        final_budget_val = total_budget_input
        if st.session_state.is_supervisor:
            col_sup1, col_sup2 = st.columns([1, 2])
            with col_sup2: 
                override = st.number_input("主管覆寫成交價", value=total_budget_input)
                if override != total_budget_input: final_budget_val = override
        
        c_d1, c_d2 = st.columns(2)
        start_date = c_d1.date_input("開始日", datetime(2026,1,1))
        end_date = c_d2.date_input("結束日", datetime(2026,1,31))
        days_count = (end_date - start_date).days + 1
        
        with st.expander("備註欄位", expanded=False):
            rc1, rc2, rc3 = st.columns(3)
            sign_deadline = rc1.date_input("回簽截止", datetime.now()+timedelta(days=3))
            billing_month = rc2.text_input("請款月", "2026年2月")
            payment_date = rc3.date_input("付款日", datetime(2026,3,31))

        # --- Media Selection (簡化版 UI，邏輯保留) ---
        st.markdown("### 媒體投放")
        config = {}
        col_m1, col_m2, col_m3 = st.columns(3)
        
        # 1. 全家廣播
        if col_m1.checkbox("全家廣播", key="cb_rad"):
            is_nat = col_m1.checkbox("廣播-全省", True)
            regs = ["全省"] if is_nat else col_m1.multiselect("廣播-區域", REGIONS_ORDER, REGIONS_ORDER)
            secs = col_m1.multiselect("廣播-秒數", DURATIONS, [20])
            share = col_m1.slider("廣播 %", 0, 100, key="rad_share")
            sec_shares = {secs[0]: 100} if secs else {}
            config["全家廣播"] = {"is_national": is_nat, "regions": regs, "sec_shares": sec_shares, "share": share}
        
        # 2. 新鮮視
        if col_m2.checkbox("新鮮視", key="cb_fv"):
            is_nat = col_m2.checkbox("新鮮視-全省", False)
            regs = ["全省"] if is_nat else col_m2.multiselect("新鮮視-區域", REGIONS_ORDER, ["北區"])
            secs = col_m2.multiselect("新鮮視-秒數", DURATIONS, [10])
            share = col_m2.slider("新鮮視 %", 0, 100, key="fv_share")
            sec_shares = {secs[0]: 100} if secs else {}
            config["新鮮視"] = {"is_national": is_nat, "regions": regs, "sec_shares": sec_shares, "share": share}

        # 3. 家樂福
        if col_m3.checkbox("家樂福", key="cb_cf"):
            secs = col_m3.multiselect("家樂福-秒數", DURATIONS, [20])
            share = col_m3.slider("家樂福 %", 0, 100, key="cf_share")
            sec_shares = {secs[0]: 100} if secs else {}
            config["家樂福"] = {"regions": ["全省"], "sec_shares": sec_shares, "share": share}

        # --- Calculation ---
        if config:
            rows, total_list_accum = calculate_plan_data(config, total_budget_input, days_count, PRICING_DB, SEC_FACTORS, STORE_COUNTS_NUM, REGIONS_ORDER)
            
            rem_list = get_remarks_text(sign_deadline, billing_month, payment_date)
            vat = int(round(final_budget_val * 0.05))
            grand_total = final_budget_val + vat
            p_display = f"{product_name}"
            
            # HTML
            html = generate_html_preview(rows, days_count, start_date, end_date, client_name, p_display, format_type, rem_list, total_list_accum, grand_total, final_budget_val, prod_cost_input)
            st.components.v1.html(html, height=500, scrolling=True)
            
            # Files
            xlsx_data = generate_excel_from_scratch(format_type, start_date, end_date, client_name, product_name, rows, rem_list, final_budget_val, prod_cost_input, sales_person)
            
            st.markdown("---")
            c_dl1, c_dl2, c_up = st.columns([1, 1, 2])
            
            with c_dl2:
                pdf_data, _, err = xlsx_bytes_to_pdf_bytes(xlsx_data)
                if pdf_data: st.download_button("📥 PDF", pdf_data, "cue.pdf", "application/pdf")
                else: st.warning("無 PDF 預覽")
            
            with c_dl1:
                if st.session_state.is_supervisor:
                    st.download_button("📥 Excel", xlsx_data, "cue.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else: st.info("Excel 限主管")
            
            # Ragic Upload
            with c_up:
                st.subheader("☁️ 上傳至 Ragic")
                if not st.session_state.ragic_confirm:
                    if st.button("🚀 準備上傳", type="primary"): st.session_state.ragic_confirm = True; st.rerun()
                else:
                    st.warning(f"確認上傳: {client_name} - {product_name} ?")
                    col_y, col_n = st.columns(2)
                    if col_n.button("❌ 取消"): st.session_state.ragic_confirm = False; st.rerun()
                    if col_y.button("✅ 確認"):
                        with st.spinner("上傳中..."):
                            
                            # =======================================================
                            # [關鍵設定區] 請填入您的 Ragic Field ID
                            # 請依照您在 Ragic 表單設計頁面看到的 ID 修改下方數字
                            # =======================================================
                            RAGIC_MAP = {
                                'client':     '1000080', # 客戶名稱
                                'product':    '1000081', # 產品名稱
                                'budget_raw': '1000082', # 原始預算 (Net)
                                'budget_fin': '1000083', # 優惠總價 (成交價)
                                'prod_cost':  '1000084', # 製作費
                                'format':     '1000078', # 格式類型
                                'sales':      '1000079', # 業務人員
                                'date_start': '1000085', # 走期-開始日
                                'date_end':   '1000086', # 走期-結束日
                                'date_sign':  '1000087', # 回簽截止日
                                'bill_month': '1000089', # 請款月份
                                'date_pay':   '1000088', # 付款兌現日
                                'details':    '1000090', # 投放參數詳情 (多行文字)
                                'file_xls':   '1000091', # Excel 檔案上傳
                                'file_pdf':   '1000092'  # PDF 檔案上傳
                            }
                            # =======================================================

                            # 1. 整理參數詳情文字
                            campaign_summary = format_campaign_details(config)

                            # 2. 準備資料 Payload (Ragic 接受字串格式)
                            data = {
                                RAGIC_MAP['client']:     client_name,
                                RAGIC_MAP['product']:    product_name,
                                RAGIC_MAP['budget_raw']: total_budget_input,
                                RAGIC_MAP['budget_fin']: final_budget_val,
                                RAGIC_MAP['prod_cost']:  prod_cost_input,
                                RAGIC_MAP['format']:     format_type,
                                RAGIC_MAP['sales']:      sales_person,
                                RAGIC_MAP['date_start']: str(start_date),
                                RAGIC_MAP['date_end']:   str(end_date),
                                RAGIC_MAP['date_sign']:  str(sign_deadline),
                                RAGIC_MAP['bill_month']: billing_month,
                                RAGIC_MAP['date_pay']:   str(payment_date),
                                RAGIC_MAP['details']:    campaign_summary,
                            }
                            
                            # 3. 準備檔案 (Binary Upload)
                            files = {}
                            # 上傳 Excel (必備)
                            files[RAGIC_MAP['file_xls']] = (f"Cue_{safe_filename(client_name)}.xlsx", xlsx_data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                            # 上傳 PDF (選備)
                            if pdf_data:
                                files[RAGIC_MAP['file_pdf']] = (f"Cue_{safe_filename(client_name)}.pdf", pdf_data, 'application/pdf')
                            
                            # 4. 送出至 Ragic
                            ok, msg = upload_to_ragic(st.session_state.ragic_url, st.session_state.ragic_key, data, files)
                            
                            if ok: st.success(msg); time.sleep(3)
                            else: st.error(msg)
                            
                        st.session_state.ragic_confirm = False
                        st.rerun()

    except Exception as e:
        st.error("系統發生錯誤，請聯絡管理員")
        st.error(traceback.format_exc())

if __name__ == "__main__":
    main()
