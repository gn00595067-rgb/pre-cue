import streamlit as st
import pandas as pd
import math
import io
import os
import re
import shutil
import tempfile
import subprocess
import gc
import requests
from datetime import timedelta, datetime, date
from itertools import groupby
from copy import copy

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker, XDRPositiveSize2D

# =========================================================
# 0. Streamlit Page
# =========================================================
st.set_page_config(layout="wide", page_title="Cue Sheet Pro (B方案-樣板套用)")

# =========================================================
# 1. Global Constants
# =========================================================
GSHEET_SHARE_URL = "https://docs.google.com/spreadsheets/d/1bzmG-N8XFsj8m3LUPqA8K70AcIqaK4Qhq1VPWcK0w_s/edit?usp=sharing"

# 如果你樣板內已經有 Logo（強烈建議），就不一定要用 URL 下載
BOLIN_LOGO_URL = "https://docs.google.com/drawings/d/17Uqgp-7LJJj9E4bV7Azo7TwXESPKTTIsmTbf-9tU9eE/export/png"

FONT_MAIN = "微軟正黑體"

REGIONS_ORDER = ["北區", "桃竹苗", "中區", "雲嘉南", "高屏", "東區"]
DURATIONS = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60]

REGION_DISPLAY_MAP = {
    "北區": "北區-北北基",
    "桃竹苗": "桃區-桃竹苗",
    "中區": "中區-中彰投",
    "雲嘉南": "雲嘉南區-雲嘉南",
    "高屏": "高屏區-高屏",
    "東區": "東區-宜花東",
    "全省量販": "全省量販",
    "全省超市": "全省超市"
}

MEDIA_ORDER = {"全家廣播": 1, "新鮮視": 2, "家樂福": 3}

TEMPLATE_PATHS = {
    "Dongwu": os.path.join("templates", "東吳樣板.xlsx"),
    "Shenghuo": os.path.join("templates", "生活樣板.xlsx"),
    "Bolin": os.path.join("templates", "鉑霖樣板.xlsx"),
}

# =========================================================
# 2. Small Helpers
# =========================================================
def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

def region_display(region: str) -> str:
    return REGION_DISPLAY_MAP.get(region, region)

def parse_gsheet_id(url: str):
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return m.group(1) if m else None

def col_width_to_pixels(excel_width: float) -> int:
    """
    近似換算：Excel column width -> pixels
    常見公式：px ≈ width*7 + 5
    """
    if excel_width is None:
        return 64
    return int(excel_width * 7 + 5)

def px_to_emu(px: int) -> int:
    # 1 px at 96 dpi = 9525 EMU
    return int(px * 9525)

def find_soffice_path():
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        return soffice
    if os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
    return None

@st.cache_data(show_spinner="正在下載 Logo...", ttl=3600)
def fetch_logo_bytes(url: str):
    try:
        r = requests.get(url, timeout=15)
        if r.status_code == 200:
            return r.content
    except:
        pass
    return None

@st.cache_data(show_spinner="正在生成 PDF (LibreOffice)...", ttl=3600)
def xlsx_bytes_to_pdf_bytes(xlsx_bytes: bytes):
    soffice = find_soffice_path()
    if not soffice:
        return None, "Fail", "找不到 LibreOffice (soffice)。雲端請用 packages.txt 安裝 libreoffice。"
    try:
        with tempfile.TemporaryDirectory() as tmp:
            xlsx_path = os.path.join(tmp, "cue.xlsx")
            with open(xlsx_path, "wb") as f:
                f.write(xlsx_bytes)

            # pdf:calc_pdf_Export 比較穩
            subprocess.run(
                [soffice, "--headless", "--nologo", "--convert-to", "pdf:calc_pdf_Export", "--outdir", tmp, xlsx_path],
                capture_output=True,
                timeout=90,
            )

            pdf_path = os.path.join(tmp, "cue.pdf")
            if not os.path.exists(pdf_path):
                # LibreOffice 有時候會用原檔名輸出
                for fn in os.listdir(tmp):
                    if fn.lower().endswith(".pdf"):
                        pdf_path = os.path.join(tmp, fn)
                        break

            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    return f.read(), "LibreOffice", ""
            return None, "Fail", "LibreOffice 未產出 PDF"
    except subprocess.TimeoutExpired:
        return None, "Fail", "轉檔逾時"
    except Exception as e:
        return None, "Fail", str(e)
    finally:
        gc.collect()

# =========================================================
# 3. Load Config from Google Sheet
# =========================================================
@st.cache_data(ttl=300)
def load_config_from_cloud(share_url):
    file_id = parse_gsheet_id(share_url)
    if not file_id:
        return None, None, None, None, "GSHEET 連結格式錯誤"

    def read_sheet(sheet_name):
        url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
        return pd.read_csv(url)

    try:
        df_store = read_sheet("Stores")
        df_store.columns = [c.strip() for c in df_store.columns]
        store_counts = dict(zip(df_store["Key"], df_store["Display_Name"]))
        store_counts_num = dict(zip(df_store["Key"], df_store["Count"]))

        df_fact = read_sheet("Factors")
        df_fact.columns = [c.strip() for c in df_fact.columns]
        sec_factors = {}
        for _, row in df_fact.iterrows():
            m = row["Media"]
            sec_factors.setdefault(m, {})
            sec_factors[m][int(row["Seconds"])] = float(row["Factor"])
        # alias
        name_map = {"全家新鮮視": "新鮮視", "全家廣播": "全家廣播", "家樂福": "家樂福"}
        for k, v in name_map.items():
            if k in sec_factors and v not in sec_factors:
                sec_factors[v] = sec_factors[k]

        df_price = read_sheet("Pricing")
        df_price.columns = [c.strip() for c in df_price.columns]
        pricing_db = {}
        for _, row in df_price.iterrows():
            m = row["Media"]
            r = row["Region"]
            if m == "家樂福":
                pricing_db.setdefault(m, {})
                pricing_db[m][r] = {
                    "List": int(row["List_Price"]),
                    "Net": int(row["Net_Price"]),
                    "Std_Spots": int(row["Std_Spots"]),
                    "Day_Part": row["Day_Part"],
                }
            else:
                pricing_db.setdefault(m, {"Std_Spots": int(row["Std_Spots"]), "Day_Part": row["Day_Part"]})
                pricing_db[m][r] = [int(row["List_Price"]), int(row["Net_Price"])]

        return store_counts, store_counts_num, pricing_db, sec_factors, None
    except Exception as e:
        return None, None, None, None, f"讀取失敗: {str(e)}"

def get_sec_factor(media_type, seconds, sec_factors):
    factors = sec_factors.get(media_type)
    if not factors:
        return 1.0
    if seconds in factors:
        return factors[seconds]
    # fallback: linear scaling
    for base in [10, 20, 15, 30]:
        if base in factors:
            return (seconds / base) * factors[base]
    return 1.0

def calculate_schedule(total_spots, days):
    if days <= 0:
        return []
    # 強制偶數
    if total_spots % 2 != 0:
        total_spots += 1
    half = total_spots // 2
    base, rem = divmod(half, days)
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
        f"6.付款兌現日期：{p_str}",
    ]

# =========================================================
# 4. Core Calculation (含你說的全省/分區 + 1.1 規則)
# =========================================================
def calculate_plan_data(config, total_budget, days_count, pricing_db, sec_factors, store_counts_num):
    """
    產出 rows：
      - rate_display: 顯示用（你要求 rate(Net) 顯示『該分區總價』，不是單檔）
      - pkg_display:  非全省：分區總價（若未達標 *1.1）
                     全省：分區顯示原價(不*1.1)，但 nat_pkg_display（package-cost）要 *1.1（若未達標）
      - nat_pkg_display: 全省打包價（合併格顯示）
      - is_pkg_member: 是否屬於全省合併 package-cost 的群組
    """
    rows = []
    total_list_accum = 0

    for m, cfg in config.items():
        m_budget_total = total_budget * (cfg["share"] / 100.0)

        for sec, sec_pct in cfg["sec_shares"].items():
            s_budget = m_budget_total * (sec_pct / 100.0)
            if s_budget <= 0:
                continue

            factor = get_sec_factor(m, sec, sec_factors)

            if m in ["全家廣播", "新鮮視"]:
                db = pricing_db[m]
                std_spots = db["Std_Spots"]
                daypart = db["Day_Part"]

                # 計算用區域：全省=>用全省；分區=>用選到的分區
                is_nat = cfg["is_national"]
                calc_regs = ["全省"] if is_nat else cfg["regions"]
                display_regs = REGIONS_ORDER if is_nat else cfg["regions"]

                # --- Net 用來決定 spots（含未達標時成本 *1.1）---
                unit_net_sum = 0.0
                for r in calc_regs:
                    unit_net_sum += (db[r][1] / std_spots) * factor

                if unit_net_sum <= 0:
                    continue

                spots_init = math.ceil(s_budget / unit_net_sum)
                is_under_target = spots_init < std_spots
                calc_penalty = 1.1 if is_under_target else 1.0

                spots_final = math.ceil(s_budget / (unit_net_sum * calc_penalty))
                if spots_final % 2 != 0:
                    spots_final += 1
                if spots_final <= 0:
                    spots_final = 2

                sch = calculate_schedule(spots_final, days_count)

                # --- 顯示規則（你的要求） ---
                # A) 沒選全省：分區價與加總都要遵守未達標 *1.1
                # B) 有選全省：package-cost(全省打包價) 若未達標要 *1.1
                #              但分區顯示價格不要再 *1.1（避免價差過大讓客戶懷疑）
                row_display_penalty = (1.1 if is_under_target else 1.0) if (not is_nat) else 1.0
                nat_pkg_penalty = (1.1 if is_under_target else 1.0) if is_nat else 1.0

                # 全省打包價（合併顯示）
                nat_pkg_display = 0
                if is_nat:
                    nat_list = db["全省"][0]
                    nat_pkg_display = int((nat_list / std_spots) * factor * nat_pkg_penalty * spots_final)
                    total_list_accum += nat_pkg_display

                for r in display_regs:
                    # 分區顯示用 list total（不一定等於全省計價）
                    list_price_region = db[r][0]
                    # 你要 rate(Net) 顯示「該分區總價」：所以這裡直接算總價
                    total_rate_display = int((list_price_region / std_spots) * factor * row_display_penalty * spots_final)

                    # package-cost 欄位：
                    # - 非全省：就顯示分區總價（同 total_rate_display）
                    # - 全省：該欄位由 nat_pkg_display 合併顯示（分區各列不填）
                    pkg_display = total_rate_display

                    # 非全省才把分區加總列入 total_list_accum（用於折扣率等）
                    if not is_nat:
                        total_list_accum += pkg_display

                    program_num_key = (f"新鮮視_{r}" if m == "新鮮視" else r)
                    rows.append({
                        "media": m,
                        "region": region_display(r),
                        "program_num": int(store_counts_num.get(program_num_key, 0)),
                        "daypart": daypart,
                        "seconds": int(sec),
                        "schedule": sch,
                        "spots": sum(sch),  # 右側檔次欄通常要顯示總檔次
                        "rate_display": total_rate_display,  # ✅你要求：顯示總價
                        "pkg_display": pkg_display,
                        "is_pkg_member": is_nat,
                        "nat_pkg_display": nat_pkg_display,
                    })

            elif m == "家樂福":
                db = pricing_db["家樂福"]
                base_std = db["量販_全省"]["Std_Spots"]
                daypart_h = db["量販_全省"]["Day_Part"]
                daypart_s = db["超市_全省"]["Day_Part"]

                # 用量販 Net 推 spots（你的原邏輯）
                unit_net = (db["量販_全省"]["Net"] / base_std) * factor
                spots_init = math.ceil(s_budget / unit_net)
                penalty = 1.1 if spots_init < base_std else 1.0

                spots_final = math.ceil(s_budget / (unit_net * penalty))
                if spots_final % 2 != 0:
                    spots_final += 1
                if spots_final <= 0:
                    spots_final = 2

                sch_h = calculate_schedule(spots_final, days_count)
                # List 顯示（總價）
                unit_list_h = (db["量販_全省"]["List"] / base_std) * factor * penalty
                total_rate_h = int(unit_list_h * spots_final)
                total_list_accum += total_rate_h

                rows.append({
                    "media": "家樂福",
                    "region": "全省量販",
                    "program_num": int(store_counts_num.get("家樂福_量販", 0)),
                    "daypart": daypart_h,
                    "seconds": int(sec),
                    "schedule": sch_h,
                    "spots": sum(sch_h),
                    "rate_display": total_rate_h,
                    "pkg_display": total_rate_h,
                    "is_pkg_member": False,
                    "nat_pkg_display": 0,
                })

                # 超市：計量販（依你原邏輯）
                ratio = db["超市_全省"]["Std_Spots"] / base_std
                spots_s = int(round(spots_final * ratio))
                sch_s = calculate_schedule(spots_s, days_count)

                rows.append({
                    "media": "家樂福",
                    "region": "全省超市",
                    "program_num": int(store_counts_num.get("家樂福_超市", 0)),
                    "daypart": daypart_s,
                    "seconds": int(sec),
                    "schedule": sch_s,
                    "spots": sum(sch_s),
                    "rate_display": "計量販",
                    "pkg_display": "計量販",
                    "is_pkg_member": False,
                    "nat_pkg_display": 0,
                })

    # 排序
    rows.sort(key=lambda x: (MEDIA_ORDER.get(x["media"], 99), x["seconds"], x["region"]))
    return rows, total_list_accum

# =========================================================
# 5. B方案：用樣板產出 Excel（保留所有樣式/Logo/框線）
# =========================================================
def _remove_merges_in_range(ws, min_row, max_row, min_col, max_col):
    to_remove = []
    for r in list(ws.merged_cells.ranges):
        # merged range bounds
        if (r.min_row <= max_row and r.max_row >= min_row and
            r.min_col <= max_col and r.max_col >= min_col):
            to_remove.append(str(r))
    for addr in to_remove:
        try:
            ws.unmerge_cells(addr)
        except:
            pass

def _clear_values(ws, min_row, max_row, min_col, max_col):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(r, c).value = None

def align_logo_right_to_table(ws, img: XLImage, anchor_row_1based: int, anchor_start_col_1based: int, table_last_col_1based: int):
    """
    讓圖片右側切齊 table_last_col 的右邊界
    做法：以 anchor_start_col 為起點，計算 (anchor_start_col..table_last_col) 的像素寬，
         設 colOff = totalWidth - imgWidth
    """
    # 取得欄寬(px)
    total_px = 0
    for c in range(anchor_start_col_1based, table_last_col_1based + 1):
        letter = get_column_letter(c)
        w = ws.column_dimensions[letter].width
        total_px += col_width_to_pixels(w)

    # openpyxl Image 寬高為 px
    img_w = int(img.width)
    offset_px = max(0, total_px - img_w)

    marker = AnchorMarker(
        col=anchor_start_col_1based - 1,
        colOff=px_to_emu(offset_px),
        row=anchor_row_1based - 1,
        rowOff=px_to_emu(0),
    )
    ext = XDRPositiveSize2D(cx=px_to_emu(img_w), cy=px_to_emu(int(img.height)))
    img.anchor = OneCellAnchor(_from=marker, ext=ext)

def generate_excel_from_template(format_type: str,
                                template_path: str,
                                start_dt: date,
                                end_dt: date,
                                client_name: str,
                                product_name: str,
                                rows: list,
                                remarks: list,
                                final_budget_val: int,
                                prod_cost: int,
                                store_counts=None):
    """
    依 format_type 套用對應樣板。
    假設你的樣板已經把：
      - 欄寬、列高、顏色、框線、字體、Logo、頁首頁尾等都設好
    我們只做：
      1) 更新客戶/走期/產品等文字
      2) 更新日期欄
      3) 清掉舊資料、填入新 rows
      4) 重建 Station 合併、全省 package-cost 合併
      5) 設列印區域 + fitToWidth 避免 PDF 裁切
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到樣板：{template_path}。請放在 templates/ 內。")

    wb = openpyxl.load_workbook(template_path)
    ws = wb.active  # 你也可以改成指定名稱

    eff_days = (end_dt - start_dt).days + 1
    if eff_days <= 0:
        raise ValueError("日期區間錯誤：結束日必須 >= 開始日")

    # ----- 不同樣板的座標設定（依你現有生成器慣例）-----
    if format_type == "Dongwu":
        # 固定欄 A-G + 日期從 H 開始 + 最後檔次欄
        fixed_cols = 7
        day_col_start = 8
        last_col = fixed_cols + eff_days + 1  # spots col
        header_day_row = 7
        header_wk_row = 8
        data_start_row = 9
        pkg_col = 7
        station_col = 1
        # Header cells（沿用你 scratch 版本位置）
        # A3(客戶) / A4(Product) / A5(Period) / A6(Medium) 你樣板若不同可自行調整
        ws["B3"].value = client_name
        ws["B4"].value = f"{'、'.join([f'{s}秒' for s in sorted(set(r['seconds'] for r in rows))])} {product_name}"
        ws["B5"].value = f"{start_dt.strftime('%Y. %m. %d')} - {end_dt.strftime('%Y. %m. %d')}"
        ws["B6"].value = "/".join(sorted(set(r["media"] for r in rows), key=lambda x: MEDIA_ORDER.get(x, 99)))

    else:
        # Shenghuo / Bolin：固定欄 A-E + 日期從 F 開始 + 末端 檔次/定價/專案價
        fixed_cols = 5
        day_col_start = 6
        end_c_start = fixed_cols + eff_days + 1  # 檔次欄
        last_col = end_c_start + 2  # + 定價 + 專案價
        if format_type == "Shenghuo":
            header_day_row = 8  # 日期數字列
            header_wk_row = None  # 生活樣板常只有日期列，或你可自行改
            data_start_row = 9
        else:  # Bolin
            header_day_row = 6
            header_wk_row = 7
            data_start_row = 8

        station_col = 1
        pkg_col = end_c_start + 2  # 專案價欄（你原本就是用這欄當 package-cost 顯示）

        # 常用欄位（依你 scratch 版本）
        # 鉑霖：B2 client / B4 client / B5 product / 右側 period/spec 在第4列
        if format_type == "Bolin":
            ws["B2"].value = client_name
            ws["B4"].value = client_name
            ws["B5"].value = product_name
            # Spec + Period（你樣板若有固定合併格就只要寫入左上角）
            sec_str = " ".join([f"{s}秒廣告" for s in sorted(set(r["seconds"] for r in rows))])
            # 這兩格位置需配合你的鉑霖樣板（若不同自行改）
            ws["F4"].value = f"廣告規格：{sec_str}"
            ws.cell(4, last_col - 1).value = f"執行期間：{start_dt.strftime('%Y.%m.%d')} - {end_dt.strftime('%Y.%m.%d')}"

    # ----- 日期列寫入（盡量沿用樣板格式，只填值）-----
    weekdays_zh = ["一", "二", "三", "四", "五", "六", "日"]
    curr = start_dt
    for i in range(eff_days):
        col = day_col_start + i
        d_cell = ws.cell(header_day_row, col)
        d_cell.value = curr.day
        if header_wk_row:
            w_cell = ws.cell(header_wk_row, col)
            w_cell.value = weekdays_zh[curr.weekday()]
        curr += timedelta(days=1)

    # ----- 清掉舊資料區（只清 value，不動格式）-----
    # 預設清 200 行，足夠大多數情況；你可視需要調整
    clear_max_row = data_start_row + 200
    _remove_merges_in_range(ws, data_start_row, clear_max_row, 1, last_col)
    _clear_values(ws, data_start_row, clear_max_row, 1, last_col)

    # ----- 寫入資料（同樣只寫 value；格式靠樣板原本 cell style）-----
    def media_display_name(m):
        if format_type in ["Shenghuo", "Bolin"]:
            # 這兩個樣板通常 Station 欄就是 "全家廣播/新鮮視/家樂福"
            return m
        # Dongwu 的 Station 欄會顯示兩行
        if m == "全家廣播":
            return "全家便利商店\n通路廣播廣告"
        if m == "新鮮視":
            return "全家便利商店\n新鮮視廣告"
        return "家樂福"

    # 寫 rows
    out_row = data_start_row
    # group key：同 media + seconds（方便合併 Station 與全省 package）
    rows_sorted = sorted(rows, key=lambda x: (MEDIA_ORDER.get(x["media"], 99), x["seconds"], x["region"]))
    groups = []
    for k, g in groupby(rows_sorted, key=lambda x: (x["media"], x["seconds"])):
        groups.append((k, list(g)))

    for (m, sec), g_list in groups:
        group_start = out_row
        for r in g_list:
            # A: Station / B: Region / C: store count / D: daypart / E: spec
            ws.cell(out_row, 1).value = media_display_name(m)
            ws.cell(out_row, 2).value = r["region"]

            # store count
            cnt = r.get("program_num", 0)
            if format_type == "Dongwu":
                ws.cell(out_row, 3).value = cnt
            else:
                suffix = "面" if m == "新鮮視" else "店"
                ws.cell(out_row, 3).value = f"{int(cnt):,}{suffix}" if isinstance(cnt, (int, float)) else cnt

            ws.cell(out_row, 4).value = r["daypart"]

            # 規格欄
            if format_type == "Dongwu":
                ws.cell(out_row, 5).value = f"{r['seconds']}秒"
            else:
                if m == "新鮮視":
                    ws.cell(out_row, 5).value = f"{r['seconds']}秒\n影片/影像 1920x1080 (mp4)"
                else:
                    ws.cell(out_row, 5).value = f"{r['seconds']}秒廣告"

            # 日期排程
            for i, v in enumerate(r["schedule"][:eff_days]):
                ws.cell(out_row, day_col_start + i).value = int(v)

            # 檔次 / 金額欄
            if format_type == "Dongwu":
                # F: rate(Net) / G: Package-cost(Net)
                # 你要求 rate(Net) 顯示總價 → 直接填 total
                ws.cell(out_row, 6).value = r["rate_display"]

                # 全省：合併顯示 nat_pkg_display；分區：顯示自身 pkg_display
                if r.get("is_pkg_member"):
                    # 先不填（後面統一合併後在第一列填 nat_pkg_display）
                    pass
                else:
                    ws.cell(out_row, 7).value = r["pkg_display"]

                # spots 欄（最後一欄）
                ws.cell(out_row, last_col).value = int(r["spots"])

            else:
                # 檔次欄
                end_c_start = fixed_cols + eff_days + 1
                ws.cell(out_row, end_c_start).value = int(r["spots"])
                # 定價欄：用 rate_display
                ws.cell(out_row, end_c_start + 1).value = r["rate_display"]
                # 專案價欄：全省合併顯示 nat_pkg_display；分區顯示 pkg_display
                if r.get("is_pkg_member"):
                    pass
                else:
                    ws.cell(out_row, end_c_start + 2).value = r["pkg_display"]

            out_row += 1

        group_end = out_row - 1

        # ----- 合併 Station 欄（整段 group）-----
        if group_end > group_start:
            ws.merge_cells(start_row=group_start, start_column=station_col, end_row=group_end, end_column=station_col)

        # ----- 全省 package-cost 合併 -----
        if g_list and g_list[0].get("is_pkg_member"):
            # package-cost 欄（Dongwu=G；Bolin/Shenghuo=專案價）
            ws.merge_cells(start_row=group_start, start_column=pkg_col, end_row=group_end, end_column=pkg_col)
            # 在第一列填 nat_pkg_display
            ws.cell(group_start, pkg_col).value = g_list[0].get("nat_pkg_display", 0)

    last_data_row = out_row - 1

    # ----- Remarks（如果樣板有固定區塊，你可改成寫到指定位置）-----
    # 這裡用「寫在資料後方」的方式，避免破壞你樣板既有區塊
    remark_row = last_data_row + 2
    ws.cell(remark_row, 1).value = "Remarks："
    for i, rm in enumerate(remarks, start=1):
        ws.cell(remark_row + i, 1).value = rm

    # ----- 頁面設定：避免 PDF 左右被切 -----
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # 讓高度可延伸，重點是寬度不要裁切
    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.3
    ws.page_margins.bottom = 0.3

    # Print Area：鎖住 A1 到表格最右欄、最後備註列（更不容易被裁切）
    ws.print_area = f"A1:{get_column_letter(last_col)}{remark_row + len(remarks) + 2}"

    # ----- 鉑霖 Logo 右對齊（B方案關鍵修正）-----
    if format_type == "Bolin":
        # 1) 先嘗試抓樣板內既有圖片（最理想：你樣板本來就有 Logo）
        img = None
        if hasattr(ws, "_images") and ws._images:
            img = ws._images[0]  # 假設第一張就是 Logo
        else:
            # 2) 沒有就用 URL 下載加上去
            logo_bytes = fetch_logo_bytes(BOLIN_LOGO_URL)
            if logo_bytes:
                img = XLImage(io.BytesIO(logo_bytes))
                ws.add_image(img)

        if img:
            # 讓 Logo 的右側切齊「表格最右邊界」
            # 我們把 anchor 起點放在倒數第二欄（通常是定價那個大欄），再把右邊對齊到 last_col
            anchor_row = 1
            anchor_start_col = max(1, last_col - 1)  # 倒數第二欄
            try:
                align_logo_right_to_table(ws, img, anchor_row, anchor_start_col, last_col)
            except:
                # 即便對齊失敗也不中斷（至少不會把整份報表改壞）
                pass

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# =========================================================
# 6. Streamlit UI
# =========================================================
def main():
    with st.spinner("正在讀取 Google 試算表設定檔..."):
        STORE_COUNTS, STORE_COUNTS_NUM, PRICING_DB, SEC_FACTORS, err = load_config_from_cloud(GSHEET_SHARE_URL)
    if err:
        st.error(f"❌ 設定檔載入失敗: {err}")
        st.stop()

    st.title("📺 Cue 表生成器（B方案：樣板套用，極致擬真）")

    format_type = st.radio("選擇樣板格式", ["Dongwu", "Shenghuo", "Bolin"], horizontal=True)
    template_path = TEMPLATE_PATHS[format_type]
    st.caption(f"使用樣板：{template_path}")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        client_name = st.text_input("客戶名稱", "萬國通路")
    with c2:
        product_name = st.text_input("產品名稱", "統一布丁")
    with c3:
        total_budget_input = st.number_input("總預算 (未稅 Net)", value=1000000, step=10000)
    with c4:
        prod_cost_input = st.number_input("製作費 (未稅)", value=0, step=1000)

    d1, d2 = st.columns(2)
    with d1:
        start_date = st.date_input("開始日", date(2026, 1, 1))
    with d2:
        end_date = st.date_input("結束日", date(2026, 1, 31))
    days_count = (end_date - start_date).days + 1
    if days_count <= 0:
        st.error("日期區間錯誤：結束日必須 >= 開始日")
        st.stop()
    st.info(f"📅 走期共 **{days_count}** 天")

    with st.expander("📝 備註欄位設定", expanded=False):
        rc1, rc2, rc3 = st.columns(3)
        sign_deadline = rc1.date_input("回簽截止日", date.today() + timedelta(days=3))
        billing_month = rc2.text_input("請款月份", "2026年2月")
        payment_date = rc3.date_input("付款兌現日", date(2026, 3, 31))

    st.markdown("### 1) 媒體投放設定")

    colA, colB, colC = st.columns(3)
    config = {}

    with colA:
        st.markdown("#### 📻 全家廣播")
        rad_on = st.checkbox("啟用", value=True, key="rad_on")
        if rad_on:
            rad_nat = st.checkbox("全省聯播", value=True, key="rad_nat")
            rad_regs = ["全省"] if rad_nat else st.multiselect("區域", REGIONS_ORDER, default=REGIONS_ORDER, key="rad_regs")
            if (not rad_nat) and len(rad_regs) == 6:
                rad_nat = True
                rad_regs = ["全省"]
                st.info("✅ 已選滿 6 區，自動視為全省聯播")
            rad_secs = st.multiselect("秒數", DURATIONS, default=[20], key="rad_secs")
            rad_share = st.slider("預算佔比%", 0, 100, 70, key="rad_share")

            sec_shares = {}
            if len(rad_secs) > 1:
                rem = 100
                for i, s in enumerate(sorted(rad_secs)):
                    if i < len(rad_secs) - 1:
                        v = st.slider(f"{s}秒佔比", 0, rem, int(rem / 2), key=f"rad_s_{s}")
                        sec_shares[int(s)] = int(v)
                        rem -= v
                    else:
                        sec_shares[int(s)] = int(rem)
            elif rad_secs:
                sec_shares[int(rad_secs[0])] = 100

            config["全家廣播"] = {
                "is_national": bool(rad_nat),
                "regions": rad_regs,
                "sec_shares": sec_shares,
                "share": int(rad_share),
            }

    with colB:
        st.markdown("#### 📺 新鮮視")
        fv_on = st.checkbox("啟用", value=True, key="fv_on")
        if fv_on:
            fv_nat = st.checkbox("全省聯播 ", value=False, key="fv_nat")
            fv_regs = ["全省"] if fv_nat else st.multiselect("區域", REGIONS_ORDER, default=["北區", "中區"], key="fv_regs")
            if (not fv_nat) and len(fv_regs) == 6:
                fv_nat = True
                fv_regs = ["全省"]
                st.info("✅ 已選滿 6 區，自動視為全省聯播")
            fv_secs = st.multiselect("秒數", DURATIONS, default=[10], key="fv_secs")
            fv_share = st.slider("預算佔比% ", 0, 100, 20, key="fv_share")

            sec_shares = {}
            if len(fv_secs) > 1:
                rem = 100
                for i, s in enumerate(sorted(fv_secs)):
                    if i < len(fv_secs) - 1:
                        v = st.slider(f"{s}秒佔比 ", 0, rem, int(rem / 2), key=f"fv_s_{s}")
                        sec_shares[int(s)] = int(v)
                        rem -= v
                    else:
                        sec_shares[int(s)] = int(rem)
            elif fv_secs:
                sec_shares[int(fv_secs[0])] = 100

            config["新鮮視"] = {
                "is_national": bool(fv_nat),
                "regions": fv_regs,
                "sec_shares": sec_shares,
                "share": int(fv_share),
            }

    with colC:
        st.markdown("#### 🛒 家樂福")
        cf_on = st.checkbox("啟用", value=True, key="cf_on")
        if cf_on:
            cf_secs = st.multiselect("秒數", DURATIONS, default=[20], key="cf_secs")
            cf_share = st.slider("預算佔比%", 0, 100, 10, key="cf_share")

            sec_shares = {}
            if len(cf_secs) > 1:
                rem = 100
                for i, s in enumerate(sorted(cf_secs)):
                    if i < len(cf_secs) - 1:
                        v = st.slider(f"{s}秒佔比  ", 0, rem, int(rem / 2), key=f"cf_s_{s}")
                        sec_shares[int(s)] = int(v)
                        rem -= v
                    else:
                        sec_shares[int(s)] = int(rem)
            elif cf_secs:
                sec_shares[int(cf_secs[0])] = 100

            config["家樂福"] = {
                "is_national": True,
                "regions": ["全省"],
                "sec_shares": sec_shares,
                "share": int(cf_share),
            }

    # Normalize shares (optional but recommended)
    if config:
        total_share = sum(v["share"] for v in config.values())
        if total_share != 100 and total_share > 0:
            st.warning(f"目前預算佔比合計 {total_share}%，建議調整為 100%（系統仍可照比例運算）")

    st.markdown("### 2) 生成結果")
    if st.button("🚀 生成 Cue 表", type="primary"):
        if not config:
            st.error("請至少啟用一個媒體")
            st.stop()

        remarks = get_remarks_text(sign_deadline, billing_month, payment_date)

        with st.spinner("正在計算排程與金額..."):
            rows, total_list_accum = calculate_plan_data(
                config=config,
                total_budget=total_budget_input,
                days_count=days_count,
                pricing_db=PRICING_DB,
                sec_factors=SEC_FACTORS,
                store_counts_num=STORE_COUNTS_NUM,
            )

        if not rows:
            st.error("沒有產出任何列（請檢查是否預算/秒數/區域設定為 0）")
            st.stop()

        with st.spinner("正在套用樣板產出 Excel（B方案）..."):
            try:
                xlsx_bytes = generate_excel_from_template(
                    format_type=format_type,
                    template_path=template_path,
                    start_dt=start_date,
                    end_dt=end_date,
                    client_name=client_name,
                    product_name=product_name,
                    rows=rows,
                    remarks=remarks,
                    final_budget_val=int(total_budget_input),
                    prod_cost=int(prod_cost_input),
                    store_counts=STORE_COUNTS,
                )
            except Exception as e:
                st.error("產生 Excel 失敗")
                st.exception(e)
                st.stop()

        st.success("✅ Excel 產生完成（B方案樣板套用）")
        st.download_button(
            "📥 下載 Excel",
            data=xlsx_bytes,
            file_name=f"Cue_{safe_filename(client_name)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # PDF
        st.info("PDF 需要 LibreOffice（本機安裝或雲端 packages.txt 裝 libreoffice）")
        with st.spinner("正在轉出 PDF..."):
            pdf_bytes, method, err = xlsx_bytes_to_pdf_bytes(xlsx_bytes)
        if pdf_bytes:
            st.download_button(
                "📥 下載 PDF",
                data=pdf_bytes,
                file_name=f"Cue_{safe_filename(client_name)}.pdf",
                mime="application/pdf",
            )
        else:
            st.warning(f"PDF 生成失敗：{err}")

if __name__ == "__main__":
    main()
