import pandas as pd
import numpy as np
import os
import shutil
import warnings
import openpyxl
import time
import glob
import stat
import math
import sys
import io
import threading
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

warnings.filterwarnings('ignore')

# ---- customtkinter 全局外观设置 ----
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ================================================================
# ==================== 第一段：全局配置区 =======================
# ================================================================

# 脚本所在目录（自动计算，无需手动修改）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ---------- 各平台数据输入目录 ----------
SHOPEE_DATA_DIR     = os.path.join(BASE_DIR, 'Data', 'Shopee')
TK_NOHORSE_DATA_DIR = os.path.join(BASE_DIR, 'Data', 'TK_NoHorse')
TK_HORSE_DATA_DIR   = os.path.join(BASE_DIR, 'Data', 'TK_Horse')

# ---------- 各平台结果输出目录 ----------
SHOPEE_RESULT_DIR     = os.path.join(BASE_DIR, 'Result', 'Shopee')
TK_NOHORSE_RESULT_DIR = os.path.join(BASE_DIR, 'Result', 'TK_NoHorse')
TK_HORSE_RESULT_DIR   = os.path.join(BASE_DIR, 'Result', 'TK_Horse')

# ---------- 成本表文件名（Shopee / TK无马帮 共用） ----------
COST_FILE_NAME = '商品成本.xlsx'

# ===== Shopee 平台参数（可随时修改） =====
SHOPEE_EXCHANGE_RATE = 0.117789   # 汇率：1 PHP = ? CNY
SHOPEE_UNIT_FEE      = 3.0        # 发单费/空包费单价（CNY）
SHOPEE_TARGET_YEAR   = 2026       # 目标年份
SHOPEE_TARGET_MONTH  = 1          # 目标月份（1-12）

SHOPEE_SHEET_NAMES = {
    'order':     ['订单明细原表', '订单明细', '订单明细表'],
    'income':    ['收入明细表', '收入明细'],
    'brush':     ['补单明细', '补单', '补单明细 '],
    'ads':       ['广告费', '广告'],
    'sales':     ['销售明细', '销售明细表'],
    'pnl':       ['利润表', 'Profit'],
    'unmatched': ['未匹配', '未匹配订单'],
}

# ===== TK无马帮 平台参数（可随时修改） =====
TK_NOHORSE_EXCHANGE_RATE = 0.117789   # 汇率：1 PHP = ? CNY
TK_NOHORSE_UNIT_FEE      = 3.0        # 发单费/空包费单价（CNY）
TK_NOHORSE_TARGET_YEAR   = 2026       # 目标年份
TK_NOHORSE_TARGET_MONTH  = 1          # 目标月份（1-12）

TK_NOHORSE_SHEET_NAMES = {
    'order':     ['订单明细原表', '订单明细'],
    'income':    ['收入明细表', '收入明细'],
    'brush':     ['补单明细', '补单'],
    'ads':       ['广告费', '广告'],
    'sales':     ['销售明细', '销售明细表'],
    'pnl':       ['利润表', 'Profit'],
    'unmatched': ['未匹配', '未匹配订单'],
}

# ===== TK有马帮 平台参数（可随时修改） =====
TK_HORSE_TARGET_MONTH = 2   # 目标月份（1-12），用于账单筛选和利润表列定位

# ================================================================
# ==================== 第二段：公共工具函数 ======================
# ================================================================

def find_header_row(df, keywords):
    """在前20行中查找包含所有关键词的行，返回行索引"""
    for idx, row in df.head(20).iterrows():
        row_str = " ".join([str(x) for x in row.values])
        if all(k in row_str for k in keywords):
            return idx
    return 0

def load_and_clean_sheet(excel_file, sheet_candidates, keywords):
    """尝试多个Sheet候选名，找到后自动定位表头行并读取"""
    sheet_name = None
    for name in sheet_candidates:
        if name in excel_file.sheet_names:
            sheet_name = name
            break
    if not sheet_name:
        return pd.DataFrame()
    try:
        df_preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=20)
        header_idx = find_header_row(df_preview, keywords)
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_idx, dtype=str)
        return df
    except:
        return pd.DataFrame()

def clean_currency(x):
    """清洗货币字符串，返回浮点数"""
    if pd.isna(x):
        return 0.0
    s = str(x).replace('PHP', '').replace('RM', '').replace(',', '').replace(' ', '').strip()
    try:
        return float(s)
    except:
        return 0.0

def safe_float(x):
    """强力转浮点数（兼容NaN/None/特殊符号）"""
    if x is None:
        return 0.0
    try:
        if isinstance(x, (int, float)):
            if math.isnan(x):
                return 0.0
            return float(x)
        s = str(x).strip().lower()
        if s in ['nan', '', 'none']:
            return 0.0
        s_clean = s.replace(',', '').replace(' ', '').replace('¥', '').replace('₱', '').replace('rp', '')
        val = float(s_clean)
        return 0.0 if math.isnan(val) else val
    except:
        return 0.0

def safe_border(ws, row_idx, col_idx):
    """给单元格加细边框并居中"""
    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    cell = ws.cell(row=row_idx, column=col_idx)
    cell.border = border
    cell.alignment = Alignment(horizontal='center', vertical='center')

def find_month_col(ws_pnl, target_month):
    """
    统一利润表月份列定位函数（仿照TK_Horse样式）。
    遍历前5行，查找包含 "X月" 的表头单元格，返回列索引。
    找不到时默认返回2（B列），并打印警告。
    """
    target_str = f"{target_month}月"
    for row in ws_pnl.iter_rows(min_row=1, max_row=5):
        for cell in row:
            if cell.value and str(cell.value).strip().replace(" ", "") == target_str:
                print(f"   找到 '{target_str}' 数据列: 第 {cell.column} 列 ({get_column_letter(cell.column)}列)")
                return cell.column
    print(f"   [警告] 利润表中未找到 '{target_str}' 的表头，将默认写入第2列(B列)")
    return 2


# ================================================================
# ================ 第二段 A：Shopee 平台处理函数 =================
# ================================================================

def process_shopee(source_path, cost_map, result_dir):
    """Shopee平台单文件处理：读取→计算→写销售明细→写补单→写利润表→写未匹配"""
    file_name = os.path.basename(source_path)
    SN = SHOPEE_SHEET_NAMES
    ER = SHOPEE_EXCHANGE_RATE
    UF = SHOPEE_UNIT_FEE
    TY = SHOPEE_TARGET_YEAR
    TM = SHOPEE_TARGET_MONTH

    print(f"\n{'='*60}")
    print(f">>> [Shopee] 正在处理: {file_name}")

    # ---------- 1. 读取各Sheet ----------
    xls = pd.ExcelFile(source_path)
    df_order  = load_and_clean_sheet(xls, SN['order'],  ['Order ID', 'SKU'])
    df_income = load_and_clean_sheet(xls, SN['income'], ['Order ID', 'Total Released Amount'])
    df_brush  = load_and_clean_sheet(xls, SN['brush'],  ['线上单号', '佣金'])
    df_ads    = load_and_clean_sheet(xls, SN['ads'],    ['花费'])

    if df_order.empty or df_income.empty:
        print("   [跳过] 缺少订单或收入明细表")
        return

    # ---------- 2. 时间筛选 ----------
    if 'Order Creation Date' in df_order.columns:
        date_col = 'Order Creation Date'
    elif 'Order created time' in df_order.columns:
        date_col = 'Order created time'
    else:
        date_col = df_order.columns[0]

    df_order['Created Time'] = pd.to_datetime(df_order[date_col], errors='coerce')
    mask_order = (df_order['Created Time'].dt.year == TY) & (df_order['Created Time'].dt.month == TM)
    df_order_m = df_order[mask_order].copy()
    print(f"   {TM}月订单数: {len(df_order_m)}")
    if len(df_order_m) == 0:
        print(f"   [跳过] {TM}月无订单")
        return

    inc_date_col = 'Order Creation Date' if 'Order Creation Date' in df_income.columns else 'Order created time'
    df_income['Created Time'] = pd.to_datetime(df_income[inc_date_col], errors='coerce')
    mask_income = (df_income['Created Time'].dt.year == TY) & (df_income['Created Time'].dt.month == TM)
    df_income_m = df_income[mask_income].copy()

    # ---------- 3. 成本计算 ----------
    sku_col = next((c for c in df_order_m.columns if 'Reference' in c and 'Parent' not in c), None)
    if not sku_col:
        sku_col = next((c for c in df_order_m.columns if 'SKU' in c), 'Seller SKU')
    qty_col    = 'Quantity'
    track_col  = next((c for c in df_order_m.columns if 'Tracking' in c), 'Tracking Number*')
    status_col = 'Order Status'

    df_order_m['Line_Total_Cost'] = df_order_m.apply(
        lambda row: cost_map.get(str(row[sku_col]).strip(), 0) * clean_currency(row[qty_col]), axis=1
    )

    order_agg = df_order_m.groupby('Order ID').agg({
        'Created Time': 'first',
        sku_col:        lambda x: '、'.join([str(i) for i in x if str(i) != 'nan']),
        qty_col:        lambda x: sum([clean_currency(i) for i in x]),
        'Line_Total_Cost': 'sum',
        status_col:     'first',
        track_col:      'first'
    }).reset_index()
    order_agg.rename(columns={sku_col: 'Merged_SKU', 'Line_Total_Cost': 'Total_Cost_CNY'}, inplace=True)

    # ---------- 4. 跨表匹配，生成销售明细 ----------
    cols      = df_income.columns
    c_price   = next((c for c in cols if 'Original Product Price' in c), None)
    c_promo   = next((c for c in cols if 'Seller Product Promotion' in c), None)
    c_voucher = next((c for c in cols if 'Voucher Sponsored by Seller' in c), None)
    c_refund  = next((c for c in cols if 'Refund Amount' in c), None)
    income_dict = df_income.set_index('Order ID').to_dict('index')

    sale_details = []
    for _, row in order_agg.iterrows():
        oid     = row['Order ID']
        inc_row = income_dict.get(oid, {})
        sale_details.append({
            '日期':           row['Created Time'],
            '订单号':         oid,
            '商品SKU':        row['Merged_SKU'],
            '销售数量':       row[qty_col],
            '销售单价':       clean_currency(inc_row.get(c_price, 0)),
            '优惠金额':       clean_currency(inc_row.get(c_promo, 0)),
            '优惠劵金额':     clean_currency(inc_row.get(c_voucher, 0)),
            '退款金额':       clean_currency(inc_row.get(c_refund, 0)),
            '商品成本总额(CNY)': row['Total_Cost_CNY'],
            'Status':         row[status_col],
            'Tracking':       row[track_col],
        })
    df_sales = pd.DataFrame(sale_details)
    df_sales = df_sales[df_sales['销售单价'] > 0.001].copy()
    df_sales.reset_index(drop=True, inplace=True)

    # ---------- 5. 未匹配统计 ----------
    ids_order   = set(df_order_m['Order ID'])
    ids_income  = set(df_income['Order ID'])
    unmatched_ids = ids_order - ids_income
    df_unmatched  = df_order_m[df_order_m['Order ID'].isin(unmatched_ids)].copy()

    count_cancel_shipped = count_cancel_unshipped = 0
    if not df_unmatched.empty:
        def check_status(row):
            track  = str(row.get(track_col, ''))
            status = str(row.get(status_col, ''))
            if 'Cancel' in status:
                return 'Shipped' if len(track) > 5 and track.lower() != 'nan' else 'Unshipped'
            return 'Other'
        df_unmatched['Check_Status'] = df_unmatched.apply(check_status, axis=1)
        count_cancel_shipped   = len(df_unmatched[df_unmatched['Check_Status'] == 'Shipped'])
        count_cancel_unshipped = len(df_unmatched[df_unmatched['Check_Status'] == 'Unshipped'])

    # ---------- 6. 补单与费用计算 ----------
    brushing_ids = set()
    if not df_brush.empty:
        brush_id_col = next((c for c in df_brush.columns if '单号' in c), df_brush.columns[0])
        brushing_ids = set(df_brush[brush_id_col].astype(str).str.strip())

    c_rebate = next((c for c in cols if 'Rebate' in c and 'Shopee' in c), None)
    cols_brush_amt = [c for c in [c_price, c_promo, c_voucher, c_refund, c_rebate] if c]

    all_cols = list(df_income.columns)
    try:
        idx_start = next(i for i, c in enumerate(all_cols) if 'Seller shipping fee' in c or 'Shipping Fee Paid by Seller' in c)
        idx_end   = next(i for i, c in enumerate(all_cols) if 'Withholding Tax' in c)
        cols_platform_fee_range = all_cols[idx_start: idx_end + 1]
    except:
        cols_platform_fee_range = [c for c in all_cols if 'Fee' in c or 'Tax' in c or 'Commission' in c]

    brush_refund_sum_php = brush_platform_fee_php = 0
    for bid in brushing_ids:
        if bid in income_dict:
            row = income_dict[bid]
            for c in cols_brush_amt:          brush_refund_sum_php  += clean_currency(row.get(c, 0))
            for c in cols_platform_fee_range: brush_platform_fee_php += clean_currency(row.get(c, 0))

    cols_pure_platform   = [c for c in cols_platform_fee_range if 'Tax' not in c]
    total_platform_fee_php = sum(df_income_m[c].apply(clean_currency).sum() for c in cols_pure_platform if c in df_income_m.columns)
    c_tax          = next((c for c in all_cols if 'Withholding Tax' in c), None)
    total_tax_php  = df_income_m[c_tax].apply(clean_currency).sum() if c_tax else 0

    ads_col  = next((c for c in df_ads.columns if '花费' in c or 'Expense' in c), None) if not df_ads.empty else None
    ads_cny  = df_ads[ads_col].apply(clean_currency).sum() * ER if ads_col else 0

    brush_comm = brush_diff = 0
    if not df_brush.empty:
        c_comm = next((c for c in df_brush.columns if '佣金' in c), None)
        c_diff = next((c for c in df_brush.columns if '汇差' in c), None)
        if c_comm: brush_comm = df_brush[c_comm].apply(clean_currency).sum()
        if c_diff: brush_diff = df_brush[c_diff].apply(clean_currency).sum()

    metrics = {
        'total_platform_fee_php': total_platform_fee_php,
        'total_tax_php':          total_tax_php,
        'brush_amt_php':          brush_refund_sum_php,
        'brush_platform_fee_php': brush_platform_fee_php,
        'brush_commission':       brush_comm,
        'brush_diff':             brush_diff,
        'brush_empty_pkg':        len(brushing_ids) * UF,
        'ads_cny':                ads_cny,
        'cancel_shipped':         count_cancel_shipped,
        'cancel_unshipped':       count_cancel_unshipped,
        'order_count':            len(df_sales),
    }

    # ---------- 7. 复制源文件到 Result 目录 ----------
    current_time  = time.strftime("%H点%M分%S秒")
    new_file_name = f"新_{file_name.replace('.xlsx', '')}_{current_time}.xlsx"
    file_output   = os.path.join(result_dir, new_file_name)
    print(f"   生成副本: {new_file_name}")
    try:
        shutil.copy(source_path, file_output)
        os.chmod(file_output, stat.S_IWRITE)
        time.sleep(1)
    except Exception as e:
        print(f"   [错误] 复制失败: {e}")
        return

    # ---------- 8. 写入 Excel ----------
    try:
        wb = openpyxl.load_workbook(file_output)

        # --- A. 销售明细 ---
        sales_sheet_name = None
        for n in SN['sales']:
            if n in wb.sheetnames:
                sales_sheet_name = n
                break
        if not sales_sheet_name:
            ws_sales = wb.create_sheet('销售明细')
            ws_sales.append(['日期', '订单号', '商品SKU', '销售数量', '销售单价', '优惠金额',
                             '优惠劵金额', '退款金额', '销售金额(PHP)', '销售金额(CNY)', '商品成本总额(CNY)', '毛利(CNY)'])
            sales_sheet_name = '销售明细'
        else:
            ws_sales = wb[sales_sheet_name]
            if ws_sales.max_row >= 2:
                ws_sales.delete_rows(2, ws_sales.max_row)

        for i, row in df_sales.iterrows():
            r = i + 2
            ws_sales.cell(r, 1).value  = str(row['日期'])[:10]
            ws_sales.cell(r, 2).value  = row['订单号']
            ws_sales.cell(r, 3).value  = row['商品SKU']
            ws_sales.cell(r, 4).value  = row['销售数量']
            ws_sales.cell(r, 5).value  = row['销售单价']
            ws_sales.cell(r, 6).value  = row['优惠金额']
            ws_sales.cell(r, 7).value  = row['优惠劵金额']
            ws_sales.cell(r, 8).value  = row['退款金额']
            ws_sales.cell(r, 9).value  = f'=E{r}+F{r}+G{r}+H{r}'
            ws_sales.cell(r, 10).value = f'=I{r}*{ER}'
            ws_sales.cell(r, 11).value = row['商品成本总额(CNY)']
            ws_sales.cell(r, 12).value = f'=J{r}-K{r}'
            for c in range(1, 13):
                safe_border(ws_sales, r, c)

        last_sales_row = len(df_sales) + 1

        # --- B. 补单明细 ---
        brush_sheet_name = None
        for n in SN['brush']:
            if n in wb.sheetnames:
                brush_sheet_name = n
                break

        v_shopee_fee   = abs(metrics['total_platform_fee_php']) * ER
        v_tax          = abs(metrics['total_tax_php']) * ER
        v_brush_amt    = metrics['brush_amt_php'] * ER
        v_brush_plat   = abs(metrics['brush_platform_fee_php']) * ER
        brush_comm     = metrics['brush_commission']
        brush_empty_pkg = metrics['brush_empty_pkg']
        brush_diff_expr = 0

        if brush_sheet_name and not df_brush.empty:
            ws_brush = wb[brush_sheet_name]
            header_row = col_idx_xiaoji = col_idx_chengben = col_idx_huicha = None

            for row in ws_brush.iter_rows(min_row=1, max_row=20):
                for cell in row:
                    if cell.value:
                        val_str = str(cell.value).strip()
                        if '垫付小计' in val_str:   col_idx_xiaoji  = cell.column
                        elif '垫付成本' in val_str:  col_idx_chengben = cell.column
                        elif '汇率差' in val_str or '汇差' in val_str: col_idx_huicha = cell.column
                if col_idx_xiaoji and col_idx_chengben:
                    header_row = row[0].row
                    break

            if header_row and col_idx_xiaoji and col_idx_chengben:
                if not col_idx_huicha:
                    col_idx_huicha = ws_brush.max_column + 1
                    ws_brush.cell(row=header_row, column=col_idx_huicha, value="汇率差")

                col_let_xj = get_column_letter(col_idx_xiaoji)
                col_let_cb = get_column_letter(col_idx_chengben)
                col_let_hc = get_column_letter(col_idx_huicha)
                max_brush_row = ws_brush.max_row

                for r in range(header_row + 1, max_brush_row + 1):
                    c_xj = ws_brush.cell(row=r, column=col_idx_xiaoji)
                    c_cb = ws_brush.cell(row=r, column=col_idx_chengben)
                    for c in [c_xj, c_cb]:
                        if c.value:
                            cleaned = str(c.value).replace('₱', '').replace('PHP', '').replace(',', '').strip()
                            try:
                                c.value = float(cleaned)
                            except:
                                pass
                    if c_xj.value is not None and c_cb.value is not None:
                        ws_brush.cell(row=r, column=col_idx_huicha).value = f'={col_let_xj}{r}-{ER}*{col_let_cb}{r}'

                brush_diff_expr = f"=SUM('{brush_sheet_name}'!{col_let_hc}{header_row+1}:{col_let_hc}{max_brush_row})"
        else:
            v_brush_amt = brush_comm = v_brush_plat = brush_empty_pkg = 0
            brush_diff_expr = 0

        # --- C. 利润表（统一动态月份列写入） ---
        pnl_sheet_name = None
        for n in SN['pnl']:
            if n in wb.sheetnames:
                pnl_sheet_name = n
                break

        if pnl_sheet_name:
            ws_pnl = wb[pnl_sheet_name]
            target_col_idx = find_month_col(ws_pnl, TM)
            col_letter     = get_column_letter(target_col_idx)

            val_map = {
                '本期销售':     f'=SUM({sales_sheet_name}!J2:J{last_sales_row})',
                '其中:补单金额': v_brush_amt,
                '销售退款':     f'=ABS(SUM({sales_sheet_name}!H2:H{last_sales_row}))*{ER}',
                '实际产品成本': f'=SUM({sales_sheet_name}!K2:K{last_sales_row})',
                '发单费':       f'=({metrics["order_count"]}+{metrics["cancel_shipped"]})*{UF}',
                '广告费':       metrics['ads_cny'],
                '补单佣金':     brush_comm,
                '补单平台费用': v_brush_plat,
                '补单汇率差额': brush_diff_expr,
                '补单空包费':   brush_empty_pkg,
                '平台费用':     v_shopee_fee,
                '预扣税':       v_tax,
                '团队运营费用': 0,
            }

            # 按长度从长到短排列关键词，防止短词错误匹配长词
            search_keys = [
                '实际主营收入', '其中:补单金额', '实际主营成本', '实际产品成本',
                '团队运营费用', '补单平台费用', '补单汇率差额', '补单空包费',
                'shopee结算', '本期销售', '销售退款', '营业费用', '推广费用',
                '补单佣金', '平台费用', '毛利润', '本期利润', '发单费', '广告费', '预扣税'
            ]

            row_idx_map = {}
            for row in ws_pnl.iter_rows(min_row=1, max_row=100):
                cell_a    = row[0]
                cell_val  = str(cell_a.value).strip() if cell_a.value else ""
                for key in search_keys:
                    if key in cell_val:
                        row_idx_map[key] = cell_a.row
                        if key in val_map:
                            ws_pnl.cell(row=cell_a.row, column=target_col_idx).value = val_map[key]
                        break
                if "未匹配" in cell_val or "补充" in cell_val:
                    cell_a.value = f"未匹配说明：取消已发货 {metrics['cancel_shipped']} 单，取消未发货 {metrics['cancel_unshipped']} 单"

            def get_cell(key):
                return f"{col_letter}{row_idx_map.get(key, 0)}"

            if '实际主营收入' in row_idx_map:
                ws_pnl.cell(row_idx_map['实际主营收入'], target_col_idx).value = \
                    f"={get_cell('本期销售')}-{get_cell('其中:补单金额')}-{get_cell('销售退款')}"
            if '实际主营成本' in row_idx_map:
                ws_pnl.cell(row_idx_map['实际主营成本'], target_col_idx).value = \
                    f"={get_cell('实际产品成本')}+{get_cell('发单费')}"
            if '推广费用' in row_idx_map:
                sub = [get_cell(k) for k in ['广告费', '补单佣金', '补单平台费用', '补单汇率差额', '补单空包费'] if k in row_idx_map]
                ws_pnl.cell(row_idx_map['推广费用'], target_col_idx).value = f"={'+'.join(sub)}"
            if 'shopee结算' in row_idx_map:
                ws_pnl.cell(row_idx_map['shopee结算'], target_col_idx).value = \
                    f"={get_cell('平台费用')}+{get_cell('预扣税')}"
            if '营业费用' in row_idx_map:
                ws_pnl.cell(row_idx_map['营业费用'], target_col_idx).value = \
                    f"={get_cell('推广费用')}+{get_cell('shopee结算')}+{get_cell('团队运营费用')}"
            profit_key = '本期利润' if '本期利润' in row_idx_map else '毛利润'
            if profit_key in row_idx_map:
                ws_pnl.cell(row_idx_map[profit_key], target_col_idx).value = \
                    f"={get_cell('实际主营收入')}-{get_cell('实际主营成本')}-{get_cell('营业费用')}"

        # --- D. 未匹配订单 ---
        un_sheet_name = None
        for n in SN['unmatched']:
            if n in wb.sheetnames:
                un_sheet_name = n
                break
        if not df_unmatched.empty:
            if un_sheet_name:
                ws_un = wb[un_sheet_name]
                if ws_un.max_row >= 1:
                    ws_un.delete_rows(1, ws_un.max_row)
            else:
                ws_un = wb.create_sheet('未匹配订单')
            ws_un.append(df_unmatched.columns.tolist())
            for r in dataframe_to_rows(df_unmatched, index=False, header=False):
                ws_un.append(r)

        wb.save(file_output)
        wb.close()
        print("   [成功] 处理完成")

    except Exception as e:
        print(f"   [失败] 处理出错: {e}")


# ================================================================
# ============== 第二段 B：TK无马帮 平台处理函数（占位符） ========
# ================================================================
# （TK无马帮函数定义在文件末尾，Python允许在main调用前定义）

# ================================================================
# ==================== 第三段：主入口与平台选择 ==================
# ================================================================

def run_shopee(params: dict):
    """运行 Shopee 平台批量处理，params 由 GUI 传入"""
    data_dir   = params['data_dir']
    result_dir = params['result_dir']

    # 将 GUI 参数注入全局变量（供 process_shopee 读取）
    global SHOPEE_EXCHANGE_RATE, SHOPEE_UNIT_FEE, SHOPEE_TARGET_YEAR, SHOPEE_TARGET_MONTH
    SHOPEE_EXCHANGE_RATE = params['exchange_rate']
    SHOPEE_UNIT_FEE      = params['unit_fee']
    SHOPEE_TARGET_YEAR   = params['target_year']
    SHOPEE_TARGET_MONTH  = params['target_month']

    print(f"\n启动 Shopee 批量处理...")
    print(f"数据目录: {data_dir}")
    print(f"结果目录: {result_dir}")
    print(f"分析时间: {SHOPEE_TARGET_YEAR}年 {SHOPEE_TARGET_MONTH}月")
    print(f"汇率: {SHOPEE_EXCHANGE_RATE}  发单费: {SHOPEE_UNIT_FEE}")

    if not os.path.exists(data_dir):
        print(f"错误: 找不到数据文件夹 '{data_dir}'")
        return
    os.makedirs(result_dir, exist_ok=True)

    cost_path = os.path.join(data_dir, COST_FILE_NAME)
    if not os.path.exists(cost_path):
        print(f"错误: 找不到成本表 '{cost_path}'")
        return

    print(">>> 加载商品成本表...")
    try:
        df_cost  = pd.read_excel(cost_path, dtype=str)
        sku_col  = next((c for c in df_cost.columns if 'sku' in c.lower()), 'SKU')
        cost_col = next((c for c in df_cost.columns if '成本' in c), '人民币成本')
        cost_map = dict(zip(df_cost[sku_col].astype(str).str.strip(), df_cost[cost_col].apply(clean_currency)))
        print(f"   成本表加载成功，共 {len(cost_map)} 条SKU记录")
    except Exception as e:
        print(f"   读取成本表失败: {e}")
        return

    xlsx_files = glob.glob(os.path.join(data_dir, "*.xlsx"))
    count = 0
    for file_path in xlsx_files:
        file_name = os.path.basename(file_path)
        if file_name == COST_FILE_NAME or "新_" in file_name or file_name.startswith("~$"):
            continue
        process_shopee(file_path, cost_map, result_dir)
        count += 1

    print(f"\n{'='*60}")
    print(f"Shopee 批量处理完成！共处理 {count} 个文件。")
    print(f"新生成的文件存放在: {result_dir}")


def run_tk_no_horse(params: dict):
    """运行 TK无马帮 平台批量处理，params 由 GUI 传入"""
    data_dir   = params['data_dir']
    result_dir = params['result_dir']

    global TK_NOHORSE_EXCHANGE_RATE, TK_NOHORSE_UNIT_FEE, TK_NOHORSE_TARGET_YEAR, TK_NOHORSE_TARGET_MONTH
    TK_NOHORSE_EXCHANGE_RATE = params['exchange_rate']
    TK_NOHORSE_UNIT_FEE      = params['unit_fee']
    TK_NOHORSE_TARGET_YEAR   = params['target_year']
    TK_NOHORSE_TARGET_MONTH  = params['target_month']

    print(f"\n启动 TK无马帮 批量处理...")
    print(f"数据目录: {data_dir}")
    print(f"结果目录: {result_dir}")
    print(f"分析时间: {TK_NOHORSE_TARGET_YEAR}年 {TK_NOHORSE_TARGET_MONTH}月")
    print(f"汇率: {TK_NOHORSE_EXCHANGE_RATE}  发单费: {TK_NOHORSE_UNIT_FEE}")

    if not os.path.exists(data_dir):
        print(f"错误: 找不到数据文件夹 '{data_dir}'")
        return
    os.makedirs(result_dir, exist_ok=True)

    cost_path = os.path.join(data_dir, COST_FILE_NAME)
    if not os.path.exists(cost_path):
        print(f"错误: 找不到成本表 '{cost_path}'")
        return

    print(">>> 加载商品成本表...")
    try:
        df_cost  = pd.read_excel(cost_path, dtype=str)
        sku_col  = next((c for c in df_cost.columns if 'sku' in c.lower()), 'SKU')
        cost_col = next((c for c in df_cost.columns if '成本' in c), '人民币成本')
        cost_map = dict(zip(df_cost[sku_col].astype(str).str.strip(), df_cost[cost_col].apply(clean_currency)))
        print(f"   成本表加载成功，共 {len(cost_map)} 条SKU记录")
    except Exception as e:
        print(f"   读取成本表失败: {e}")
        return

    xlsx_files = glob.glob(os.path.join(data_dir, "*.xlsx"))
    count = 0
    for file_path in xlsx_files:
        file_name = os.path.basename(file_path)
        if file_name == COST_FILE_NAME or "新_" in file_name or file_name.startswith("~$"):
            continue
        process_tk_no_horse(file_path, cost_map, result_dir)
        count += 1

    print(f"\n{'='*60}")
    print(f"TK无马帮 批量处理完成！共处理 {count} 个文件。")
    print(f"新生成的文件存放在: {result_dir}")


def run_tk_horse(params: dict):
    """运行 TK有马帮 平台批量处理，params 由 GUI 传入"""
    data_dir   = params['data_dir']
    result_dir = params['result_dir']

    global TK_HORSE_TARGET_MONTH
    TK_HORSE_TARGET_MONTH = params['target_month']

    print(f"\n启动 TK有马帮 批量处理...")
    print(f"数据目录: {data_dir}")
    print(f"结果目录: {result_dir}")
    print(f"目标月份: {TK_HORSE_TARGET_MONTH}月")

    if not os.path.exists(data_dir):
        print(f"错误: 找不到数据文件夹 '{data_dir}'")
        return
    os.makedirs(result_dir, exist_ok=True)

    xlsx_files = glob.glob(os.path.join(data_dir, "*.xlsx"))
    count = 0
    for file_path in xlsx_files:
        file_name = os.path.basename(file_path)
        if "新_" in file_name or file_name.startswith("~$"):
            continue
        process_tk_horse(file_path, result_dir)
        count += 1

    print(f"\n{'='*60}")
    print(f"TK有马帮 批量处理完成！共处理 {count} 个文件。")
    print(f"新生成的文件存放在: {result_dir}")


# ================================================================
# ==================== 第四段：GUI 主类 =========================
# ================================================================

class AutoExcelApp(ctk.CTk):
    """Auto Excel 批量报表处理工具 - GUI 主窗口"""

    # 各平台默认目录映射
    _DEFAULT_DIRS = {
        'Shopee':   (SHOPEE_DATA_DIR,     SHOPEE_RESULT_DIR),
        'TK无马帮': (TK_NOHORSE_DATA_DIR, TK_NOHORSE_RESULT_DIR),
        'TK有马帮': (TK_HORSE_DATA_DIR,   TK_HORSE_RESULT_DIR),
    }

    def __init__(self):
        super().__init__()
        self.title("Auto Excel 批量报表处理工具")
        self.geometry("780x800")
        self.resizable(False, False)
        self.configure(fg_color="#F5F5F7")

        # 当前选中平台
        self._platform = tk.StringVar(value="Shopee")

        self._build_title()
        self._build_platform_switcher()
        self._build_params_area()
        self._build_dir_area()
        self._build_console()
        self._build_run_button()

        # 初始化显示
        self._on_platform_change("Shopee")

    # ── 标题区 ──────────────────────────────────────────────────
    def _build_title(self):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="x", padx=36, pady=(28, 0))

        ctk.CTkLabel(
            frame,
            text="Auto Excel",
            font=ctk.CTkFont(family="Arial", size=26, weight="bold"),
            text_color="#1D1D1F",
        ).pack(side="left")

        ctk.CTkLabel(
            frame,
            text="  批量报表处理工具",
            font=ctk.CTkFont(family="Arial", size=16),
            text_color="#6E6E73",
        ).pack(side="left", pady=(6, 0))

    # ── 平台切换区 ───────────────────────────────────────────────
    def _build_platform_switcher(self):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="x", padx=36, pady=(18, 0))

        ctk.CTkLabel(
            frame,
            text="选择平台",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#1D1D1F",
        ).pack(anchor="w", pady=(0, 8))

        self._seg_btn = ctk.CTkSegmentedButton(
            frame,
            values=["Shopee", "TK无马帮", "TK有马帮"],
            variable=self._platform,
            command=self._on_platform_change,
            font=ctk.CTkFont(size=13),
            height=36,
            corner_radius=10,
            selected_color="#0071E3",
            selected_hover_color="#0077ED",
            unselected_color="#E8E8ED",
            unselected_hover_color="#D8D8DD",
            text_color="#FFFFFF",
            text_color_disabled="#6E6E73",
        )
        self._seg_btn.pack(fill="x")

    # ── 参数输入区（动态显示/隐藏） ──────────────────────────────
    def _build_params_area(self):
        """构建参数输入卡片"""
        self._params_card = ctk.CTkFrame(
            self, fg_color="#FFFFFF", corner_radius=12,
            border_width=1, border_color="#E0E0E5"
        )
        self._params_card.pack(fill="x", padx=36, pady=(14, 0))

        ctk.CTkLabel(
            self._params_card,
            text="参数配置",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#1D1D1F",
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=16, pady=(12, 6))

        # ---- 年份 / 月份 ----
        self._year_var  = tk.StringVar(value="2026")
        self._month_var = tk.StringVar(value="1")

        self._lbl_year = ctk.CTkLabel(self._params_card, text="目标年份", text_color="#6E6E73", font=ctk.CTkFont(size=12))
        self._lbl_year.grid(row=1, column=0, sticky="w", padx=(16, 4), pady=4)
        self._ent_year = ctk.CTkEntry(self._params_card, textvariable=self._year_var, width=130, height=34, corner_radius=8)
        self._ent_year.grid(row=1, column=1, sticky="w", padx=(0, 20), pady=4)

        self._lbl_month = ctk.CTkLabel(self._params_card, text="目标月份", text_color="#6E6E73", font=ctk.CTkFont(size=12))
        self._lbl_month.grid(row=1, column=2, sticky="w", padx=(0, 4), pady=4)
        self._ent_month = ctk.CTkEntry(self._params_card, textvariable=self._month_var, width=130, height=34, corner_radius=8)
        self._ent_month.grid(row=1, column=3, sticky="w", padx=(0, 16), pady=4)

        # ---- 汇率 / 发单费（Shopee & TK无马帮 专用） ----
        self._rate_var = tk.StringVar(value="0.117789")
        self._fee_var  = tk.StringVar(value="3.0")

        self._lbl_rate = ctk.CTkLabel(self._params_card, text="汇率", text_color="#6E6E73", font=ctk.CTkFont(size=12))
        self._lbl_rate.grid(row=2, column=0, sticky="w", padx=(16, 4), pady=(4, 12))
        self._ent_rate = ctk.CTkEntry(self._params_card, textvariable=self._rate_var, width=130, height=34, corner_radius=8)
        self._ent_rate.grid(row=2, column=1, sticky="w", padx=(0, 20), pady=(4, 12))

        self._lbl_fee = ctk.CTkLabel(self._params_card, text="发单费单价(CNY)", text_color="#6E6E73", font=ctk.CTkFont(size=12))
        self._lbl_fee.grid(row=2, column=2, sticky="w", padx=(0, 4), pady=(4, 12))
        self._ent_fee = ctk.CTkEntry(self._params_card, textvariable=self._fee_var, width=130, height=34, corner_radius=8)
        self._ent_fee.grid(row=2, column=3, sticky="w", padx=(0, 16), pady=(4, 12))

    # ── 目录选择区 ───────────────────────────────────────────────
    def _build_dir_area(self):
        card = ctk.CTkFrame(
            self, fg_color="#FFFFFF", corner_radius=12,
            border_width=1, border_color="#E0E0E5"
        )
        card.pack(fill="x", padx=36, pady=(14, 0))

        ctk.CTkLabel(
            card, text="目录设置",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#1D1D1F",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(12, 6))

        # Data 目录
        self._data_dir_var = tk.StringVar()
        ctk.CTkLabel(card, text="📁  Data 目录", text_color="#6E6E73", font=ctk.CTkFont(size=12)).grid(
            row=1, column=0, sticky="w", padx=(16, 8), pady=4)
        ctk.CTkEntry(card, textvariable=self._data_dir_var, width=460, height=32, corner_radius=8,
                     state="normal").grid(row=1, column=1, sticky="w", pady=4)
        ctk.CTkButton(card, text="选择", width=60, height=32, corner_radius=8,
                      fg_color="#E8E8ED", text_color="#1D1D1F", hover_color="#D0D0D5",
                      command=self._pick_data_dir).grid(row=1, column=2, padx=(8, 16), pady=4)

        # Result 目录
        self._result_dir_var = tk.StringVar()
        ctk.CTkLabel(card, text="📁  Result 目录", text_color="#6E6E73", font=ctk.CTkFont(size=12)).grid(
            row=2, column=0, sticky="w", padx=(16, 8), pady=(4, 12))
        ctk.CTkEntry(card, textvariable=self._result_dir_var, width=460, height=32, corner_radius=8,
                     state="normal").grid(row=2, column=1, sticky="w", pady=(4, 12))
        ctk.CTkButton(card, text="选择", width=60, height=32, corner_radius=8,
                      fg_color="#E8E8ED", text_color="#1D1D1F", hover_color="#D0D0D5",
                      command=self._pick_result_dir).grid(row=2, column=2, padx=(8, 16), pady=(4, 12))

    # ── 控制台输出区 ─────────────────────────────────────────────
    def _build_console(self):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=36, pady=(14, 0))

        ctk.CTkLabel(
            frame, text="运行日志",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#1D1D1F",
        ).pack(anchor="w", pady=(0, 6))

        self._console = ctk.CTkTextbox(
            frame,
            height=180,
            corner_radius=10,
            font=ctk.CTkFont(family="Consolas", size=12),
            fg_color="#1C1C1E",
            text_color="#E5E5EA",
            border_width=0,
            wrap="word",
            state="disabled",
        )
        self._console.pack(fill="both", expand=True)

    # ── 执行按钮 ─────────────────────────────────────────────────
    def _build_run_button(self):
        self._run_btn = ctk.CTkButton(
            self,
            text="🚀  开始处理",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=48,
            corner_radius=12,
            fg_color="#0071E3",
            hover_color="#0077ED",
            command=self._on_run,
        )
        self._run_btn.pack(fill="x", padx=36, pady=(14, 28))

    # ── 平台切换回调 ─────────────────────────────────────────────
    def _on_platform_change(self, platform: str):
        """切换平台时更新默认目录，并显示/隐藏汇率&发单费行"""
        data_dir, result_dir = self._DEFAULT_DIRS.get(platform, ("", ""))
        self._data_dir_var.set(data_dir)
        self._result_dir_var.set(result_dir)

        # 汇率和发单费仅 Shopee / TK无马帮 可用；年份所有平台均可填写
        show_rate_fee = (platform != "TK有马帮")
        state_rf = "normal" if show_rate_fee else "disabled"
        fg_rf    = "#6E6E73" if show_rate_fee else "#AEAEB2"

        for w in [self._lbl_rate, self._lbl_fee, self._ent_rate, self._ent_fee]:
            if isinstance(w, ctk.CTkEntry):
                w.configure(state=state_rf)
            else:
                w.configure(text_color=fg_rf)

        # 年份始终可编辑，颜色保持正常
        self._lbl_year.configure(text_color="#6E6E73")
        self._ent_year.configure(state="normal")

    # ── 目录选择回调 ─────────────────────────────────────────────
    def _pick_data_dir(self):
        d = filedialog.askdirectory(title="选择 Data 目录", initialdir=self._data_dir_var.get())
        if d:
            self._data_dir_var.set(d)

    def _pick_result_dir(self):
        d = filedialog.askdirectory(title="选择 Result 目录", initialdir=self._result_dir_var.get())
        if d:
            self._result_dir_var.set(d)

    # ── 控制台写入（线程安全） ────────────────────────────────────
    def _console_write(self, text: str):
        self._console.configure(state="normal")
        self._console.insert("end", text)
        self._console.see("end")
        self._console.configure(state="disabled")

    def _console_clear(self):
        self._console.configure(state="normal")
        self._console.delete("1.0", "end")
        self._console.configure(state="disabled")

    # ── 执行按钮回调 ─────────────────────────────────────────────
    def _on_run(self):
        """收集参数，在后台线程中执行处理逻辑"""
        platform = self._platform.get()

        # 参数校验
        try:
            month = int(self._month_var.get())
            assert 1 <= month <= 12
        except Exception:
            self._console_write("❌ 错误：目标月份必须是 1~12 的整数\n")
            return

        params = {
            'data_dir':    self._data_dir_var.get().strip(),
            'result_dir':  self._result_dir_var.get().strip(),
            'target_month': month,
        }

        try:
            params['target_year'] = int(self._year_var.get())
        except Exception:
            self._console_write("❌ 错误：目标年份必须是整数\n")
            return

        if platform != "TK有马帮":
            try:
                params['exchange_rate'] = float(self._rate_var.get())
                params['unit_fee']      = float(self._fee_var.get())
            except Exception:
                self._console_write("❌ 错误：汇率/发单费格式不正确\n")
                return

        self._console_clear()
        self._run_btn.configure(state="disabled", text="处理中…")

        # 重定向 stdout → 控制台
        class _GuiWriter(io.TextIOBase):
            def __init__(self_, write_fn):
                self_._write_fn = write_fn
            def write(self_, s):
                if s:
                    self.after(0, self._console_write, s)
                return len(s)
            def flush(self_):
                pass

        self._old_stdout = sys.stdout
        sys.stdout = _GuiWriter(self._console_write)

        def _worker():
            try:
                if platform == "Shopee":
                    run_shopee(params)
                elif platform == "TK无马帮":
                    run_tk_no_horse(params)
                else:
                    run_tk_horse(params)
            except Exception as e:
                print(f"\n❌ 未捕获异常: {e}")
                import traceback
                traceback.print_exc()
            finally:
                sys.stdout = self._old_stdout
                self.after(0, self._run_btn.configure, {"state": "normal", "text": "🚀  开始处理"})

        threading.Thread(target=_worker, daemon=True).start()


def main():
    """启动 GUI 界面"""
    app = AutoExcelApp()
    app.mainloop()


if __name__ == '__main__':
    main()


# ================================================================
# ============== 第二段 C：TK有马帮 平台处理函数 =================
# ================================================================

def process_tk_horse(source_path, result_dir):
    """TK有马帮平台单文件处理：账单+订单明细+马帮表匹配，动态月份列写利润表"""
    file_name = os.path.basename(source_path)
    TM = TK_HORSE_TARGET_MONTH

    print(f"\n{'='*60}")
    print(f">>> [TK有马帮] 正在处理: {file_name}")

    try:
        xl_dict = pd.read_excel(source_path, sheet_name=None, dtype=str, header=None)
    except Exception as e:
        print(f"   [失败] 读取错误: {e}")
        return

    # ---------- 1. 映射 Sheet 名称 ----------
    sheet_map = {}
    for sheet_name in xl_dict.keys():
        name_lower = sheet_name.lower()
        if "账单" in name_lower:      sheet_map['bill']      = sheet_name
        elif "订单明细" in name_lower: sheet_map['detail']    = sheet_name
        elif "马帮" in name_lower:     sheet_map['mabang']    = sheet_name
        elif "广告" in name_lower:     sheet_map['ads']       = sheet_name
        elif "利润" in name_lower:     sheet_map['pnl']       = sheet_name
        elif "未匹配" in name_lower:   sheet_map['unmatched'] = sheet_name
        elif "销售明细" in name_lower: sheet_map['sales']     = sheet_name

    if not sheet_map.get('bill') or not sheet_map.get('detail') or not sheet_map.get('mabang'):
        print("   [错误] 缺少必要Sheet (账单/订单明细/马帮)")
        return

    try:
        df_mabang = pd.read_excel(source_path, sheet_name=sheet_map['mabang'], dtype=str)
        df_bill   = pd.read_excel(source_path, sheet_name=sheet_map['bill'],   dtype=str)
        df_detail = pd.read_excel(source_path, sheet_name=sheet_map['detail'], dtype=str)
        df_ads    = pd.read_excel(source_path, sheet_name=sheet_map['ads'],    dtype=str) if sheet_map.get('ads') else None
    except Exception as e:
        print(f"   [错误] 读取Sheet内容失败: {e}")
        return

    print(f"   [2/5] 正在计算...")

    # ---------- 2. 马帮：计算平均汇率和成本字典 ----------
    col_mb_rate = next((c for c in df_mabang.columns if '汇率' in str(c)), None)
    col_mb_cost = next((c for c in df_mabang.columns if '支出-成本' in str(c)), None)
    col_mb_oid  = df_mabang.columns[0]

    rates    = df_mabang[col_mb_rate].apply(safe_float) if col_mb_rate else pd.Series([])
    avg_rate = rates[rates > 0].mean() if len(rates) > 0 else 1.0
    if pd.isna(avg_rate) or avg_rate == 0:
        avg_rate = 1.0
    print(f"   平均汇率: {avg_rate:.6f}")

    df_mabang['clean_oid'] = df_mabang[col_mb_oid].astype(str).str.strip().str.replace('\t', '')
    df_mabang['cost_val']  = df_mabang[col_mb_cost].apply(safe_float) if col_mb_cost else 0
    cost_dict = df_mabang.groupby('clean_oid')['cost_val'].sum().to_dict()

    # ---------- 3. 广告费 ----------
    ads_cost_cny = 0
    if df_ads is not None:
        for col in df_ads.columns:
            sum_col = sum(safe_float(v) for v in df_ads[col] if safe_float(v) > 0)
            if sum_col > 0:
                ads_cost_cny = sum_col * 7.1   # 广告费汇率按 7.1
                break

    # ---------- 4. 账单处理 ----------
    col_bill_id    = next((c for c in df_bill.columns if 'Order/adjustment ID' in str(c) or 'Order ID' in str(c)), df_bill.columns[0])
    col_bill_type  = next((c for c in df_bill.columns if 'Type' in str(c)), 'Type')
    col_created    = next((c for c in df_bill.columns if 'created time' in str(c)), 'Order created time')
    col_total_fees = next((c for c in df_bill.columns if 'Total Fees' in str(c)), None)
    col_adj        = next((c for c in df_bill.columns if 'Ajustment amount' in str(c) or 'Adjustment amount' in str(c)), None)

    df_bill['clean_oid']  = df_bill[col_bill_id].astype(str).str.strip()
    df_bill_orders        = df_bill[df_bill[col_bill_type].astype(str).str.lower() == 'order'].copy()
    df_bill_unique        = df_bill_orders.drop_duplicates(subset=['clean_oid'], keep='first')
    print(f"   账单内有效订单数(Type=Order): {len(df_bill_unique)}")

    # ---------- 5. 调整金额（按目标月份筛选） ----------
    total_adjustment_idr = 0
    if col_adj and col_created:
        temp_dates = pd.to_datetime(df_bill[col_created], errors='coerce')
        mask_month = temp_dates.dt.month == TM
        total_adjustment_idr = df_bill.loc[mask_month, col_adj].apply(safe_float).sum()
        print(f"   {TM}月份调整金额总和(IDR): {total_adjustment_idr}")

    # ---------- 6. 未匹配处理 ----------
    billed_oids    = set(df_bill_orders['clean_oid'])
    col_detail_id  = next((c for c in df_detail.columns if 'Order ID' in str(c)), df_detail.columns[0])
    df_detail['clean_oid'] = df_detail[col_detail_id].astype(str).str.strip()
    df_unmatched   = df_detail[~df_detail['clean_oid'].isin(billed_oids)].copy()

    col_ship = next((c for c in df_unmatched.columns if 'Shipped Time' in str(c)), None)
    unmatched_shipped_count = 0
    unmatched_shipped_oids  = []

    if col_ship:
        for _, row in df_unmatched.iterrows():
            st = str(row[col_ship])
            if len(st) > 5 and 'nan' not in st.lower():
                unmatched_shipped_count += 1
                unmatched_shipped_oids.append(row['clean_oid'])
    else:
        unmatched_shipped_count = len(df_unmatched)
        unmatched_shipped_oids  = df_unmatched['clean_oid'].tolist()

    # ---------- 7. 生成销售明细 & 平台费计算 ----------
    sales_rows             = []
    total_platform_fee_idr = 0

    for _, bill_row in df_bill_unique.iterrows():
        oid     = bill_row['clean_oid']
        matches = df_detail[df_detail['clean_oid'] == oid]
        if matches.empty:
            continue

        if col_total_fees:
            total_platform_fee_idr += safe_float(bill_row.get(col_total_fees))

        t_sales  = safe_float(bill_row.get('Subtotal before discounts'))
        t_disc   = safe_float(bill_row.get('Seller discounts'))
        t_refund = safe_float(bill_row.get('Refund subtotal after seller discounts'))
        t_cost   = cost_dict.get(oid, 0)

        sku_parts = []
        qty_total = 0
        for _, d_row in matches.iterrows():
            s_sku = str(d_row.get('Seller SKU', '')).strip()
            s_qty = safe_float(d_row.get('Quantity', 0))
            fmt_qty = int(s_qty) if s_qty == int(s_qty) else s_qty
            sku_parts.append(f"{s_sku}*{fmt_qty}")
            qty_total += s_qty
        sku_info = ",\n".join(sku_parts)

        sales_rows.append({
            'date':      bill_row.get(col_created),
            'oid':       oid,
            'sku':       sku_info,
            'qty':       qty_total,
            'sales_idr': t_sales,
            'disc_idr':  t_disc,
            'refund_idr': t_refund,
            'cost_cny':  t_cost,
        })

    normal_shipped_count = len(sales_rows)
    sum_cost_cny         = sum(r['cost_cny'] for r in sales_rows)

    platform_fee_cny  = abs(total_platform_fee_idr * avg_rate)
    adjustment_cny    = abs(total_adjustment_idr * avg_rate)
    tk_settlement_cny = platform_fee_cny + adjustment_cny

    total_shipped_count = normal_shipped_count + unmatched_shipped_count
    ship_fee            = total_shipped_count * 3.5

    unmatched_cost_total = sum(cost_dict.get(str(oid), 0) for oid in unmatched_shipped_oids)

    # ---------- 8. 复制源文件到 Result 目录 ----------
    print(f"   [3/5] 写入Excel...")
    current_time  = time.strftime("%H点%M分%S秒")
    new_file_name = f"新_{file_name.replace('.xlsx', '')}_{current_time}.xlsx"
    output_path   = os.path.join(result_dir, new_file_name)

    try:
        shutil.copy(source_path, output_path)
        os.chmod(output_path, stat.S_IWRITE)
        time.sleep(1)

        wb = openpyxl.load_workbook(output_path)

        # --- A. 写入销售明细 ---
        if sheet_map.get('sales') and sheet_map['sales'] in wb.sheetnames:
            ws_sales = wb[sheet_map['sales']]
            if ws_sales.max_row >= 2:
                ws_sales.delete_rows(2, ws_sales.max_row)

            bold_border = Side(border_style="medium", color="000000")
            full_border = Border(left=bold_border, right=bold_border, top=bold_border, bottom=bold_border)

            for idx, item in enumerate(sales_rows):
                r = idx + 2
                ws_sales.cell(r, 1).value  = item['date']
                ws_sales.cell(r, 2).value  = item['oid']
                ws_sales.cell(r, 3).value  = item['sku']
                ws_sales.cell(r, 4).value  = item['qty']
                ws_sales.cell(r, 5).value  = item['sales_idr']
                ws_sales.cell(r, 6).value  = item['disc_idr']
                ws_sales.cell(r, 7).value  = item['refund_idr']
                ws_sales.cell(r, 8).value  = f"=E{r}+F{r}+G{r}"
                ws_sales.cell(r, 9).value  = f"=H{r}*{avg_rate:.6f}"
                ws_sales.cell(r, 10).value = item['cost_cny']
                ws_sales.cell(r, 11).value = f"=I{r}-J{r}"
                for c_idx in range(1, 12):
                    ws_sales.cell(r, c_idx).border = full_border

        last_sales_row = len(sales_rows) + 1 if sales_rows else 2

        # --- B. 写入利润表（统一动态月份列写入） ---
        if sheet_map.get('pnl') and sheet_map['pnl'] in wb.sheetnames:
            ws_pnl         = wb[sheet_map['pnl']]
            target_col_idx = find_month_col(ws_pnl, TM)
            col_letter     = get_column_letter(target_col_idx)

            # 构建行关键词→行号映射
            row_map = {}
            for row in ws_pnl.iter_rows(min_row=1, max_col=1):
                cell_val = str(row[0].value).replace(' ', '').replace('\t', '') if row[0].value else ""
                if cell_val:
                    row_map[cell_val] = row[0].row

            def write_pnl(key, value=None, formula=None):
                """写入利润表指定关键词所在行的目标月份列"""
                for k_map, r_idx in row_map.items():
                    if key in k_map:
                        cell = ws_pnl.cell(row=r_idx, column=target_col_idx)
                        cell.value = formula if formula else safe_float(value)
                        return r_idx
                return None

            sales_sheet = sheet_map.get('sales', '销售明细')

            # 1. 收入部分
            formula_sales = f"=SUM({sales_sheet}!I2:I{last_sales_row})+ABS(SUM({sales_sheet}!G2:G{last_sales_row})*{avg_rate:.6f})"
            r_sales   = write_pnl('本期销售',  formula=formula_sales)
            r_refund  = write_pnl('销售退款',  formula=f"=ABS(SUM({sales_sheet}!G2:G{last_sales_row})*{avg_rate:.6f})")

            r_rev = None
            for k, r in row_map.items():
                if '实际主营收入' in k:
                    r_rev = r
                    break
            if r_rev and r_sales and r_refund:
                ws_pnl.cell(r_rev, target_col_idx).value = f"={col_letter}{r_sales}-{col_letter}{r_refund}"

            # 2. 成本部分
            r_prod_cost = write_pnl('实际产品成本', value=sum_cost_cny + unmatched_cost_total)
            r_ship_fee  = write_pnl('发单费',       value=ship_fee)
            if r_ship_fee:
                comment_text = f"正常发货: {normal_shipped_count} + 取消已发货: {unmatched_shipped_count}"
                ws_pnl.cell(r_ship_fee, target_col_idx).comment = Comment(comment_text, "System")

            r_main_cost = None
            for k, r in row_map.items():
                if '实际主营成本' in k and '产品' not in k:
                    r_main_cost = r
                    break
            if r_main_cost and r_prod_cost and r_ship_fee:
                ws_pnl.cell(r_main_cost, target_col_idx).value = f"={col_letter}{r_prod_cost}+{col_letter}{r_ship_fee}"

            # 3. 费用部分
            r_ads  = write_pnl('广告费', value=ads_cost_cny)
            r_promo = None
            for k, r in row_map.items():
                if '推广费用' in k:
                    r_promo = r
                    break
            if r_promo and r_ads:
                ws_pnl.cell(r_promo, target_col_idx).value = f"={col_letter}{r_ads}"

            r_plat = write_pnl('平台费用', value=platform_fee_cny)
            if not r_plat:
                r_plat = write_pnl('平台费', value=platform_fee_cny)
            r_adj  = write_pnl('调整金额', value=adjustment_cny)

            r_tk_settle = None
            for k, r in row_map.items():
                if 'TK结算' in k or 'shopee结算' in k:
                    r_tk_settle = r
                    break
            if r_tk_settle and r_plat and r_adj:
                ws_pnl.cell(r_tk_settle, target_col_idx).value = f"={col_letter}{r_plat}+{col_letter}{r_adj}"

            r_op_exp = None
            for k, r in row_map.items():
                if '营业费用' in k and '推广' not in k:
                    r_op_exp = r
                    break
            if r_op_exp and r_promo and r_tk_settle:
                ws_pnl.cell(r_op_exp, target_col_idx).value = f"={col_letter}{r_promo}+{col_letter}{r_tk_settle}"

            # 4. 利润汇总
            r_gross = None
            for k, r in row_map.items():
                if '毛利润' in k or '本期利润' in k:
                    r_gross = r
                    break
            if r_gross and r_rev and r_main_cost and r_op_exp:
                ws_pnl.cell(r_gross, target_col_idx).value = f"={col_letter}{r_rev}-{col_letter}{r_main_cost}-{col_letter}{r_op_exp}"

            # 5. 底部红字说明（清理旧说明，写新说明）
            max_limit = min(ws_pnl.max_row, 200)
            for row in ws_pnl.iter_rows(min_row=1, max_row=max_limit):
                for cell in row:
                    if cell.value and ("补充" in str(cell.value) or "未匹配：取消已发货" in str(cell.value)):
                        cell.value = None

            last_pnl_row = ws_pnl.max_row + 1
            cell_red = ws_pnl.cell(last_pnl_row, 1)
            cell_red.value = f"未匹配：取消已发货：{unmatched_shipped_count}单数"
            cell_red.font  = Font(color="FF0000")

        # --- C. 写入未匹配 ---
        if not df_unmatched.empty:
            if sheet_map.get('unmatched') and sheet_map['unmatched'] in wb.sheetnames:
                ws_u = wb[sheet_map['unmatched']]
                if ws_u.max_row >= 1:
                    ws_u.delete_rows(1, ws_u.max_row)
            else:
                ws_u = wb.create_sheet('未匹配')

            headers = [c for c in df_unmatched.columns.tolist() if c != 'clean_oid']
            ws_u.append(headers)
            df_to_write = df_unmatched.drop(columns=['clean_oid'], errors='ignore')
            for row in dataframe_to_rows(df_to_write, index=False, header=False):
                ws_u.append(row)

        wb.save(output_path)
        wb.close()
        print(f"   [成功] 完成: {new_file_name}")

    except Exception as e:
        print(f"   [失败] 写入过程出错: {e}")
        import traceback
        traceback.print_exc()


# ================================================================
# ============== 第二段 B：TK无马帮 平台处理函数 =================
# ================================================================

def process_tk_no_horse(source_path, cost_map, result_dir):
    """TK无马帮平台单文件处理：逻辑与Shopee相同，使用独立参数配置"""
    file_name = os.path.basename(source_path)
    SN = TK_NOHORSE_SHEET_NAMES
    ER = TK_NOHORSE_EXCHANGE_RATE
    UF = TK_NOHORSE_UNIT_FEE
    TY = TK_NOHORSE_TARGET_YEAR
    TM = TK_NOHORSE_TARGET_MONTH

    print(f"\n{'='*60}")
    print(f">>> [TK无马帮] 正在处理: {file_name}")

    # ---------- 1. 读取各Sheet ----------
    xls = pd.ExcelFile(source_path)
    df_order  = load_and_clean_sheet(xls, SN['order'],  ['Order ID', 'SKU'])
    df_income = load_and_clean_sheet(xls, SN['income'], ['Order ID', 'Total Released Amount'])
    df_brush  = load_and_clean_sheet(xls, SN['brush'],  ['线上单号', '佣金'])
    df_ads    = load_and_clean_sheet(xls, SN['ads'],    ['花费'])

    if df_order.empty or df_income.empty:
        print("   [跳过] 缺少订单或收入明细表")
        return

    # ---------- 2. 时间筛选 ----------
    if 'Order Creation Date' in df_order.columns:
        date_col = 'Order Creation Date'
    elif 'Order created time' in df_order.columns:
        date_col = 'Order created time'
    else:
        date_col = df_order.columns[0]

    df_order['Created Time'] = pd.to_datetime(df_order[date_col], errors='coerce')
    mask_order = (df_order['Created Time'].dt.year == TY) & (df_order['Created Time'].dt.month == TM)
    df_order_m = df_order[mask_order].copy()
    print(f"   {TM}月订单数: {len(df_order_m)}")
    if len(df_order_m) == 0:
        print(f"   [跳过] {TM}月无订单")
        return

    inc_date_col = 'Order Creation Date' if 'Order Creation Date' in df_income.columns else 'Order created time'
    df_income['Created Time'] = pd.to_datetime(df_income[inc_date_col], errors='coerce')
    mask_income = (df_income['Created Time'].dt.year == TY) & (df_income['Created Time'].dt.month == TM)
    df_income_m = df_income[mask_income].copy()

    # ---------- 3. 成本计算 ----------
    sku_col = next((c for c in df_order_m.columns if 'Reference' in c and 'Parent' not in c), None)
    if not sku_col:
        sku_col = next((c for c in df_order_m.columns if 'SKU' in c), 'Seller SKU')
    qty_col    = 'Quantity'
    track_col  = next((c for c in df_order_m.columns if 'Tracking' in c), 'Tracking Number*')
    status_col = 'Order Status'

    df_order_m['Line_Total_Cost'] = df_order_m.apply(
        lambda row: cost_map.get(str(row[sku_col]).strip(), 0) * clean_currency(row[qty_col]), axis=1
    )

    order_agg = df_order_m.groupby('Order ID').agg({
        'Created Time': 'first',
        sku_col:        lambda x: '、'.join([str(i) for i in x if str(i) != 'nan']),
        qty_col:        lambda x: sum([clean_currency(i) for i in x]),
        'Line_Total_Cost': 'sum',
        status_col:     'first',
        track_col:      'first'
    }).reset_index()
    order_agg.rename(columns={sku_col: 'Merged_SKU', 'Line_Total_Cost': 'Total_Cost_CNY'}, inplace=True)

    # ---------- 4. 跨表匹配，生成销售明细 ----------
    cols      = df_income.columns
    c_price   = next((c for c in cols if 'Original Product Price' in c), None)
    c_promo   = next((c for c in cols if 'Seller Product Promotion' in c), None)
    c_voucher = next((c for c in cols if 'Voucher Sponsored by Seller' in c), None)
    c_refund  = next((c for c in cols if 'Refund Amount' in c), None)
    income_dict = df_income.set_index('Order ID').to_dict('index')

    sale_details = []
    for _, row in order_agg.iterrows():
        oid     = row['Order ID']
        inc_row = income_dict.get(oid, {})
        sale_details.append({
            '日期':           row['Created Time'],
            '订单号':         oid,
            '商品SKU':        row['Merged_SKU'],
            '销售数量':       row[qty_col],
            '销售单价':       clean_currency(inc_row.get(c_price, 0)),
            '优惠金额':       clean_currency(inc_row.get(c_promo, 0)),
            '优惠劵金额':     clean_currency(inc_row.get(c_voucher, 0)),
            '退款金额':       clean_currency(inc_row.get(c_refund, 0)),
            '商品成本总额(CNY)': row['Total_Cost_CNY'],
            'Status':         row[status_col],
            'Tracking':       row[track_col],
        })
    df_sales = pd.DataFrame(sale_details)
    df_sales = df_sales[df_sales['销售单价'] > 0.001].copy()
    df_sales.reset_index(drop=True, inplace=True)

    # ---------- 5. 未匹配统计 ----------
    ids_order     = set(df_order_m['Order ID'])
    ids_income    = set(df_income['Order ID'])
    unmatched_ids = ids_order - ids_income
    df_unmatched  = df_order_m[df_order_m['Order ID'].isin(unmatched_ids)].copy()

    count_cancel_shipped = count_cancel_unshipped = 0
    if not df_unmatched.empty:
        def check_status_nohorse(row):
            track  = str(row.get(track_col, ''))
            status = str(row.get(status_col, ''))
            if 'Cancel' in status:
                return 'Shipped' if len(track) > 5 and track.lower() != 'nan' else 'Unshipped'
            return 'Other'
        df_unmatched['Check_Status'] = df_unmatched.apply(check_status_nohorse, axis=1)
        count_cancel_shipped   = len(df_unmatched[df_unmatched['Check_Status'] == 'Shipped'])
        count_cancel_unshipped = len(df_unmatched[df_unmatched['Check_Status'] == 'Unshipped'])

    # ---------- 6. 补单与费用计算 ----------
    brushing_ids = set()
    if not df_brush.empty:
        brush_id_col = next((c for c in df_brush.columns if '单号' in c), df_brush.columns[0])
        brushing_ids = set(df_brush[brush_id_col].astype(str).str.strip())

    c_rebate = next((c for c in cols if 'Rebate' in c and 'Shopee' in c), None)
    cols_brush_amt = [c for c in [c_price, c_promo, c_voucher, c_refund, c_rebate] if c]

    all_cols = list(df_income.columns)
    try:
        idx_start = next(i for i, c in enumerate(all_cols) if 'Seller shipping fee' in c or 'Shipping Fee Paid by Seller' in c)
        idx_end   = next(i for i, c in enumerate(all_cols) if 'Withholding Tax' in c)
        cols_platform_fee_range = all_cols[idx_start: idx_end + 1]
    except:
        cols_platform_fee_range = [c for c in all_cols if 'Fee' in c or 'Tax' in c or 'Commission' in c]

    brush_refund_sum_php = brush_platform_fee_php = 0
    for bid in brushing_ids:
        if bid in income_dict:
            row = income_dict[bid]
            for c in cols_brush_amt:          brush_refund_sum_php  += clean_currency(row.get(c, 0))
            for c in cols_platform_fee_range: brush_platform_fee_php += clean_currency(row.get(c, 0))

    cols_pure_platform     = [c for c in cols_platform_fee_range if 'Tax' not in c]
    total_platform_fee_php = sum(df_income_m[c].apply(clean_currency).sum() for c in cols_pure_platform if c in df_income_m.columns)
    c_tax         = next((c for c in all_cols if 'Withholding Tax' in c), None)
    total_tax_php = df_income_m[c_tax].apply(clean_currency).sum() if c_tax else 0

    ads_col = next((c for c in df_ads.columns if '花费' in c or 'Expense' in c), None) if not df_ads.empty else None
    ads_cny = df_ads[ads_col].apply(clean_currency).sum() * ER if ads_col else 0

    brush_comm = brush_diff = 0
    if not df_brush.empty:
        c_comm = next((c for c in df_brush.columns if '佣金' in c), None)
        c_diff = next((c for c in df_brush.columns if '汇差' in c), None)
        if c_comm: brush_comm = df_brush[c_comm].apply(clean_currency).sum()
        if c_diff: brush_diff = df_brush[c_diff].apply(clean_currency).sum()

    metrics = {
        'total_platform_fee_php': total_platform_fee_php,
        'total_tax_php':          total_tax_php,
        'brush_amt_php':          brush_refund_sum_php,
        'brush_platform_fee_php': brush_platform_fee_php,
        'brush_commission':       brush_comm,
        'brush_diff':             brush_diff,
        'brush_empty_pkg':        len(brushing_ids) * UF,
        'ads_cny':                ads_cny,
        'cancel_shipped':         count_cancel_shipped,
        'cancel_unshipped':       count_cancel_unshipped,
        'order_count':            len(df_sales),
    }

    # ---------- 7. 复制源文件到 Result 目录 ----------
    current_time  = time.strftime("%H点%M分%S秒")
    new_file_name = f"新_{file_name.replace('.xlsx', '')}_{current_time}.xlsx"
    file_output   = os.path.join(result_dir, new_file_name)
    print(f"   生成副本: {new_file_name}")
    try:
        shutil.copy(source_path, file_output)
        os.chmod(file_output, stat.S_IWRITE)
        time.sleep(1)
    except Exception as e:
        print(f"   [错误] 复制失败: {e}")
        return

    # ---------- 8. 写入 Excel ----------
    try:
        wb = openpyxl.load_workbook(file_output)

        # --- A. 销售明细 ---
        sales_sheet_name = None
        for n in SN['sales']:
            if n in wb.sheetnames:
                sales_sheet_name = n
                break
        if not sales_sheet_name:
            ws_sales = wb.create_sheet('销售明细')
            ws_sales.append(['日期', '订单号', '商品SKU', '销售数量', '销售单价', '优惠金额',
                             '优惠劵金额', '退款金额', '销售金额(PHP)', '销售金额(CNY)', '商品成本总额(CNY)', '毛利(CNY)'])
            sales_sheet_name = '销售明细'
        else:
            ws_sales = wb[sales_sheet_name]
            if ws_sales.max_row >= 2:
                ws_sales.delete_rows(2, ws_sales.max_row)

        for i, row in df_sales.iterrows():
            r = i + 2
            ws_sales.cell(r, 1).value  = str(row['日期'])[:10]
            ws_sales.cell(r, 2).value  = row['订单号']
            ws_sales.cell(r, 3).value  = row['商品SKU']
            ws_sales.cell(r, 4).value  = row['销售数量']
            ws_sales.cell(r, 5).value  = row['销售单价']
            ws_sales.cell(r, 6).value  = row['优惠金额']
            ws_sales.cell(r, 7).value  = row['优惠劵金额']
            ws_sales.cell(r, 8).value  = row['退款金额']
            ws_sales.cell(r, 9).value  = f'=E{r}+F{r}+G{r}+H{r}'
            ws_sales.cell(r, 10).value = f'=I{r}*{ER}'
            ws_sales.cell(r, 11).value = row['商品成本总额(CNY)']
            ws_sales.cell(r, 12).value = f'=J{r}-K{r}'
            for c in range(1, 13):
                safe_border(ws_sales, r, c)

        last_sales_row = len(df_sales) + 1

        # --- B. 补单明细 ---
        brush_sheet_name = None
        for n in SN['brush']:
            if n in wb.sheetnames:
                brush_sheet_name = n
                break

        v_shopee_fee    = abs(metrics['total_platform_fee_php']) * ER
        v_tax           = abs(metrics['total_tax_php']) * ER
        v_brush_amt     = metrics['brush_amt_php'] * ER
        v_brush_plat    = abs(metrics['brush_platform_fee_php']) * ER
        brush_comm      = metrics['brush_commission']
        brush_empty_pkg = metrics['brush_empty_pkg']
        brush_diff_expr = 0

        if brush_sheet_name and not df_brush.empty:
            ws_brush = wb[brush_sheet_name]
            header_row = col_idx_xiaoji = col_idx_chengben = col_idx_huicha = None

            for row in ws_brush.iter_rows(min_row=1, max_row=20):
                for cell in row:
                    if cell.value:
                        val_str = str(cell.value).strip()
                        if '垫付小计' in val_str:   col_idx_xiaoji   = cell.column
                        elif '垫付成本' in val_str:  col_idx_chengben = cell.column
                        elif '汇率差' in val_str or '汇差' in val_str: col_idx_huicha = cell.column
                if col_idx_xiaoji and col_idx_chengben:
                    header_row = row[0].row
                    break

            if header_row and col_idx_xiaoji and col_idx_chengben:
                if not col_idx_huicha:
                    col_idx_huicha = ws_brush.max_column + 1
                    ws_brush.cell(row=header_row, column=col_idx_huicha, value="汇率差")

                col_let_xj = get_column_letter(col_idx_xiaoji)
                col_let_cb = get_column_letter(col_idx_chengben)
                col_let_hc = get_column_letter(col_idx_huicha)
                max_brush_row = ws_brush.max_row

                for r in range(header_row + 1, max_brush_row + 1):
                    c_xj = ws_brush.cell(row=r, column=col_idx_xiaoji)
                    c_cb = ws_brush.cell(row=r, column=col_idx_chengben)
                    for c in [c_xj, c_cb]:
                        if c.value:
                            cleaned = str(c.value).replace('₱', '').replace('PHP', '').replace(',', '').strip()
                            try:
                                c.value = float(cleaned)
                            except:
                                pass
                    if c_xj.value is not None and c_cb.value is not None:
                        ws_brush.cell(row=r, column=col_idx_huicha).value = f'={col_let_xj}{r}-{ER}*{col_let_cb}{r}'

                brush_diff_expr = f"=SUM('{brush_sheet_name}'!{col_let_hc}{header_row+1}:{col_let_hc}{max_brush_row})"
        else:
            v_brush_amt = brush_comm = v_brush_plat = brush_empty_pkg = 0
            brush_diff_expr = 0

        # --- C. 利润表（统一动态月份列写入） ---
        pnl_sheet_name = None
        for n in SN['pnl']:
            if n in wb.sheetnames:
                pnl_sheet_name = n
                break

        if pnl_sheet_name:
            ws_pnl = wb[pnl_sheet_name]
            target_col_idx = find_month_col(ws_pnl, TM)
            col_letter     = get_column_letter(target_col_idx)

            val_map = {
                '本期销售':     f'=SUM({sales_sheet_name}!J2:J{last_sales_row})',
                '其中:补单金额': v_brush_amt,
                '销售退款':     f'=ABS(SUM({sales_sheet_name}!H2:H{last_sales_row}))*{ER}',
                '实际产品成本': f'=SUM({sales_sheet_name}!K2:K{last_sales_row})',
                '发单费':       f'=({metrics["order_count"]}+{metrics["cancel_shipped"]})*{UF}',
                '广告费':       metrics['ads_cny'],
                '补单佣金':     brush_comm,
                '补单平台费用': v_brush_plat,
                '补单汇率差额': brush_diff_expr,
                '补单空包费':   brush_empty_pkg,
                '平台费用':     v_shopee_fee,
                '预扣税':       v_tax,
                '团队运营费用': 0,
            }

            search_keys = [
                '实际主营收入', '其中:补单金额', '实际主营成本', '实际产品成本',
                '团队运营费用', '补单平台费用', '补单汇率差额', '补单空包费',
                'shopee结算', '本期销售', '销售退款', '营业费用', '推广费用',
                '补单佣金', '平台费用', '毛利润', '本期利润', '发单费', '广告费', '预扣税'
            ]

            row_idx_map = {}
            for row in ws_pnl.iter_rows(min_row=1, max_row=100):
                cell_a   = row[0]
                cell_val = str(cell_a.value).strip() if cell_a.value else ""
                for key in search_keys:
                    if key in cell_val:
                        row_idx_map[key] = cell_a.row
                        if key in val_map:
                            ws_pnl.cell(row=cell_a.row, column=target_col_idx).value = val_map[key]
                        break
                if "未匹配" in cell_val or "补充" in cell_val:
                    cell_a.value = f"未匹配说明：取消已发货 {metrics['cancel_shipped']} 单，取消未发货 {metrics['cancel_unshipped']} 单"

            def get_cell_nh(key):
                return f"{col_letter}{row_idx_map.get(key, 0)}"

            if '实际主营收入' in row_idx_map:
                ws_pnl.cell(row_idx_map['实际主营收入'], target_col_idx).value = \
                    f"={get_cell_nh('本期销售')}-{get_cell_nh('其中:补单金额')}-{get_cell_nh('销售退款')}"
            if '实际主营成本' in row_idx_map:
                ws_pnl.cell(row_idx_map['实际主营成本'], target_col_idx).value = \
                    f"={get_cell_nh('实际产品成本')}+{get_cell_nh('发单费')}"
            if '推广费用' in row_idx_map:
                sub = [get_cell_nh(k) for k in ['广告费', '补单佣金', '补单平台费用', '补单汇率差额', '补单空包费'] if k in row_idx_map]
                ws_pnl.cell(row_idx_map['推广费用'], target_col_idx).value = f"={'+'.join(sub)}"
            if 'shopee结算' in row_idx_map:
                ws_pnl.cell(row_idx_map['shopee结算'], target_col_idx).value = \
                    f"={get_cell_nh('平台费用')}+{get_cell_nh('预扣税')}"
            if '营业费用' in row_idx_map:
                ws_pnl.cell(row_idx_map['营业费用'], target_col_idx).value = \
                    f"={get_cell_nh('推广费用')}+{get_cell_nh('shopee结算')}+{get_cell_nh('团队运营费用')}"
            profit_key = '本期利润' if '本期利润' in row_idx_map else '毛利润'
            if profit_key in row_idx_map:
                ws_pnl.cell(row_idx_map[profit_key], target_col_idx).value = \
                    f"={get_cell_nh('实际主营收入')}-{get_cell_nh('实际主营成本')}-{get_cell_nh('营业费用')}"

        # --- D. 未匹配订单 ---
        un_sheet_name = None
        for n in SN['unmatched']:
            if n in wb.sheetnames:
                un_sheet_name = n
                break
        if not df_unmatched.empty:
            if un_sheet_name:
                ws_un = wb[un_sheet_name]
                if ws_un.max_row >= 1:
                    ws_un.delete_rows(1, ws_un.max_row)
            else:
                ws_un = wb.create_sheet('未匹配订单')
            ws_un.append(df_unmatched.columns.tolist())
            for r in dataframe_to_rows(df_unmatched, index=False, header=False):
                ws_un.append(r)

        wb.save(file_output)
        wb.close()
        print("   [成功] 处理完成")

    except Exception as e:
        print(f"   [失败] 处理出错: {e}")
