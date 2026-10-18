import pandas as pd
import numpy as np
import os
import shutil
import warnings
import openpyxl
import time
import glob
import stat
import re
from copy import copy
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 忽略警告
warnings.filterwarnings('ignore')

# ================= 配置区域 =================

# 1. 待处理文件所在的文件夹名称（自动定位脚本所在目录下的 Shopee_Data，Mac/Windows 通用）
TARGET_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Shopee_Data_Shunchang')

# 1.1 输出文件夹（脚本同目录下的 output_Shunchang）
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output_Shunchang')

# 2. 成本表文件名
COST_FILE_NAME = '顺昌商品成本.xlsx'

# 3. 参数配置
PH_EXCHANGE_RATE = 0.1065        # 1 PHP = 0.1065 CNY
ID_EXCHANGE_RATE = 0.0003754     # 1 IDR = 0.0003754 CNY
MY_EXCHANGE_RATE = 1.629         # 1 MYR = 1.629 CNY
ID_ADS_MULTIPLIER = 1.11         # 印尼广告费税费系数
MY_SKU_UNIT_COST = 8.0       # 马来每个基础 SKU 的人民币成本
UNIT_FEE = 3.0          # 发单费/空包费单价

# 4. 时间筛选配置 (可随时修改，控制只处理该年月的数据)
# TARGET_MONTH = 0 时自动判断月份（取订单数据中订单量最多的年月）；非 0 时按指定年月筛选
TARGET_YEAR = 2026
TARGET_MONTH = 0

# Sheet 名称映射
SHEET_NAMES = {
    'order': ['订单明细原表', '订单明细','订单明细表'],
    'income': ['收入明细表', '收入明细', '账单明细'],
    'brush': ['补单明细', '补单','补单明细 '],
    'ads': ['广告费', '广告'],
    'sales': ['销售明细', '销售明细表'],
    'pnl': ['利润表', 'Profit'],
    'unmatched': ['未匹配', '未匹配订单']
}

# ================= 1. 基础工具函数 =================

def find_header_row(df, keywords):
    for idx, row in df.head(20).iterrows():
        row_str = " ".join([str(x) for x in row.values])
        if all(k in row_str for k in keywords):
            return idx
    return 0

def load_and_clean_sheet(excel_file, sheet_candidates, keywords):
    sheet_name = None
    for name in sheet_candidates:
        if name in excel_file.sheet_names:
            sheet_name = name
            break
    if not sheet_name: return pd.DataFrame()
    try:
        df_preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=20)
        header_idx = find_header_row(df_preview, keywords)
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_idx, dtype=str)
        return df
    except: return pd.DataFrame()

def detect_market(excel_file, file_name):
    """按文件名和收入表表头识别菲律宾/马来西亚。"""
    if '印尼' in file_name:
        return 'ID'
    if '马来' in file_name:
        return 'MY'
    for name in SHEET_NAMES['income']:
        if name not in excel_file.sheet_names:
            continue
        try:
            preview = pd.read_excel(excel_file, sheet_name=name, header=None, nrows=8)
            text = ' '.join(str(v) for v in preview.values.ravel() if pd.notna(v))
            if '订单编号' in text or '拨款金额 (RM)' in text:
                return 'MY'
        except Exception:
            pass
    return 'PH'

def normalize_malaysia_income(df):
    """把马来账单的中英文字段统一为脚本内部字段名，列顺序保持不变。"""
    aliases = {
        '订单编号': 'Order ID',
        '订单成立时间': 'Order Creation Date',
        '退款金額': 'Refund Amount',
        'Shopee回扣金额': 'Rebate Provided by Shopee',
        '卖家赞助的优惠券': 'Voucher Sponsored by Seller',
    }
    return df.rename(columns={k: v for k, v in aliases.items() if k in df.columns})

def normalize_indonesia_order(df):
    """把印尼订单表字段统一为脚本内部字段名。"""
    aliases = {
        'No. Pesanan': 'Order ID',
        'Status Pesanan': 'Order Status',
        'No. Resi': 'Tracking Number*',
        'Waktu Pengiriman Diatur': 'Ship Time',
        'Waktu Pesanan Dibuat': 'Order Creation Date',
        'Nomor Referensi SKU': 'SKU Reference No.',
        'Jumlah': 'Quantity',
    }
    return df.rename(columns={k: v for k, v in aliases.items() if k in df.columns})

def normalize_indonesia_income(df):
    """印尼账单可能导出为印尼语表头，统一为英文字段名（列顺序不变）。"""
    aliases = {
        'No. Pesanan': 'Order ID',
        'Waktu Pesanan Dibuat': 'Order Creation Date',
        'Total Penghasilan': 'Total Released Amount (Rp)',
        'Harga Produk': 'Product Price',
        'Jumlah Pengembalian Dana ke Pembeli': 'Refund Amount',
        'Ongkir Dibayar Pembeli': 'Shipping Fee Paid by Buyer',
        'Voucher disponsor oleh Penjual': 'Voucher Sponsored by Seller',
        'Diskon Produk dari Shopee': 'Rebate Provided by Shopee',
        'Biaya Isi Saldo Otomatis (dari Penghasilan)': 'Ads Escrow Top Up Fee',
        'PPh 22': 'WHT 22',
    }
    return df.rename(columns={k: v for k, v in aliases.items() if k in df.columns})

def resolve_cost(sku, cost_map):
    """按成本表解析单件/倍数/组合 SKU，返回（成本，缺失基础SKU）。"""
    text = str(sku or '').strip()
    if not text or text.lower() == 'nan':
        return 0.0, []
    text = text.replace('×', '*').replace('＊', '*').replace('＋', '+')
    text = re.sub(r'\s*\*\s*(\d+)', r'*\1', text)
    if text in cost_map:
        return cost_map[text], []
    # 成本表有时只给出 SKU*n 的总价；输入为基础 SKU 时按倍数反推单件价。
    inferred_unit_costs = []
    variant_pattern = re.compile(rf'{re.escape(text)}\s*\*\s*(\d+)$')
    for configured_sku, configured_cost in cost_map.items():
        match = variant_pattern.fullmatch(str(configured_sku).strip())
        if match and int(match.group(1)) > 0:
            inferred_unit_costs.append(configured_cost / int(match.group(1)))
    if inferred_unit_costs:
        return inferred_unit_costs[0], []
    total, missing = 0.0, []
    for part in re.split(r'[、，,+\r\n]+', text):
        part = part.strip()
        if not part:
            continue
        if part in cost_map:
            total += cost_map[part]
            continue
        multiple = re.fullmatch(r'(.+?)\*\s*(\d+)', part)
        if multiple:
            base, qty = multiple.group(1).strip(), int(multiple.group(2))
            if base in cost_map:
                total += cost_map[base] * qty
            else:
                missing.append(base)
            continue
        if '*' in part:
            components = [v.strip() for v in part.split('*') if v.strip()]
            unknown = [v for v in components if v not in cost_map]
            if unknown:
                missing.extend(unknown)
            else:
                total += sum(cost_map[v] for v in components)
            continue
        missing.append(part)
    return total, missing

def clean_currency(x):
    if pd.isna(x): return 0.0
    s = str(x).replace('PHP', '').replace('RM', '').replace(',', '').replace(' ', '').replace('₱', '').strip()
    try: return float(s)
    except: return 0.0

def malaysia_sku_count(sku):
    """计算马来 SKU 字段包含的基础 SKU 数量，支持顿号/加号组合和 SKU*n。"""
    if sku is None or pd.isna(sku):
        return 0
    text = str(sku).strip()
    if not text or text.lower() == 'nan':
        return 0
    text = re.sub(r'[\r\n]+\s*([*xX×])\s*(\d+)', r'\1\2', text)
    total = 0
    for part in re.split(r'[、，,+\r\n]+', text):
        part = part.strip()
        if not part:
            continue
        match = re.search(r'(?:\*|[xX×])\s*(\d+)\s*$', part)
        total += int(match.group(1)) if match else 1
    return total

ADS_AMOUNT_HEADERS = ('花费', '总额', 'Amount', 'Jumlah')

def load_ads_amount(excel_file):
    """读取广告表金额列合计，返回 (合计, 是否旧版“花费”报表)。
    新版为广告交易历史记录（总额/Amount/Jumlah 为负数扣费，含正数返还），取净额。"""
    sheet_name = next((n for n in SHEET_NAMES['ads'] if n in excel_file.sheet_names), None)
    if not sheet_name:
        return 0.0, False
    raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, dtype=str)
    for idx, row in raw.head(20).iterrows():
        cells = [str(v).strip() for v in row.values]
        for key in ADS_AMOUNT_HEADERS:
            if key in cells:
                col = cells.index(key)
                total = raw.iloc[idx + 1:, col].apply(clean_currency).sum()
                return total, key == '花费'
    print("   [警告] 广告表找不到金额列（花费/总额/Amount/Jumlah），广告费按 0 计算")
    return 0.0, False

def safe_border(ws, row_idx, col_idx):
    """安全地给单元格加边框"""
    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    cell = ws.cell(row=row_idx, column=col_idx)
    cell.border = border
    cell.alignment = Alignment(horizontal='center', vertical='center')

# ================= 2. 核心处理逻辑 =================

def process_single_file(source_path, cost_map):
    file_name = os.path.basename(source_path)
    print(f"\n{'='*60}")
    print(f">>> 正在处理: {file_name}")

    # 1. 读取数据
    xls = pd.ExcelFile(source_path)
    market = detect_market(xls, file_name)
    rates = {'PH': PH_EXCHANGE_RATE, 'ID': ID_EXCHANGE_RATE, 'MY': MY_EXCHANGE_RATE}
    currencies = {'PH': 'PHP', 'ID': 'IDR', 'MY': 'MYR'}
    market_names = {'PH': '菲律宾', 'ID': '印尼', 'MY': '马来西亚'}
    exchange_rate = rates[market]
    currency = currencies[market]
    print(f"   市场: {market_names[market]}，汇率: {exchange_rate}")

    # 马来02店的订单表 A1 为空，因此使用其它稳定表头定位首行，再补回 Order ID。
    order_keywords = (['No. Pesanan', 'Nomor Referensi SKU'] if market == 'ID'
                      else ['SKU Reference No.', 'Order Creation Date'])
    df_order = load_and_clean_sheet(xls, SHEET_NAMES['order'], order_keywords)
    income_keywords = ['订单编号', 'Product Price'] if market == 'MY' else ['Order ID', 'Total Released Amount']
    df_income = load_and_clean_sheet(xls, SHEET_NAMES['income'], income_keywords)
    if market == 'MY':
        df_income = normalize_malaysia_income(df_income)
    elif market == 'ID':
        df_order = normalize_indonesia_order(df_order)
        df_income = normalize_indonesia_income(df_income)
    if 'Order ID' not in df_order.columns and len(df_order.columns):
        first_col = df_order.columns[0]
        if pd.isna(first_col) or str(first_col).startswith('Unnamed'):
            df_order.rename(columns={first_col: 'Order ID'}, inplace=True)
    df_brush = load_and_clean_sheet(xls, SHEET_NAMES['brush'], ['线上单号', '佣金'])
    ads_sum, ads_is_spend_report = load_ads_amount(xls)

    if df_order.empty or df_income.empty:
        print("   [跳过] 缺少订单或收入明细表")
        return False

    # 2. 数据计算
    # --- 时间筛选：按配置的年月精准筛选订单 ---
    if 'Order Creation Date' in df_order.columns: date_col = 'Order Creation Date'
    elif 'Order created time' in df_order.columns: date_col = 'Order created time'
    else: date_col = df_order.columns[0]

    df_order['Created Time'] = pd.to_datetime(df_order[date_col], errors='coerce')

    # --- 确定目标年月：TARGET_MONTH=0 时自动判断（取订单量最多的年月），否则用指定年月 ---
    if TARGET_MONTH == 0:
        valid_dates = df_order['Created Time'].dropna()
        if valid_dates.empty:
            print("   [跳过] 无有效订单日期，无法自动判断月份")
            return False
        ym_mode = valid_dates.dt.to_period('M').mode()[0]
        cur_year, cur_month = ym_mode.year, ym_mode.month
        print(f"   自动判断目标年月: {cur_year}年 {cur_month}月")
    else:
        cur_year, cur_month = TARGET_YEAR, TARGET_MONTH

    mask_order = (df_order['Created Time'].dt.year == cur_year) & (df_order['Created Time'].dt.month == cur_month)
    df_order_dec = df_order[mask_order].copy()
    print(f"   {cur_month}月订单数: {len(df_order_dec)}")

    if len(df_order_dec) == 0:
        print(f"   [跳过] {cur_month}月无订单")
        return False

    # --- 收入表筛选：按目标年月精准筛选 ---
    inc_date_col = 'Order Creation Date' if 'Order Creation Date' in df_income.columns else 'Order created time'
    df_income['Created Time'] = pd.to_datetime(df_income[inc_date_col], errors='coerce')
    mask_income = (df_income['Created Time'].dt.year == cur_year) & (df_income['Created Time'].dt.month == cur_month)
    df_income_dec = df_income[mask_income].copy()

    # 成本计算
    sku_col = next((c for c in df_order_dec.columns if 'Reference' in c and 'Parent' not in c), None)
    if not sku_col: sku_col = next((c for c in df_order_dec.columns if 'SKU' in c), 'Seller SKU')
    qty_col = 'Quantity'

    def get_line_cost(sku, qty):
        qty_num = clean_currency(qty)
        if market == 'MY':
            return MY_SKU_UNIT_COST * malaysia_sku_count(sku) * qty_num
        unit_cost, _ = resolve_cost(sku, cost_map)
        return unit_cost * qty_num

    def get_row_cost(row):
        return get_line_cost(row[sku_col], row[qty_col])

    df_order_dec['Line_Total_Cost'] = df_order_dec.apply(get_row_cost, axis=1)
    if market != 'MY':
        missing_skus = set()
        for sku in df_order_dec[sku_col].dropna():
            _, missing = resolve_cost(sku, cost_map)
            missing_skus.update(missing)
        if missing_skus:
            print(f"   [警告] 成本表缺少 SKU: {', '.join(sorted(missing_skus))}（对应成本按 0 计算）")
    track_col = next((c for c in df_order_dec.columns if 'Tracking' in c), 'Tracking Number*')
    status_col = 'Order Status'

    order_agg = df_order_dec.groupby('Order ID').agg({
        'Created Time': 'first',
        sku_col: lambda x: '、'.join([str(i) for i in x if str(i)!='nan']),
        qty_col: lambda x: sum([clean_currency(i) for i in x]),
        'Line_Total_Cost': 'sum',
        status_col: 'first',
        track_col: 'first'
    }).reset_index()
    order_agg.rename(columns={sku_col: 'Merged_SKU', 'Line_Total_Cost': 'Total_Cost_CNY'}, inplace=True)

    # 跨表匹配
    cols = df_income.columns
    if market in ('MY', 'ID'):
        c_price = next((c for c in cols if str(c).strip() == 'Product Price'), None)
        c_promo = next((c for c in cols if str(c).strip() == 'Voucher Sponsored by Seller'), None)
        c_voucher = next((c for c in cols if str(c).strip() == 'Rebate Provided by Shopee'), None)
    else:
        c_price = next((c for c in cols if 'Original Product Price' in c), None)
        c_promo = next((c for c in cols if 'Seller Product Promotion' in c), None)
        c_voucher = next((c for c in cols if 'Voucher Sponsored by Seller' in c), None)
    c_refund = next((c for c in cols if 'Refund Amount' in c), None)
    income_dict = df_income.set_index('Order ID').to_dict('index')

    # 补单ID集合
    brushing_ids = set()
    if not df_brush.empty:
        # 优先匹配"线上单号"（Shopee订单号），其次才是其他含"单号"的列
        brush_id_col = next((c for c in df_brush.columns if '线上单号' in c),
                            next((c for c in df_brush.columns if '单号' in c), df_brush.columns[0]))
        brushing_ids = set(df_brush[brush_id_col].astype(str).str.strip())

    sale_details = []
    for _, row in order_agg.iterrows():
        oid = row['Order ID']
        inc_row = income_dict.get(oid, {})
        sale_details.append({
            '日期': row['Created Time'],
            '订单号': oid,
            '商品SKU': row['Merged_SKU'],
            '销售数量': row[qty_col],
            '销售单价': clean_currency(inc_row.get(c_price, 0)),
            '优惠金额': clean_currency(inc_row.get(c_promo, 0)),
            '优惠劵金额': clean_currency(inc_row.get(c_voucher, 0)),
            '退款金额': clean_currency(inc_row.get(c_refund, 0)),
            '商品成本总额(CNY)': row['Total_Cost_CNY'],
            'Status': row[status_col],
            'Tracking': row[track_col]
        })
    df_sales_detail = pd.DataFrame(sale_details)

    df_sales_detail = df_sales_detail[df_sales_detail['销售单价'] > 0.001].copy()
    df_sales_detail.reset_index(drop=True, inplace=True)

    # 未匹配计算
    ids_order = set(df_order_dec['Order ID'])
    ids_income = set(df_income['Order ID'])
    unmatched_ids = ids_order - ids_income
    df_unmatched = df_order_dec[df_order_dec['Order ID'].isin(unmatched_ids)].copy()

    # 已发货/未发货以"发货时间"(Ship Time)为准：未匹配表中该列有发货时间即视为已发货
    ship_time_col = next((c for c in df_order_dec.columns if 'Ship Time' in c or '发货时间' in c), None)
    count_cancel_shipped = 0
    count_cancel_unshipped = 0
    if not df_unmatched.empty:
        def check_status(row):
            ship_time = str(row.get(ship_time_col, '')).strip() if ship_time_col else ''
            has_ship = ship_time != '' and ship_time.lower() != 'nan'
            return 'Shipped' if has_ship else 'Unshipped'
        df_unmatched['Check_Status'] = df_unmatched.apply(check_status, axis=1)
        count_cancel_shipped = len(df_unmatched[df_unmatched['Check_Status'] == 'Shipped'])
        count_cancel_unshipped = len(df_unmatched[df_unmatched['Check_Status'] == 'Unshipped'])

    # 补单与费用
    c_rebate = next((c for c in cols if 'Rebate' in c and 'Shopee' in c), None)
    cols_brush_amt = list(dict.fromkeys(c for c in [c_price, c_promo, c_voucher, c_refund, c_rebate] if c))

    all_cols = list(df_income.columns)
    if market == 'MY':
        # 马来账单：M:AD（第13至30列）合计后，扣除已在销售明细提取的
        # Voucher Sponsored by Seller 与 Rebate Provided by Shopee。
        brush_platform_cols = all_cols[12:30]
    elif market == 'ID':
        idx_start = all_cols.index('Shipping Fee Paid by Buyer')
        idx_end = all_cols.index('WHT 22')
        brush_platform_cols = all_cols[idx_start:idx_end + 1]
    else:
        try:
            idx_start = next(i for i, c in enumerate(all_cols) if 'Seller shipping fee' in c or 'Shipping Fee Paid by Seller' in c)
            idx_end = next(i for i, c in enumerate(all_cols) if 'Withholding Tax' in c)
            brush_platform_cols = all_cols[idx_start : idx_end+1]
        except Exception:
            brush_platform_cols = [c for c in all_cols if 'Fee' in c or 'Tax' in c or 'Commission' in c]

    def platform_fee_amount(row, fee_cols):
        amount = sum(clean_currency(row.get(c, 0)) for c in fee_cols)
        if market in ('MY', 'ID'):
            amount -= clean_currency(row.get(c_promo, 0))
            amount -= clean_currency(row.get(c_voucher, 0))
        if market == 'ID':
            amount -= clean_currency(row.get('Ads Escrow Top Up Fee', 0))
        return amount

    brush_refund_sum_php = 0
    brush_platform_fee_php = 0
    for bid in brushing_ids:
        if bid in income_dict:
            row = income_dict[bid]
            for c in cols_brush_amt: brush_refund_sum_php += clean_currency(row.get(c, 0))
            brush_platform_fee_php += platform_fee_amount(row, brush_platform_cols)

    if market == 'MY':
        sales_platform_cols = all_cols[12:30]
    elif market == 'ID':
        sales_platform_cols = brush_platform_cols
    else:
        # 菲律宾保持原口径：T:AD + AG；AE 预扣税单独核算。
        sales_platform_cols = brush_platform_cols
        if len(df_income.columns) > 32:
            sales_platform_cols = list(df_income.columns[19:30]) + [df_income.columns[32]]
    matched_ids = set(df_sales_detail['订单号'])
    total_platform_fee_php = 0
    for oid in matched_ids:
        inc_row = income_dict.get(oid, {})
        total_platform_fee_php += platform_fee_amount(inc_row, sales_platform_cols)
    c_tax = next((c for c in all_cols if 'Withholding Tax' in c), None)
    total_tax_php = df_income_dec[c_tax].apply(clean_currency).sum() if c_tax else 0

    ads_cny = abs(ads_sum) * exchange_rate
    if market == 'ID':
        ads_cny *= ID_ADS_MULTIPLIER
    elif ads_is_spend_report:
        ads_cny = ads_sum * exchange_rate

    # 补单汇率差：直接用Python计算（垫付小计 - 汇率 × 垫付成本）
    brush_comm = 0
    brush_diff_value = 0
    if not df_brush.empty:
        c_comm = next((c for c in df_brush.columns if '佣金' in c), None)
        c_xiaoji = next((c for c in df_brush.columns if '垫付小计' in c), None)
        c_chengben = next((c for c in df_brush.columns if '垫付成本' in c), None)
        if c_comm: brush_comm = df_brush[c_comm].apply(clean_currency).sum()
        if c_xiaoji and c_chengben:
            xiaoji_sum = df_brush[c_xiaoji].apply(clean_currency).sum()
            chengben_sum = df_brush[c_chengben].apply(clean_currency).sum()
            brush_diff_value = xiaoji_sum - exchange_rate * chengben_sum

    metrics = {
        'total_platform_fee_php': total_platform_fee_php,
        'total_tax_php': total_tax_php,
        'brush_amt_php': brush_refund_sum_php,
        'brush_platform_fee_php': brush_platform_fee_php,
        'brush_commission': brush_comm,
        'brush_diff_value': brush_diff_value,
        'brush_empty_pkg': len(brushing_ids) * UNIT_FEE,
        'ads_cny': ads_cny,
        'cancel_shipped': count_cancel_shipped,
        'cancel_unshipped': count_cancel_unshipped,
        'order_count': len(df_sales_detail),
        'brush_count': int(df_sales_detail['订单号'].isin(brushing_ids).sum()) if not df_sales_detail.empty else 0
    }

    # ================= 3. 写入文件 =================

    # 1. 复制源文件到输出文件夹
    new_file_name = f"新_{file_name.replace('.xlsx', '')}.xlsx"
    file_output = os.path.join(OUTPUT_FOLDER, new_file_name)

    print(f"   生成副本: {new_file_name}")
    try:
        shutil.copy(source_path, file_output)
        # Windows 用 S_IWRITE 去掉只读标志；Mac/Linux 需同时保留读权限
        os.chmod(file_output, stat.S_IRUSR | stat.S_IWUSR | stat.S_IRGRP | stat.S_IROTH)
        time.sleep(1)
    except Exception as e:
        print(f"   [错误] 复制失败: {e}")
        return False

    # 2. 打开副本进行编辑
    try:
        wb = openpyxl.load_workbook(file_output)

        green_fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # --- A. 处理"销售明细" ---
        sales_sheet_name = None
        for n in SHEET_NAMES['sales']:
            if n in wb.sheetnames: sales_sheet_name = n; break

        sales_headers = [
            '日期', '订单号', '商品SKU', '销售数量',
            f'销售单价\n({currency})', f'优惠金额\n({currency})',
            f'优惠劵金额\n({currency})', f'退款金额\n({currency})',
            f'销售金额\n({currency})', '销售金额\n（¥）', '商品成本总额\n(¥)', '毛利\n(¥)'
        ]
        if not sales_sheet_name:
            ws_sales = wb.create_sheet('销售明细')
            ws_sales.append(sales_headers)
        else:
            ws_sales = wb[sales_sheet_name]
            max_row = ws_sales.max_row
            if max_row >= 2:
                ws_sales.delete_rows(2, max_row)
            if market == 'MY':
                # 马来旧模板只有11列；按新规则补齐“优惠金额”并保持原表头样式。
                style_source = ws_sales.cell(1, min(ws_sales.max_column, 11))
                for c, header in enumerate(sales_headers, 1):
                    cell = ws_sales.cell(1, c)
                    cell.value = header
                    if c == 12 and style_source.has_style:
                        cell._style = copy(style_source._style)
                        cell.alignment = copy(style_source.alignment)

        # 写入销售数据，补单行用绿色标记且成本置0
        for i, row in df_sales_detail.iterrows():
            r = i + 2
            is_brush = row['订单号'] in brushing_ids
            ws_sales.cell(r, 1).value = str(row['日期'])[:10]
            ws_sales.cell(r, 2).value = row['订单号']
            ws_sales.cell(r, 3).value = row['商品SKU']
            ws_sales.cell(r, 4).value = row['销售数量']
            ws_sales.cell(r, 5).value = row['销售单价']
            ws_sales.cell(r, 6).value = row['优惠金额']
            ws_sales.cell(r, 7).value = row['优惠劵金额']
            ws_sales.cell(r, 8).value = row['退款金额']
            ws_sales.cell(r, 9).value = f'=E{r}+F{r}+G{r}+H{r}'
            ws_sales.cell(r, 10).value = f'=I{r}*{exchange_rate}'
            ws_sales.cell(r, 11).value = 0 if is_brush else row['商品成本总额(CNY)']
            ws_sales.cell(r, 12).value = f'=J{r}-K{r}'

            if is_brush:
                for c in range(1, 13):
                    ws_sales.cell(r, c).fill = green_fill

            for c in range(1, 13): safe_border(ws_sales, r, c)

        last_sales_row = len(df_sales_detail) + 1

        # --- B. 处理费用变量（不再写入补单明细表）---
        brush_sheet_name = None
        for n in SHEET_NAMES['brush']:
            if n in wb.sheetnames: brush_sheet_name = n; break

        v_tax = abs(metrics['total_tax_php']) * exchange_rate

        if brush_sheet_name and not df_brush.empty:
            v_brush_amt = metrics['brush_amt_php'] * exchange_rate
            v_brush_plat = abs(metrics['brush_platform_fee_php']) * exchange_rate
            brush_commission = metrics['brush_commission']
            brush_empty_pkg = metrics['brush_empty_pkg']
        else:
            v_brush_amt = 0
            brush_commission = 0
            v_brush_plat = 0
            brush_empty_pkg = 0

        # 平台费用：菲律宾沿用原列范围；马来/印尼按各自账单费用区间扣除已提取值。
        v_shopee_fee = abs(metrics['total_platform_fee_php']) * exchange_rate - v_brush_plat

        # --- C. 处理"利润表" ---
        pnl_sheet_name = None
        for n in SHEET_NAMES['pnl']:
            if n in wb.sheetnames: pnl_sheet_name = n; break

        if pnl_sheet_name:
            ws_pnl = wb[pnl_sheet_name]

            # 找到本月目标列（按第2行月份表头匹配，找不到则新增列）
            target_month_header = f"{cur_month}月"
            target_col = None
            for col_i in range(2, ws_pnl.max_column + 20):
                cell_val = ws_pnl.cell(row=2, column=col_i).value
                if cell_val is not None and str(cell_val).strip() == target_month_header:
                    target_col = col_i
                    break
                if cell_val is None or str(cell_val).strip() == '':
                    target_col = col_i
                    ws_pnl.cell(row=2, column=col_i).value = target_month_header
                    break
            if target_col is None:
                target_col = ws_pnl.max_column + 1
                ws_pnl.cell(row=2, column=target_col).value = target_month_header
            target_col_letter = get_column_letter(target_col)

            sn = sales_sheet_name or "销售明细"
            val_map = {
                '本期销售': f'=SUM({sn}!J2:J{last_sales_row})+ABS(SUM({sn}!H2:H{last_sales_row}))*{exchange_rate}',
                '其中:补单金额': v_brush_amt,
                '补单金额': v_brush_amt,
                '销售退款': f'=ABS(SUM({sn}!H2:H{last_sales_row}))*{exchange_rate}',
                '实际产品成本': f'=SUM({sn}!K2:K{last_sales_row})',
                '发单费': f'=({metrics["order_count"]}+{metrics["cancel_shipped"]}-{metrics["brush_count"]})*{UNIT_FEE}',
                '广告费': metrics['ads_cny'],
                '补单佣金': brush_commission,
                '补单平台费用': v_brush_plat,
                '补单汇率差额': metrics['brush_diff_value'],
                '补单空包费': brush_empty_pkg,
                '平台费用': v_shopee_fee,
                '预扣税': v_tax,
                '团队运营费用': 0,
            }

            row_idx_map = {}

            search_keys = [
                '实际主营收入', '其中:补单金额', '补单金额', '实际主营成本', '实际产品成本',
                '团队运营费用', '补单平台费用', '补单汇率差额', '补单空包费',
                'shopee结算', '本期销售', '销售退款', '营业费用', '推广费用',
                '补单佣金', '平台费用', '毛利润', '本期利润', '发单费', '广告费', '预扣税'
            ]

            # 从第3行开始扫描（跳过A1标题行和A2表头行，保护A1格式不被修改）
            for row in ws_pnl.iter_rows(min_row=3, max_row=100):
                cell_a = row[0]
                cell_val_str = str(cell_a.value).strip() if cell_a.value else ""

                for key in search_keys:
                    if key.lower() in cell_val_str.lower():
                        row_idx_map[key] = cell_a.row
                        if key in val_map:
                            ws_pnl.cell(row=cell_a.row, column=target_col).value = val_map[key]
                        break


            if '其中:补单金额' not in row_idx_map and '补单金额' in row_idx_map:
                row_idx_map['其中:补单金额'] = row_idx_map['补单金额']

            def get_cell(key):
                row_num = row_idx_map.get(key)
                return f"{target_col_letter}{row_num}" if row_num else None

            if '实际主营收入' in row_idx_map:
                terms = [get_cell('本期销售')]
                terms += [f"-{get_cell(k)}" for k in ['其中:补单金额', '销售退款'] if get_cell(k)]
                ws_pnl[get_cell('实际主营收入')] = '=' + ''.join(terms)

            if '实际主营成本' in row_idx_map:
                ws_pnl[get_cell('实际主营成本')] = f"={get_cell('实际产品成本')}+{get_cell('发单费')}"

            if '推广费用' in row_idx_map:
                sub_items = [get_cell(k) for k in ['广告费', '补单佣金', '补单平台费用', '补单汇率差额', '补单空包费'] if k in row_idx_map]
                if sub_items:
                    ws_pnl[get_cell('推广费用')] = f"={'+'.join(sub_items)}"

            if 'shopee结算' in row_idx_map:
                sub_items = [get_cell(k) for k in ['平台费用', '预扣税'] if get_cell(k)]
                if sub_items:
                    ws_pnl[get_cell('shopee结算')] = f"={'+'.join(sub_items)}"

            if '营业费用' in row_idx_map:
                sub_items = [get_cell(k) for k in ['推广费用', 'shopee结算', '团队运营费用'] if get_cell(k)]
                if sub_items:
                    ws_pnl[get_cell('营业费用')] = f"={'+'.join(sub_items)}"

            profit_key = '本期利润' if '本期利润' in row_idx_map else '毛利润'
            if profit_key in row_idx_map:
                ws_pnl[get_cell(profit_key)] = f"={get_cell('实际主营收入')}-{get_cell('实际主营成本')}-{get_cell('营业费用')}"

            # 备注：按月追加“{月}月未匹配”，保留以前月份的备注
            note = (f"{cur_month}月未匹配：取消已发货{metrics['cancel_shipped']}，"
                    f"取消未发货{metrics['cancel_unshipped']}")
            note_row = None
            start = row_idx_map.get(profit_key, 2) + 1
            for r in range(start, ws_pnl.max_row + 2):
                lab = str(ws_pnl.cell(r, 1).value or '').strip()
                if lab == '' or lab.startswith(f"{cur_month}月未匹配") or lab.startswith('未匹配说明'):
                    note_row = r
                    break
            ws_pnl.cell(note_row, 1).value = note

        # --- D. 处理"未匹配" ---
        un_sheet_name = None
        for n in SHEET_NAMES['unmatched']:
            if n in wb.sheetnames: un_sheet_name = n; break

        if not df_unmatched.empty:
            if un_sheet_name:
                ws_un = wb[un_sheet_name]
                if ws_un.max_row >= 1:
                    ws_un.delete_rows(1, ws_un.max_row)
            else:
                ws_un = wb.create_sheet('未匹配订单')

            # 写表头，末尾加"取消已发货成本"列
            headers = df_unmatched.columns.tolist() + ['取消已发货成本']
            ws_un.append(headers)

            # 逐行写入，取消已发货行：填成本值 + 黄色背景
            for _, un_row in df_unmatched.iterrows():
                row_data = ['' if pd.isna(un_row[c]) else un_row[c] for c in df_unmatched.columns]
                is_cancel_shipped = str(un_row.get('Check_Status', '')).strip() == 'Shipped'

                if is_cancel_shipped:
                    sku = str(un_row.get(sku_col, '')).strip()
                    qty = clean_currency(un_row.get(qty_col, 1))
                    cost_val = round(get_line_cost(sku, qty), 2)
                    row_data.append(f"取消已发货成本：{cost_val}")
                else:
                    row_data.append('')

                ws_un.append(row_data)

                if is_cancel_shipped:
                    cur_row = ws_un.max_row
                    for c in range(1, len(row_data) + 1):
                        ws_un.cell(row=cur_row, column=c).fill = yellow_fill

        wb.save(file_output)
        wb.close()
        print("   [成功] 处理完成")
        return True

    except Exception as e:
        print(f"   [失败] 处理出错: {e}")
        return False

# ================= 3. 主入口 =================

def main():
    print(f"启动批量处理... 目标文件夹: {TARGET_FOLDER}")
    if TARGET_MONTH == 0:
        print("当前设置分析时间: 自动判断月份（按各文件订单量最多的年月）")
    else:
        print(f"当前设置分析时间: {TARGET_YEAR}年 {TARGET_MONTH}月")

    if not os.path.exists(TARGET_FOLDER):
        print(f"错误: 找不到文件夹 '{TARGET_FOLDER}'")
        return

    # 读取成本表
    cost_path = os.path.join(TARGET_FOLDER, COST_FILE_NAME)
    if not os.path.exists(cost_path):
        print(f"错误: 找不到成本表 '{cost_path}'")
        return

    print(">>> 加载商品成本表...")
    try:
        df_cost = pd.read_excel(cost_path, dtype=str)
        sku_col = next((c for c in df_cost.columns if 'sku' in c.lower()), 'SKU')
        cost_col = next((c for c in df_cost.columns if '成本' in c), '人民币成本')
        cost_map = dict(zip(df_cost[sku_col].astype(str).str.strip(), df_cost[cost_col].apply(clean_currency)))
    except Exception as e:
        print(f"   读取成本表失败: {e}")
        return

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    xlsx_files = glob.glob(os.path.join(TARGET_FOLDER, "*.xlsx"))
    count = 0
    for file_path in xlsx_files:
        file_name = os.path.basename(file_path)
        if (file_name == COST_FILE_NAME) or ("新_" in file_name) or (file_name.startswith("~$")):
            continue
        if process_single_file(file_path, cost_map):
            count += 1

    print(f"\n{'='*60}")
    print(f"批量处理完成！共 {count} 个文件。")
    print(f"新生成的文件存放在: {OUTPUT_FOLDER}")
    try:
        input("按回车键退出...")
    except EOFError:
        pass

if __name__ == '__main__':
    main()
