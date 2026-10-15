import pandas as pd
import numpy as np
import os
import shutil
import warnings
import openpyxl
import time
import glob
import stat
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 忽略警告
warnings.filterwarnings('ignore')

# 脚本所在目录（自动定位，无需手动修改）
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ================= 配置区域 =================

# 1. 顺昌数据文件夹（自动定位到脚本同级的 TKdeal_Shunchang 子文件夹）
TARGET_FOLDER = os.path.join(SCRIPT_DIR, 'TKdeal_Shunchang')

# 2. 成本表文件名（位于脚本同级目录，不在 TKdeal 内）
COST_FILE_NAME = '商品成本.xlsx'

# 3. 参数配置
UNIT_FEE = 3.0              # 发单费/空包费单价

# 3a. 多国汇率配置（脚本根据文件名关键词自动匹配，单位：1本国货币 = ? CNY）
#     若文件名不含以下关键词，则使用 DEFAULT_EXCHANGE_RATE
EXCHANGE_RATE_MAP = {
    '菲律宾': 0.113474,   # 1 PHP ≈ 0.113474 CNY
    '印尼':   0.000403006,  # 1 IDR ≈ 0.000403006 CNY
    '马来':   1.68848,     # 1 MYR ≈ 1.52 CNY
    '泰国':   0.20,     # 1 THB ≈ 0.20 CNY
    '越南':   0.000028, # 1 VND ≈ 0.000028 CNY
    '新加坡': 5.28,     # 1 SGD ≈ 5.28 CNY
}
DEFAULT_EXCHANGE_RATE = 0.1178  # 未匹配时的默认汇率（默认按 PHP）

# 4. 输出文件夹（自动创建在脚本同级目录，不污染 TKdeal_Shunchang 原始数据）
OUTPUT_FOLDER = os.path.join(SCRIPT_DIR, 'output_Shunchang')

# 5. 时间筛选配置 (可随时修改)
TARGET_YEAR = 2026      # 目标年份
TARGET_MONTH = 3        # 目标月份 (1-12)

# Sheet 名称映射
SHEET_NAMES = {
    'order':     ['订单明细表', '订单明细原表', '订单明细'],          # TK: 订单明细表
    'income':    ['账单原表', '收入明细表', '收入明细'],              # TK: 账单原表
    'brush':     ['补单明细', '补单'],
    'ads':       ['广告费', '广告'],
    'sales':     ['销售明细表', '销售明细'],                          # TK: 销售明细表
    'pnl':       ['利润表', 'Profit'],
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

def clean_currency(x):
    if pd.isna(x): return 0.0
    s = str(x).replace('PHP', '').replace('RM', '').replace(',', '').replace(' ', '').strip()
    try: return float(s)
    except: return 0.0

def safe_border(ws, row_idx, col_idx):
    """安全地给单元格加边框"""
    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    cell = ws.cell(row=row_idx, column=col_idx)
    cell.border = border
    cell.alignment = Alignment(horizontal='center', vertical='center')

# ================= 2. 核心处理逻辑 =================

def get_exchange_rate(file_name):
    """根据文件名关键词自动匹配汇率"""
    for keyword, rate in EXCHANGE_RATE_MAP.items():
        if keyword in file_name:
            return rate
    return DEFAULT_EXCHANGE_RATE

def copy_as_empty(source_path, file_name):
    """当无有效数据时，仅复制原表并增加前缀作为结果"""
    current_time = time.strftime("%H点%M分%S秒")
    new_file_name = f"新_{file_name.replace('.xlsx', '')}_{current_time}.xlsx"
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    file_output = os.path.join(OUTPUT_FOLDER, new_file_name)
    try:
        shutil.copy(source_path, file_output)
        os.chmod(file_output, stat.S_IWRITE)
    except Exception as e:
        print(f"   [错误] 空文件生成副本失败: {e}")

def process_single_file(source_path, cost_map):
    file_name = os.path.basename(source_path)
    EXCHANGE_RATE = get_exchange_rate(file_name)  # 按文件名自动选汇率
    print(f"\n{'='*60}")
    print(f">>> 正在处理: {file_name}  (汇率: {EXCHANGE_RATE})")

    # 1. 读取数据
    xls = pd.ExcelFile(source_path)
    # TK 订单关键词: 'Order ID' 和 'Seller SKU' 均存在
    df_order = load_and_clean_sheet(xls, SHEET_NAMES['order'], ['Order ID', 'Seller SKU'])
    # TK 账单关键词: 'Order/adjustment ID' 和 'Total settlement amount'
    df_income = load_and_clean_sheet(xls, SHEET_NAMES['income'],
                                     ['Order/adjustment ID', 'Total settlement amount'])
    # 兼容 Shopee 收入表
    if df_income.empty:
        df_income = load_and_clean_sheet(xls, SHEET_NAMES['income'], ['Order ID', 'Total Released Amount'])
    df_brush = load_and_clean_sheet(xls, SHEET_NAMES['brush'], ['线上单号', '佣金'])
    df_ads = load_and_clean_sheet(xls, SHEET_NAMES['ads'], ['花费'])

    if df_order.empty or df_income.empty:
        print("   [跳过] 缺少订单或收入明细表，正在直接生成空原样副本")
        copy_as_empty(source_path, file_name)
        return

    # 将账单的 'Order/adjustment ID  '(含空格) 标准化为 'Order ID'
    inc_id_col = next((c for c in df_income.columns if 'Order' in c and ('ID' in c or 'adjustment' in c)), None)
    if inc_id_col and inc_id_col != 'Order ID':
        df_income = df_income.rename(columns={inc_id_col: 'Order ID'})
    # 去掉非订单行（如 adjustment、refund 等）——保留 Type==Order 的行
    if 'Type ' in df_income.columns or 'Type' in df_income.columns:
        type_col = 'Type ' if 'Type ' in df_income.columns else 'Type'
        df_income = df_income[df_income[type_col].astype(str).str.strip() == 'Order'].copy()
    # 最终校验：若 Order ID 仍不存在（表头识别失败），视为无账单数据，跳过
    if 'Order ID' not in df_income.columns:
        print(f"   [警告] 账单原表表头识别失败（实际列: {list(df_income.columns[:4])}），跳过并生成原样副本")
        copy_as_empty(source_path, file_name)
        return

    # 2. 数据计算 — 日期列定位（TK: 'Created Time', Shopee: 'Order Creation Date'/'Order created time'）
    date_col = next((c for c in df_order.columns if c in ('Created Time', 'Order Creation Date', 'Order created time')), df_order.columns[0])

    df_order['Created Time'] = pd.to_datetime(df_order[date_col], errors='coerce')
    mask_order = (df_order['Created Time'].dt.year == TARGET_YEAR) & (df_order['Created Time'].dt.month == TARGET_MONTH)
    df_order_dec = df_order[mask_order].copy()
    print(f"   {TARGET_MONTH}月订单数: {len(df_order_dec)}")

    if len(df_order_dec) == 0:
        print(f"   [跳过] {TARGET_MONTH}月无订单，将直接生成空原样副本")
        copy_as_empty(source_path, file_name)
        return

    # 收入明细时间筛选（TK 账单: 'Order created time'，Shopee: 'Order Creation Date'）
    inc_date_col = next((c for c in df_income.columns
                         if c.strip() in ('Order created time', 'Order Creation Date')), None)
    if inc_date_col:
        df_income['Created Time'] = pd.to_datetime(df_income[inc_date_col], errors='coerce')
        mask_income = (df_income['Created Time'].dt.year == TARGET_YEAR) & (df_income['Created Time'].dt.month == TARGET_MONTH)
        df_income_dec = df_income[mask_income].copy()
    else:
        df_income_dec = df_income.copy()  # 无时间列则全量使用


    sku_col = next((c for c in df_order_dec.columns if 'Seller SKU' in c), None)
    if not sku_col: sku_col = next((c for c in df_order_dec.columns if 'Reference' in c and 'Parent' not in c), None)
    if not sku_col: sku_col = next((c for c in df_order_dec.columns if 'SKU' in c), 'Seller SKU')
    qty_col = 'Quantity'
    # TK 物流单号: 'Tracking ID'， Shopee: 'Tracking Number*'
    track_col = next((c for c in df_order_dec.columns if 'Tracking' in c), 'Tracking ID')
    
    def get_row_cost(row):
        sku = str(row[sku_col]).strip()
        qty = clean_currency(row[qty_col])
        return cost_map.get(sku, 0) * qty

    df_order_dec['Line_Total_Cost'] = df_order_dec.apply(get_row_cost, axis=1)
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

    cols = df_income.columns
    # TK 账单列名， Shopee 列名兼容
    c_price   = next((c for c in cols if 'Subtotal before discounts' in c
                      or 'Original Product Price' in c), None)
    c_promo   = next((c for c in cols if 'Seller discounts' in c
                      and 'co-funded' not in c and 'Refund' not in c
                      or 'Seller Product Promotion' in c), None)
    c_voucher = next((c for c in cols if 'Seller co-funded voucher discount' in c
                      and 'Refund' not in c
                      or 'Voucher Sponsored by Seller' in c), None)
    c_refund  = next((c for c in cols if c.strip() == 'Customer refund'
                      or 'Refund Amount' in c), None)
    # TK 账单原表同一订单可能有多行，去重后建索引（保留最后一条，金额最完整）
    if 'Order ID' in df_income.columns:
        income_dict = df_income.drop_duplicates(subset='Order ID', keep='last').set_index('Order ID').to_dict('index')
    else:
        print(f"   [警告] 账单原表未找到 Order ID 列（实际列名: {list(df_income.columns[:5])}），收入金额将为 0")
        income_dict = {}
    
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

    ids_order = set(df_order_dec['Order ID'])
    ids_income = set(df_income['Order ID'])
    unmatched_ids = ids_order - ids_income
    df_unmatched = df_order_dec[df_order_dec['Order ID'].isin(unmatched_ids)].copy()
    
    count_cancel_shipped = 0
    count_cancel_unshipped = 0
    if not df_unmatched.empty:
        def check_status(row):
            track = str(row.get(track_col, ''))
            status = str(row.get(status_col, ''))
            if 'Cancel' in status:
                return 'Shipped' if len(track) > 5 and track.lower() != 'nan' else 'Unshipped'
            return 'Other'
        df_unmatched['Check_Status'] = df_unmatched.apply(check_status, axis=1)
        count_cancel_shipped = len(df_unmatched[df_unmatched['Check_Status'] == 'Shipped'])
        count_cancel_unshipped = len(df_unmatched[df_unmatched['Check_Status'] == 'Unshipped'])

    brushing_ids = set()
    if not df_brush.empty:
        brush_id_col = next((c for c in df_brush.columns if '单号' in c), df_brush.columns[0])
        brushing_ids = set(df_brush[brush_id_col].astype(str).str.strip())
    
    c_rebate = next((c for c in cols if 'Rebate' in c and 'Shopee' in c), None)
    cols_brush_amt = [c for c in [c_price, c_promo, c_voucher, c_refund, c_rebate] if c]
    
    all_cols = list(df_income.columns)
    try:
        # TK: 平台费用用 'Total Fees' 一列表示； Shopee: 范围列汇总
        if 'Total Fees' in all_cols:
            cols_platform_fee_range = ['Total Fees']
        else:
            idx_start = next(i for i, c in enumerate(all_cols) if 'Seller shipping fee' in c or 'Shipping Fee Paid by Seller' in c)
            idx_end   = next(i for i, c in enumerate(all_cols) if 'Withholding Tax' in c)
            cols_platform_fee_range = all_cols[idx_start : idx_end+1]
    except:
        cols_platform_fee_range = [c for c in all_cols if 'Fee' in c or 'Tax' in c or 'Commission' in c]
    


    brush_refund_sum_php = 0
    brush_platform_fee_php = 0
    for bid in brushing_ids:
        if bid in income_dict:
            row = income_dict[bid]
            for c in cols_brush_amt: brush_refund_sum_php += clean_currency(row.get(c, 0))
            for c in cols_platform_fee_range: brush_platform_fee_php += clean_currency(row.get(c, 0))

    cols_pure_platform = [c for c in cols_platform_fee_range if 'Tax' not in c]
    total_platform_fee_php = sum(df_income_dec[c].apply(clean_currency).sum() for c in cols_pure_platform if c in df_income_dec.columns)
    c_tax = next((c for c in all_cols if 'Withholding Tax' in c), None)
    total_tax_php = df_income_dec[c_tax].apply(clean_currency).sum() if c_tax else 0

    ads_col = next((c for c in df_ads.columns if '花费' in c or 'Expense' in c), None)
    ads_cny = df_ads[ads_col].apply(clean_currency).sum() * EXCHANGE_RATE if ads_col else 0
    
    brush_comm = 0; brush_diff = 0
    if not df_brush.empty:
        c_comm = next((c for c in df_brush.columns if '佣金' in c), None)
        c_diff = next((c for c in df_brush.columns if '汇差' in c), None)
        if c_comm: brush_comm = df_brush[c_comm].apply(clean_currency).sum()
        if c_diff: brush_diff = df_brush[c_diff].apply(clean_currency).sum()

    metrics = {
        'total_platform_fee_php': total_platform_fee_php,
        'total_tax_php': total_tax_php,
        'brush_amt_php': brush_refund_sum_php,
        'brush_platform_fee_php': brush_platform_fee_php,
        'brush_commission': brush_comm,
        'brush_diff': brush_diff,
        'brush_empty_pkg': len(brushing_ids) * UNIT_FEE,
        'ads_cny': ads_cny,
        'cancel_shipped': count_cancel_shipped,
        'cancel_unshipped': count_cancel_unshipped,
        'order_count': len(df_sales_detail)
    }

    # ================= 3. 写入文件 =================
    
    current_time = time.strftime("%H点%M分%S秒")
    new_file_name = f"新_{file_name.replace('.xlsx', '')}_{current_time}.xlsx"
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)  # 确保输出目录存在
    file_output = os.path.join(OUTPUT_FOLDER, new_file_name)

    print(f"   正在生成副本并进行写表操作...")
    try:
        shutil.copy(source_path, file_output)
        os.chmod(file_output, stat.S_IWRITE)
        time.sleep(1)
    except Exception as e:
        print(f"   [错误] 复制失败: {e}")
        return

    try:
        wb = openpyxl.load_workbook(file_output)
        
        # --- A. 处理“销售明细” ---
        curr_sym = next((sym for key, sym in {'菲律宾':'₱', '印尼':'Rp', '马来':'RM', '泰国':'฿', '越南':'₫', '新加坡':'S$'}.items() if key in file_name), '外币')
        sales_sheet_name = None
        for n in SHEET_NAMES['sales']:
            if n in wb.sheetnames: sales_sheet_name = n; break
        
        if not sales_sheet_name:
            ws_sales = wb.create_sheet('销售明细表')
            headers = ['日期', '订单号', '商品SKU', '销售数量', f'销售单价({curr_sym})', f'优惠金额({curr_sym})', f'退款金额({curr_sym})', 
                       f'销售金额({curr_sym})', '销售金额(¥)', '商品成本总额(¥)', '毛利(¥)']
            ws_sales.append(headers)
        else:
            ws_sales = wb[sales_sheet_name]
            max_row = ws_sales.max_row
            # 【修复点 1】：放弃极其耗时的单元格迭代清空，直接使用 delete_rows 瞬间删除旧的几十万行幽灵数据
            if max_row >= 2:
                ws_sales.delete_rows(2, max_row)
        
        for i, row in df_sales_detail.iterrows():
            r = i + 2
            ws_sales.cell(r, 1).value = str(row['日期'])[:10]
            ws_sales.cell(r, 2).value = row['订单号']
            ws_sales.cell(r, 3).value = row['商品SKU']
            ws_sales.cell(r, 4).value = row['销售数量']
            ws_sales.cell(r, 5).value = row['销售单价']
            ws_sales.cell(r, 6).value = row['优惠金额'] + row['优惠劵金额']
            ws_sales.cell(r, 7).value = row['退款金额']
            
            ws_sales.cell(r, 8).value = f'=E{r}+F{r}+G{r}'
            ws_sales.cell(r, 9).value = f'=H{r}*{EXCHANGE_RATE}'
            ws_sales.cell(r, 10).value = row['商品成本总额(CNY)']
            ws_sales.cell(r, 11).value = f'=I{r}-J{r}'
            
            for c in range(1, 12): safe_border(ws_sales, r, c)

        last_sales_row = len(df_sales_detail) + 1

        # --- B. 处理“补单明细” ---
        brush_sheet_name = None
        for n in SHEET_NAMES['brush']:
            if n in wb.sheetnames: brush_sheet_name = n; break

        v_brush_amt = metrics['brush_amt_php'] * EXCHANGE_RATE
        v_shopee_fee = abs(metrics['total_platform_fee_php']) * EXCHANGE_RATE 
        v_tax = abs(metrics['total_tax_php']) * EXCHANGE_RATE                  
        v_brush_plat = abs(metrics['brush_platform_fee_php']) * EXCHANGE_RATE
        
        brush_commission = metrics['brush_commission']
        brush_empty_pkg = metrics['brush_empty_pkg']
        brush_diff_expr = metrics['brush_diff']

        if brush_sheet_name and not df_brush.empty:
            ws_brush = wb[brush_sheet_name]
            
            header_row = None
            col_idx_xiaoji = None
            col_idx_chengben = None
            col_idx_huicha = None

            for row in ws_brush.iter_rows(min_row=1, max_row=20):
                for cell in row:
                    if cell.value is not None:
                        val_str = str(cell.value).strip()
                        if '垫付小计' in val_str: col_idx_xiaoji = cell.column
                        elif '垫付成本' in val_str: col_idx_chengben = cell.column
                        elif '汇率差' in val_str or '汇差' in val_str: col_idx_huicha = cell.column
                if col_idx_xiaoji and col_idx_chengben:
                    header_row = row[0].row
                    break

            if header_row and col_idx_xiaoji and col_idx_chengben:
                if not col_idx_huicha:
                    col_idx_huicha = ws_brush.max_column + 1
                    ws_brush.cell(row=header_row, column=col_idx_huicha, value="汇率差")

                col_let_xiaoji = get_column_letter(col_idx_xiaoji)
                col_let_chengben = get_column_letter(col_idx_chengben)
                col_let_huicha = get_column_letter(col_idx_huicha)

                max_brush_row = ws_brush.max_row
                for r in range(header_row + 1, max_brush_row + 1):
                    if ws_brush.cell(row=r, column=col_idx_xiaoji).value is not None or ws_brush.cell(row=r, column=col_idx_chengben).value is not None:
                        ws_brush.cell(row=r, column=col_idx_huicha).value = f'={col_let_xiaoji}{r}-{EXCHANGE_RATE}*{col_let_chengben}{r}'
                
                brush_diff_expr = f'=SUM({brush_sheet_name}!{col_let_huicha}{header_row+1}:{col_let_huicha}{max_brush_row})'
        else:
            v_brush_amt = 0
            brush_commission = 0
            v_brush_plat = 0
            brush_diff_expr = 0
            brush_empty_pkg = 0

        # --- C. 处理“利润表” ---
        pnl_sheet_name = None
        for n in SHEET_NAMES['pnl']:
            if n in wb.sheetnames: pnl_sheet_name = n; break
        
        if pnl_sheet_name:
            ws_pnl = wb[pnl_sheet_name]
            
            # 定位填报的列名位置：根据“项目”所在行，寻找“X月”所在的列
            target_col_idx = 2  # 默认在第二列(B列)
            for r_idx, row in enumerate(ws_pnl.iter_rows(min_row=1, max_row=5), start=1):
                row_vals = [str(c.value).strip() if c.value else "" for c in row]
                if any("项目" in val for val in row_vals):
                    for cell in row:
                        if cell.value and f"{TARGET_MONTH}月" in str(cell.value).replace(' ', ''):
                            target_col_idx = cell.column
                            break
                    break
            target_col_letter = get_column_letter(target_col_idx)

            val_map = {
                '本期销售': f'=SUM({sales_sheet_name or "销售明细表"}!I2:I{last_sales_row})',
                '其中:补单金额': v_brush_amt,
                '销售退款': f'=ABS(SUM({sales_sheet_name or "销售明细表"}!G2:G{last_sales_row}))*{EXCHANGE_RATE}',
                '实际产品成本': f'=SUM({sales_sheet_name or "销售明细表"}!J2:J{last_sales_row})',
                '发单费': f'=({metrics["order_count"]}+{metrics["cancel_shipped"]})*{UNIT_FEE}',
                '广告费': metrics['ads_cny'],
                '补单佣金': brush_commission,
                '补单平台费用': v_brush_plat,
                '补单汇率差额': brush_diff_expr,
                '补单空包费': brush_empty_pkg,
                '平台费用': v_shopee_fee,
                '调整金额': 0, # 新增或可选预留
                '预扣税': v_tax,
                '团队运营费用': 0,
            }
            
            row_idx_map = {}
            for row in ws_pnl.iter_rows(min_row=1, max_row=100): # 限制读取范围防卡死
                cell_a = row[0]
                cell_val_str = str(cell_a.value).strip() if cell_a.value else ""
                
                for key in ['实际主营收入', '本期销售', '其中:补单金额', '销售退款', 
                            '实际主营成本', '实际产品成本', '发单费',
                            '营业费用', '推广费用', '广告费', '补单佣金', '补单平台费用', '补单汇率差额', '补单空包费',
                            'TK结算', 'shopee结算', '平台费用', '调整金额', '预扣税', '团队运营费用', '毛利润', '本期利润']:
                    if key in cell_val_str:
                        row_idx_map[key] = cell_a.row
                        if key in val_map:
                            row[target_col_idx - 1].value = val_map[key]
                            
                if "未匹配" in cell_val_str or "补充" in cell_val_str:
                    cell_a.value = f"未匹配说明：取消已发货 {metrics['cancel_shipped']} 单，取消未发货 {metrics['cancel_unshipped']} 单"

            def get_cell(key):
                r = row_idx_map.get(key)
                return f"{target_col_letter}{r}" if r else "0"
            
            if '实际主营收入' in row_idx_map:
                ws_pnl[get_cell('实际主营收入')] = f"={get_cell('本期销售')}-{get_cell('其中:补单金额')}-{get_cell('销售退款')}"
            
            if '实际主营成本' in row_idx_map:
                ws_pnl[get_cell('实际主营成本')] = f"={get_cell('实际产品成本')}+{get_cell('发单费')}"
                
            if '推广费用' in row_idx_map:
                sub_items = [get_cell(k) for k in ['广告费', '补单佣金', '补单平台费用', '补单汇率差额', '补单空包费'] if k in row_idx_map]
                ws_pnl[get_cell('推广费用')] = f"={'+'.join(sub_items)}"
                
            # 兼容：如果有shopee结算则用shopee结算，如果是TK结算则用TK结算
            settle_key = 'TK结算' if 'TK结算' in row_idx_map else ('shopee结算' if 'shopee结算' in row_idx_map else None)
            if settle_key:
                settle_items = [get_cell(k) for k in ['平台费用', '调整金额', '预扣税'] if k in row_idx_map]
                if settle_items:
                    ws_pnl[get_cell(settle_key)] = f"={'+'.join(settle_items)}"
                
            if '营业费用' in row_idx_map:
                biz_fee_items = [get_cell('推广费用'), get_cell('团队运营费用')]
                if settle_key: biz_fee_items.append(get_cell(settle_key))
                ws_pnl[get_cell('营业费用')] = f"={'+'.join(biz_fee_items)}"
                
            profit_key = '本期利润' if '本期利润' in row_idx_map else '毛利润'
            if profit_key in row_idx_map:
                ws_pnl[get_cell(profit_key)] = f"={get_cell('实际主营收入')}-{get_cell('实际主营成本')}-{get_cell('营业费用')}"

        # --- D. 处理“未匹配” ---
        un_sheet_name = None
        for n in SHEET_NAMES['unmatched']:
            if n in wb.sheetnames: un_sheet_name = n; break
            
        if not df_unmatched.empty:
            if un_sheet_name:
                ws_un = wb[un_sheet_name]
                # 【修复点 2】：同样放弃死循环清理，直接删掉整表内容，再重新写入表头
                if ws_un.max_row >= 1:
                    ws_un.delete_rows(1, ws_un.max_row)
            else:
                ws_un = wb.create_sheet('未匹配订单')
                
            headers = df_unmatched.columns.tolist()
            ws_un.append(headers)
            for r in dataframe_to_rows(df_unmatched, index=False, header=False):
                ws_un.append(r)

        wb.save(file_output)
        wb.close()
        print("   [成功] 处理完成")

    except Exception as e:
        print(f"   [失败] 处理出错: {e}")

# ================= 3. 主入口 =================

def main():
    print(f"启动批量处理...")
    print(f"  数据来源: {TARGET_FOLDER}")
    print(f"  输出目录: {OUTPUT_FOLDER}")
    print(f"  当前设置分析时间: {TARGET_YEAR}年 {TARGET_MONTH}月")
    
    if not os.path.exists(TARGET_FOLDER):
        print(f"错误: 找不到 TKdeal_Shunchang 文件夹 '{TARGET_FOLDER}'")
        print(f"  请确认脚本与 TKdeal_Shunchang 文件夹在同一目录下。")
        input("按回车键退出...")
        return

    # 成本表从 TKdeal_Shunchang 文件夹内读取
    cost_path = os.path.join(TARGET_FOLDER, COST_FILE_NAME)
    if not os.path.exists(cost_path):
        print(f"错误: 找不到成本表 '{cost_path}'")
        print(f"  请将 '{COST_FILE_NAME}' 放在 TKdeal_Shunchang 文件夹内: {TARGET_FOLDER}")
        input("按回车键退出...")
        return

    print(f">>> 加载商品成本表: {cost_path}")
    try:
        df_cost = pd.read_excel(cost_path, dtype=str)
        sku_col = next((c for c in df_cost.columns if 'sku' in c.lower()), 'SKU')
        cost_col = next((c for c in df_cost.columns if '成本' in c), '人民币成本')
        cost_map = dict(zip(df_cost[sku_col].astype(str).str.strip(), df_cost[cost_col].apply(clean_currency)))
        print(f"   已加载 {len(cost_map)} 条 SKU 成本记录")
    except Exception as e:
        print(f"   读取成本表失败: {e}")
        input("按回车键退出...")
        return

    # 搜索 TKdeal 内所有 xlsx，跳过成本表和临时文件
    xlsx_files = sorted(glob.glob(os.path.join(TARGET_FOLDER, "*.xlsx")))
    count = 0
    for file_path in xlsx_files:
        file_name = os.path.basename(file_path)
        if file_name == COST_FILE_NAME or file_name.startswith("~$"):  # 跳过成本表和临时锁定文件
            continue
        process_single_file(file_path, cost_map)
        count += 1

    print(f"\n{'='*60}")
    print(f"批量处理完成！共处理了 {count} 个文件。")
    print(f"新生成的文件存放在: {OUTPUT_FOLDER}")
    input("按回车键退出...")

if __name__ == '__main__':
    main()