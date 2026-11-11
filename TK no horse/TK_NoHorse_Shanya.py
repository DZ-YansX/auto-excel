"""TK 山亚月度利润表生成脚本。

用法：把各店铺 xlsx 和「山亚成本价.xlsx」放进 TKdeal_Shanya，运行本脚本，
结果写到 output_Shanya/新_{原文件名}.xlsx（原文件不动）。

每个店铺文件需要的 sheet：
    订单明细原表  TikTok 导出的订单明细（英文表头，含 Order ID / Seller SKU）
    收入明细表    TikTok 导出的账单（中文表头，含 订单ID/调整单ID / 交易类型）
    广告费        表头含「广告」或「花费」的一列，金额为美元
    销售明细 / 利润表 / 未匹配   模板 sheet，由脚本填写

口径：
    - 只统计「下单时间在目标月」的订单；账单只取「交易类型=订单」的行。
    - 销售明细：订单在账单里有销售额才列出；同一订单的多行账单金额累加。
    - 成本：按 Seller SKU 查成本表，「基础SKU*N」按基础成本×N，匹配时忽略空格；查不到记 0。
    - 未匹配：订单明细里有、账单里没有的订单。其中已取消但已发货的，
      商品成本计入实际产品成本，并计发单费。
    - 利润表按「{月}月」列写入，已有该列则覆盖，没有则新增一列，历史月份不动。
"""
import glob
import os
import shutil
import stat
import warnings

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ================= 配置 =================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(SCRIPT_DIR, 'TKdeal_Shanya')
OUTPUT_FOLDER = os.path.join(SCRIPT_DIR, 'output_Shanya')
COST_FILE_NAME = '山亚成本价.xlsx'

# 汇率：1 外币 = ? CNY。按账单「货币」列选用，每月处理前更新。
CURRENCY_RATES = {
    'THB': 0.1973,
}
USD_RATE = 7.0      # 广告费按美元结算
UNIT_FEE = 3.5      # 发单费单价（CNY/单）

# 目标月份：都为 0 时按订单日期自动识别（取出现最多的年月）；否则强制指定。
TARGET_YEAR = 0
TARGET_MONTH = 0

# 订单明细与账单所在的 sheet（按表头内容区分，不依赖 sheet 名，人工放反也能识别）
DATA_SHEETS = ('订单明细原表', '订单明细表', '收入明细表', '账单原表')
SALES_SHEETS = ('销售明细', '销售明细表')
PNL_SHEET = '利润表'
UNMATCHED_SHEET = '未匹配'
ADS_SHEET = '广告费'

# 账单（中文表头）用到的列
BILL_ID = '订单ID/调整单ID'
BILL_TYPE = '交易类型'
BILL_ORDER_TYPE = '订单'
BILL_CREATED = '订单创建时间'
BILL_CURRENCY = '货币'
BILL_PRICE = '享受折扣前小计'             # → 销售单价
BILL_DISCOUNT = '商家折扣'                # → 优惠金额
BILL_REFUND = '享受商家折扣后的退款小计'   # → 退款金额
BILL_FEES = '总费用'                      # → 平台费用
BILL_REQUIRED = [BILL_ID, BILL_TYPE, BILL_CREATED, BILL_CURRENCY,
                 BILL_PRICE, BILL_DISCOUNT, BILL_REFUND, BILL_FEES]

# 订单明细（英文表头）用到的列
ORDER_REQUIRED = ['Order ID', 'Seller SKU', 'Quantity', 'Order Status', 'Created Time', 'Shipped Time']

# 利润表项目行的识别关键字（按顺序匹配，行标签包含关键字即命中）
PNL_KEYS = ['实际主营收入', '本期销售', '销售退款', '实际主营成本', '实际产品成本', '发单费',
            '营业费用', '推广费用', '广告费', 'TK结算', '平台费用', '调整金额',
            '团队运营费用', '本期利润']


class SkipFile(Exception):
    """文件缺数据、无法处理：原样复制一份到输出目录。"""


# ================= 读取 =================

def to_number(value):
    if pd.isna(value):
        return 0.0
    try:
        return float(str(value).replace(',', '').strip())
    except ValueError:
        return 0.0


def load_cost_map(cost_path):
    """成本表 → {SKU(去空格): 单位成本CNY}"""
    df = pd.read_excel(cost_path, dtype=str)
    sku_col = next(c for c in df.columns if 'sku' in c.lower())
    cost_col = next(c for c in df.columns if '成本' in c)
    return {''.join(str(sku).split()): to_number(cost)
            for sku, cost in zip(df[sku_col], df[cost_col]) if not pd.isna(sku)}


def sku_unit_cost(sku, cost_map):
    """成本表只收录基础 SKU；「基础SKU*N」的成本 = 基础成本 × N。查不到返回 0。"""
    sku = ''.join(str(sku).split())
    if sku in cost_map:
        return cost_map[sku]
    base, star, times = sku.rpartition('*')
    if star and base in cost_map:
        try:
            return cost_map[base] * float(times)
        except ValueError:
            pass
    return 0.0


def load_orders_and_bill(xls):
    """读出订单明细与账单。按表头内容识别，两张表被放反也没关系。"""
    orders = bill = None
    for name in DATA_SHEETS:
        if name not in xls.sheet_names:
            continue
        df = pd.read_excel(xls, sheet_name=name, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        if 'Seller SKU' in df.columns:
            orders = df
        elif BILL_TYPE in df.columns:
            bill = df
    if orders is None or bill is None:
        raise SkipFile('缺少订单明细或账单')
    for label, df, required in (('订单明细', orders, ORDER_REQUIRED), ('账单', bill, BILL_REQUIRED)):
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f'{label}缺少列 {missing}，表头可能变了，请先核对再处理')
    return orders, bill


def load_ads_cny(xls):
    """广告费 sheet 里表头含「广告/花费」的那一列求和（美元）→ CNY。"""
    if ADS_SHEET not in xls.sheet_names:
        return 0.0
    df = pd.read_excel(xls, sheet_name=ADS_SHEET, dtype=str)
    col = next((c for c in df.columns if '广告' in str(c) or '花费' in str(c)), None)
    return df[col].map(to_number).sum() * USD_RATE if col else 0.0


# ================= 计算 =================

def pick_period(created):
    """返回目标 (年, 月)。"""
    if TARGET_YEAR and TARGET_MONTH:
        return TARGET_YEAR, TARGET_MONTH
    periods = created.dropna().dt.to_period('M')
    if periods.empty:
        raise SkipFile('订单明细里没有可识别的下单日期')
    period = periods.mode().iloc[0]
    return period.year, period.month


def in_period(dates, year, month):
    return (dates.dt.year == year) & (dates.dt.month == month)


def compute(orders, bill, cost_map):
    """核心口径都在这里。返回写表所需的全部结果。"""
    # --- 订单明细：取目标月订单，逐行算成本 ---
    # TK 订单明细时间是 DD/MM/YYYY，必须 dayfirst，否则月份错位
    orders['Created Time'] = pd.to_datetime(orders['Created Time'], dayfirst=True, errors='coerce')
    year, month = pick_period(orders['Created Time'])
    orders = orders[in_period(orders['Created Time'], year, month)].copy()
    if orders.empty:
        raise SkipFile(f'{month}月没有订单')
    orders['_qty'] = orders['Quantity'].map(to_number)
    orders['_cost'] = orders['Seller SKU'].map(lambda s: sku_unit_cost(s, cost_map)) * orders['_qty']

    # --- 账单：只要「订单」行；同一订单多行（结算/退款/补扣）金额累加 ---
    bill = bill[bill[BILL_TYPE].astype(str).str.strip() == BILL_ORDER_TYPE].copy()
    currency = str(bill[BILL_CURRENCY].dropna().iloc[0]).strip().upper() if bill[BILL_CURRENCY].notna().any() else ''
    if currency not in CURRENCY_RATES:
        raise ValueError(f'账单货币 {currency!r} 没有配置汇率，请在 CURRENCY_RATES 里补上')
    rate = CURRENCY_RATES[currency]

    amount_cols = [BILL_PRICE, BILL_DISCOUNT, BILL_REFUND, BILL_FEES]
    bill[amount_cols] = bill[amount_cols].apply(lambda col: col.map(to_number))
    bill_by_order = bill.groupby(BILL_ID)[[BILL_PRICE, BILL_DISCOUNT, BILL_REFUND]].sum()

    # 平台费用只算「下单时间在目标月」的账单行
    bill_created = pd.to_datetime(bill[BILL_CREATED], errors='coerce')
    platform_fee_cny = abs(bill.loc[in_period(bill_created, year, month), BILL_FEES].sum()) * rate

    # --- 销售明细：一单一行，账单里有销售额的才列出 ---
    per_order = orders.groupby('Order ID').agg(
        日期=('Created Time', 'first'),
        商品SKU=('Seller SKU', lambda s: '、'.join(str(x) for x in s if not pd.isna(x))),
        销售数量=('_qty', 'sum'),
        成本=('_cost', 'sum'),
    ).reset_index()
    sales = per_order.join(bill_by_order, on='Order ID').fillna(
        {BILL_PRICE: 0.0, BILL_DISCOUNT: 0.0, BILL_REFUND: 0.0})
    sales = sales[sales[BILL_PRICE] > 0.001].sort_values('日期').reset_index(drop=True)

    # --- 未匹配：订单明细里有、账单里没有 ---
    unmatched = orders[~orders['Order ID'].isin(bill[BILL_ID])]
    cancelled = unmatched['Order Status'].astype(str).str.contains('取消|Cancel')
    shipped = unmatched['Shipped Time'].notna() & (unmatched['Shipped Time'].astype(str).str.strip() != '')
    cancel_shipped = unmatched[cancelled & shipped]
    cancel_unshipped = unmatched[cancelled & ~shipped]

    return {
        'month': month,
        'currency': currency,
        'rate': rate,
        'sales': sales,
        'platform_fee_cny': platform_fee_cny,
        'unmatched': unmatched.drop(columns=['_qty', '_cost']),
        'cancel_shipped_count': cancel_shipped['Order ID'].nunique(),
        'cancel_unshipped_count': cancel_unshipped['Order ID'].nunique(),
        # 已发货却被取消：货已寄出，成本照算
        'cancel_shipped_cost': round(cancel_shipped['_cost'].sum(), 4),
    }


# ================= 写表 =================

THIN = Side(border_style='thin', color='000000')
CELL_BORDER = Border(top=THIN, left=THIN, right=THIN, bottom=THIN)
CELL_ALIGN = Alignment(horizontal='center', vertical='center')


def write_sales_sheet(wb, result):
    """重写销售明细，返回 (sheet 名, 最后一行数据的行号)。
    列：A日期 B订单号 C商品SKU D销售数量 E销售单价 F优惠金额 G退款金额
        H销售金额(外币) I销售金额(¥) J商品成本(¥) K毛利(¥)
    """
    name = next((n for n in SALES_SHEETS if n in wb.sheetnames), None)
    if name is None:
        raise ValueError(f'模板缺少销售明细 sheet（{" / ".join(SALES_SHEETS)}）')
    ws = wb[name]
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row)

    def put(row, col, value):
        cell = ws.cell(row, col)
        cell.value = value
        cell.border = CELL_BORDER
        cell.alignment = CELL_ALIGN

    rate = result['rate']
    for i, sale in result['sales'].iterrows():
        r = i + 2
        values = [
            sale['日期'].strftime('%Y/%m/%d'), sale['Order ID'], sale['商品SKU'], sale['销售数量'],
            sale[BILL_PRICE], sale[BILL_DISCOUNT], sale[BILL_REFUND],
            f'=E{r}+F{r}+G{r}', f'=H{r}*{rate}', sale['成本'], f'=I{r}-J{r}',
        ]
        for col, value in enumerate(values, start=1):
            put(r, col, value)

    last_row = len(result['sales']) + 1
    if last_row >= 2:   # 合计行：E~K 列求和
        for col in range(5, 12):
            letter = get_column_letter(col)
            put(last_row + 1, col, f'=SUM({letter}2:{letter}{last_row})')
    return name, max(last_row, 2)


def find_month_column(ws, month):
    """在利润表「项目」表头行找「{月}月」列；没有就在右侧首个空列新增。返回列号。"""
    label = f'{month}月'
    for row in ws.iter_rows(min_row=1, max_row=5):
        if not any(c.value and '项目' in str(c.value) for c in row):
            continue
        for cell in row:
            if cell.value and str(cell.value).replace(' ', '') == label:
                return cell.column
        col = 2
        while ws.cell(row[0].row, col).value not in (None, ''):
            col += 1
        ws.cell(row[0].row, col).value = label
        return col
    raise ValueError('利润表里找不到「项目」表头行')


def write_profit_sheet(wb, result, sales_sheet, last_sales_row, ads_cny):
    ws = wb[PNL_SHEET]
    month, rate = result['month'], result['rate']
    col = find_month_column(ws, month)
    letter = get_column_letter(col)

    # 定位各项目行；记下已有的「未匹配」备注行
    rows, note_rows = {}, []
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or '').strip()
        key = next((k for k in PNL_KEYS if k in label), None)
        if key:
            rows[key] = r
        if '未匹配' in label:
            note_rows.append(r)

    def ref(key):
        return f'{letter}{rows[key]}'

    def is_blank(r):
        return ws.cell(r, col).value in (None, '')

    # 1) 明细项：直接写数值或公式
    i_range = f'{sales_sheet}!I2:I{last_sales_row}'
    g_range = f'{sales_sheet}!G2:G{last_sales_row}'
    j_range = f'{sales_sheet}!J2:J{last_sales_row}'
    cancel_cost = result['cancel_shipped_cost']
    order_count = len(result['sales']) + result['cancel_shipped_count']
    leaves = {
        # 本期销售 = 退款前的销售额（G 列退款为负数）
        '本期销售': f'=SUM({i_range})-SUM({g_range})*{rate}',
        '销售退款': f'=ABS(SUM({g_range}))*{rate}',
        '实际产品成本': f'=SUM({j_range})' + (f'+{cancel_cost}' if cancel_cost else ''),
        '发单费': f'={order_count}*{UNIT_FEE}',
        '广告费': ads_cny,
        '平台费用': result['platform_fee_cny'],
        '调整金额': 0,
    }
    for key, value in leaves.items():
        if key in rows:
            ws.cell(rows[key], col).value = value

    # 2) 团队运营费用 = 它与「本期利润」之间所有子项之和（如货损费用，名称不限，空值补 0 供人工填写）
    if '团队运营费用' in rows:
        team_row = rows['团队运营费用']
        sub_rows = [r for r in range(team_row + 1, rows.get('本期利润', team_row + 1))
                    if ws.cell(r, 1).value not in (None, '')]
        for r in sub_rows:
            if is_blank(r):
                ws.cell(r, col).value = 0
        ws.cell(team_row, col).value = '=' + '+'.join(f'{letter}{r}' for r in sub_rows) if sub_rows else 0

    # 3) 汇总项：模板已有公式就保留，空着才补
    totals = {
        '实际主营收入': ('-', ['本期销售', '销售退款']),
        '实际主营成本': ('+', ['实际产品成本', '发单费']),
        '推广费用': ('+', ['广告费']),
        'TK结算': ('+', ['平台费用', '调整金额']),
        '营业费用': ('+', ['推广费用', 'TK结算', '团队运营费用']),
        '本期利润': ('-', ['实际主营收入', '实际主营成本', '营业费用']),
    }
    for key, (op, parts) in totals.items():
        if key in rows and is_blank(rows[key]):
            ws.cell(rows[key], col).value = '=' + op.join(ref(p) for p in parts if p in rows)

    # 4) 未匹配备注：覆盖已有备注行；没有就写在「本期利润」下一行
    note = f"{month}月未匹配：取消未发货{result['cancel_unshipped_count']}，取消已发货{result['cancel_shipped_count']}"
    for r in note_rows or [rows.get('本期利润', ws.max_row) + 1]:
        ws.cell(r, 1).value = note


def write_unmatched_sheet(wb, unmatched):
    """未匹配 sheet：清掉旧内容，写入本期未匹配订单的原始行。"""
    if UNMATCHED_SHEET in wb.sheetnames:
        ws = wb[UNMATCHED_SHEET]
        ws.delete_rows(1, ws.max_row)
    elif unmatched.empty:
        return
    else:
        ws = wb.create_sheet(UNMATCHED_SHEET)
    if unmatched.empty:
        return
    ws.append(unmatched.columns.tolist())
    for row in unmatched.itertuples(index=False):
        ws.append([str(v) if isinstance(v, pd.Timestamp) else v for v in row])


# ================= 流程 =================

def copy_to_output(source_path):
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    out_path = os.path.join(OUTPUT_FOLDER, f'新_{os.path.basename(source_path)}')
    shutil.copy(source_path, out_path)
    os.chmod(out_path, stat.S_IWRITE)   # 源文件可能是只读的
    return out_path


def process_file(source_path, cost_map):
    print(f"\n>>> {os.path.basename(source_path)}")
    xls = pd.ExcelFile(source_path)
    try:
        orders, bill = load_orders_and_bill(xls)
        result = compute(orders, bill, cost_map)
    except SkipFile as reason:
        copy_to_output(source_path)
        print(f'   [跳过] {reason}，已原样复制')
        return
    ads_cny = load_ads_cny(xls)
    xls.close()

    sales = result['sales']
    print(f"   {result['month']}月 | {result['currency']} 汇率 {result['rate']} | 销售明细 {len(sales)} 单 | "
          f"未匹配：取消未发货 {result['cancel_unshipped_count']}，取消已发货 {result['cancel_shipped_count']}")
    no_cost = sales[(sales['成本'] == 0) & ((sales[BILL_PRICE] + sales[BILL_DISCOUNT] + sales[BILL_REFUND]).abs() > 0.001)]
    for _, sale in no_cost.iterrows():
        print(f"   [注意] 有收入但查不到成本：订单 {sale['Order ID']}  SKU={sale['商品SKU'] or '(空)'}")

    out_path = copy_to_output(source_path)
    wb = openpyxl.load_workbook(out_path)
    sales_sheet, last_sales_row = write_sales_sheet(wb, result)
    write_profit_sheet(wb, result, sales_sheet, last_sales_row, ads_cny)
    write_unmatched_sheet(wb, result['unmatched'])
    wb.save(out_path)
    wb.close()


def main():
    cost_path = os.path.join(INPUT_FOLDER, COST_FILE_NAME)
    if not os.path.exists(cost_path):
        print(f'找不到成本表：{cost_path}')
        return
    cost_map = load_cost_map(cost_path)
    print(f'成本表 {len(cost_map)} 个 SKU；输入 {INPUT_FOLDER}；输出 {OUTPUT_FOLDER}')

    files = [f for f in sorted(glob.glob(os.path.join(INPUT_FOLDER, '*.xlsx')))
             if os.path.basename(f) != COST_FILE_NAME and not os.path.basename(f).startswith('~$')]
    failed = []
    for path in files:
        try:
            process_file(path, cost_map)
        except Exception as e:
            failed.append(os.path.basename(path))
            print(f'   [失败] {e}')

    print(f"\n处理完成：共 {len(files)} 个文件，失败 {len(failed)} 个{'：' + '、'.join(failed) if failed else ''}")


if __name__ == '__main__':
    main()
    input('按回车键退出...')
