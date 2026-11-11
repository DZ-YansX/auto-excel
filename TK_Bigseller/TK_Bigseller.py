"""TK 大卖家（BigSeller）月度利润表生成脚本。

用法：把各店铺 xlsx 放进 TKdeal，改好配置区的本月汇率后运行，
结果写到 output/新_{原文件名}.xlsx（原文件不动）。

每个店铺文件需要的 sheet（按名称关键字识别）：
    账单原表     TikTok 导出的账单，表头可以是英文 / 中文 / 印尼语
    订单明细表   TikTok 导出的订单明细（英文表头，第 2 行是说明行）
    大卖家       BigSeller 导出的订单，含「订单号」「总成本价」
    广告费       「广告费」标签下方一格是本月广告费（美元）
    销售明细表 / 利润表 / 未匹配    模板 sheet，由脚本填写

口径：
    - 销售明细：账单里交易类型为“订单”、且订单号在订单明细表里的行，按下单日期排序。
    - 目标月份：利润表里已建好但还没填的月份列优先；否则取订单下单月份的众数。
    - 销售成本：取大卖家整单「总成本价」，每单只计一次；跨月订单（下单月≠目标月）、
      未发货订单、销售额为 0 的行不计成本。大卖家没有的订单成本记 0 并标黄，需手工补。
    - 未匹配：订单明细里有、账单里没有的订单，分为 取消未发货 / 取消已发货 / 发货未到账。
      后两类计成本和发单费：成本先取大卖家，没有再查万山成本价（查不到标黄）；跨月的成本记 0。
    - 发单费 = (销售明细行数 − 跨月行数 − 未发货单数 + 未匹配已发货单数) × 单价。
    - 调整金额：只统计目标月下单的“预扣税”行。
    - 利润表按「N月」列写入，没有该列就在最后一个月份列右侧新建，历史月份不动。
"""
import glob
import os
import posixpath
import re
import sys
import traceback
import warnings
import xml.etree.ElementTree as ET
from collections import Counter
from copy import copy
from dataclasses import dataclass, field
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Color, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ================= 配置 =================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(SCRIPT_DIR, 'TKdeal')
OUTPUT_FOLDER = os.path.join(SCRIPT_DIR, 'output')

# 万山成本价：只给“大卖家里没有的未匹配已发货订单”兜底。按顺序找，找不到也能跑。
COST_FILE_NAME = '万山成本价.xlsx'
COST_FILE_FOLDERS = (SCRIPT_DIR, os.path.join(SCRIPT_DIR, '..', '补成本'))

# 汇率：1 外币 = ? CNY。按账单「货币」列选用，每月处理前更新。
CURRENCY_RATES = {
    'PHP': 0.105,
    'MYR': 1.616,
    'THB': 0.1982,
    'IDR': 0.00037053,
}
USD_RATE = 7                                                           # 广告费按美元结算
COUNTRY_UNIT_FEES = {'IDR': 3.0, 'PHP': 3.5, 'MYR': 3.5, 'THB': 3.5}   # 发单费单价（CNY/单）

# 账单用到的列：字段 -> 各语言导出的表头（英文新版 / 英文旧版 / 中文 / 印尼语）。
# 表头对不上时脚本会直接报失败；平台改了列名就在这里补一个别名。
BILL_COL_ALIASES = {
    'oid':      ['Order/Adjustment ID', '订单ID/调整单ID', 'ID Pesanan/Penyesuaian'],
    'type':     ['Transaction type', '交易类型', 'Jenis transaksi'],
    'created':  ['Order created time', '订单创建时间', 'Waktu pemesanan'],
    'currency': ['Currency', '货币', 'Mata uang'],
    'sub':      ['Subtotal before discount', 'Subtotal before discounts',      # → 销售单价
                 '享受折扣前小计', 'Subtotal sebelum diskon'],
    'disc':     ['Seller discounts', '商家折扣', 'Diskon penjual'],             # → 优惠金额
    'refund':   ['Refund amount after discount',                               # → 退款金额
                 'Refund subtotal after seller discounts',
                 '享受商家折扣后的退款小计', 'Total pengembalian dana setelah diskon'],
    'fees':     ['Total Fees', '总费用', 'Total Biaya'],                        # → 平台费用
    'adj':      ['Adjustment amount', '调整金额', 'Jumlah penyesuaian'],        # → 调整金额
}
# 账单「交易类型」的取值（小写比较）；其余类型（物流赔付、平台赔付等）不参与计算。
ORDER_TYPES = {'order', '订单', 'pesanan'}
WITHHOLDING_TYPES = {'withholding tax', '预扣税'}

# 订单明细表（英文表头）与大卖家表（中文表头）用到的列
DETAIL_COLUMNS = {'oid': 'Order ID', 'status': 'Order Status', 'sku': 'Seller SKU',
                  'qty': 'Quantity', 'created': 'Created Time', 'shipped': 'Shipped Time',
                  'country': 'Country'}
DETAIL_LAST_COLUMN = 'Checked Status'       # 未匹配表只复制到这一列为止
BIGSELLER_COLUMNS = {'oid': '订单号', 'sku': '商品SKU', 'qty': '数量', 'cost': '总成本价'}

# 订单状态属于“在途”（小写比较）：已发货但平台还没结算 → 发货未到账
IN_TRANSIT_STATUSES = {
    'shipping', 'shipped', '已发货', '已發貨', '运输中', '運輸中', '运送中', '運送中',
    '配送中', '出货中', '出貨中', 'sedang dikirim', 'sedang dihantar',
}

# 订单明细 Country 列 -> 成本表「仓库」列
COUNTRY_WAREHOUSES = {
    'Philippines': '菲律宾', 'Indonesia': '印尼', 'Thailand': '泰国',
    'Malaysia': '马来', 'Malesia': '马来', 'ماليزيا': '马来',
}
# 这些前缀的 SKU 在成本表里查不到时，按「<前缀>-KP」的成本算（同一产品）
KP_FALLBACK_PREFIXES = ('YY', 'XY')

# sheet 名称关键字
SHEET_KEYWORDS = {'sales': '销售明细', 'bill': '账单', 'detail': '订单明细', 'ads': '广告',
                  'unmatched': '未匹配', 'bigseller': '大卖家', 'pnl': '利润'}

# 未匹配表订单号列的列宽（按站点模板），其余列用默认宽度
UNMATCHED_ID_WIDTHS = {'IDR': 20.375, 'MYR': 16.75}
UNMATCHED_ID_WIDTH_DEFAULT = 18.0

NOT_SHIPPED, CANCELLED_SHIPPED, IN_TRANSIT = '取消未发货', '取消已发货', '发货未到账'
YELLOW_FILL = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')


# ================= 小工具 =================

def to_float(value):
    """单元格 → 数字；空值、'--' 等非数字一律当 0。"""
    try:
        number = float(str(value).replace(',', '').strip())
    except ValueError:
        return 0.0
    return 0.0 if number != number else number      # NaN → 0


def tidy(number):
    """整数值写成 int，避免单元格里出现 3.0。"""
    return int(number) if float(number).is_integer() else number


def fmt_num(number):
    """数字 → 公式里的文本：最多 4 位小数，去掉尾零。"""
    return f'{round(float(number), 4):f}'.rstrip('0').rstrip('.')


def month_of(value):
    """从日期取月份。支持 datetime、YYYY/MM/DD、DD/MM/YYYY，
    以及菲律宾订单明细的美式 MM/DD/YYYY h:mm:ss AM/PM。认不出返回 None。"""
    if hasattr(value, 'month'):
        return value.month
    text = str(value or '').strip()
    match = re.match(r'\d{4}[/-](\d{1,2})', text)
    if match:
        return int(match.group(1))
    match = re.match(r'(\d{1,2})[/-](\d{1,2})[/-]\d{4}', text)
    if not match:
        return None
    first, second = int(match.group(1)), int(match.group(2))
    us_style = second > 12 or re.search(r'\b[AP]M\b', text, re.I)
    return first if us_style else second


def header_columns(ws):
    """第 1 行表头 → {表头(小写): 列号}，重名取靠左的一列。"""
    columns = {}
    for cell in ws[1]:
        if cell.value is not None:
            columns.setdefault(str(cell.value).strip().lower(), cell.column)
    return columns


def find_column(columns, names, sheet_label):
    """按别名顺序找列号；都找不到说明表头变了，直接报错而不是悄悄读成 0。"""
    for name in names:
        if name.lower() in columns:
            return columns[name.lower()]
    raise ValueError(f'{sheet_label}找不到列 {names}，表头可能变了，请在配置区补充别名')


# ================= 读取 =================

@dataclass
class BillRow:
    row: int            # 账单里的行号，用于同日期订单保持原顺序
    oid: str
    kind: str           # 交易类型（小写）
    created: object
    sub: float
    disc: float
    refund: float
    fees: float
    adj: float


def read_bill(ws):
    """账单原表 → (所有行, 币种)。"""
    columns = header_columns(ws)
    col = {key: find_column(columns, names, '账单原表') for key, names in BILL_COL_ALIASES.items()}
    rows, currency = [], None
    for r in range(2, ws.max_row + 1):
        oid = ws.cell(r, col['oid']).value
        if oid is None:
            continue
        currency = currency or str(ws.cell(r, col['currency']).value or '').strip().upper()
        rows.append(BillRow(
            row=r, oid=str(oid).strip(),
            kind=str(ws.cell(r, col['type']).value or '').strip().lower(),
            created=ws.cell(r, col['created']).value,
            **{key: to_float(ws.cell(r, col[key]).value)
               for key in ('sub', 'disc', 'refund', 'fees', 'adj')}))
    return rows, currency


@dataclass
class BigSellerOrder:
    sku: object
    qty: object
    cost: float         # 整单总成本价（CNY）


def read_bigseller(ws):
    """大卖家 sheet → {订单号: BigSellerOrder}。同一订单拆成多行时成本相加。"""
    columns = header_columns(ws)
    col = {key: find_column(columns, [name], '大卖家表') for key, name in BIGSELLER_COLUMNS.items()}
    orders = {}
    for r in range(2, ws.max_row + 1):
        oid = ws.cell(r, col['oid']).value
        if oid is None:
            continue
        oid = str(oid).strip()
        cost = to_float(ws.cell(r, col['cost']).value)
        if oid in orders:
            orders[oid].cost += cost
        else:
            orders[oid] = BigSellerOrder(ws.cell(r, col['sku']).value,
                                         ws.cell(r, col['qty']).value, cost)
    return orders


@dataclass
class DetailRow:
    row: int            # 订单明细表里的行号，写未匹配表时按它复制整行
    oid: str
    sku: str
    qty: float
    status: str
    created: object
    shipped: bool


@dataclass
class DetailOrder:
    """订单明细按订单号汇总。"""
    shipped: bool = False                         # 任一行有发货时间
    statuses: set = field(default_factory=set)    # 小写
    lines: list = field(default_factory=list)     # [(SKU, 数量), ...]


def read_details(ws):
    """订单明细表 → (数据行, {订单号: DetailOrder}, 国家)。数据从第 3 行开始。"""
    columns = header_columns(ws)
    col = {key: find_column(columns, [name], '订单明细表') for key, name in DETAIL_COLUMNS.items()}
    rows, orders, country = [], {}, None
    for r in range(3, ws.max_row + 1):
        oid = ws.cell(r, col['oid']).value
        if oid is None:
            continue
        row = DetailRow(
            row=r, oid=str(oid).strip(),
            sku=str(ws.cell(r, col['sku']).value or '').strip(),
            qty=to_float(ws.cell(r, col['qty']).value),
            status=str(ws.cell(r, col['status']).value or '').strip(),
            created=ws.cell(r, col['created']).value,
            shipped=bool(str(ws.cell(r, col['shipped']).value or '').strip()))
        rows.append(row)
        order = orders.setdefault(row.oid, DetailOrder())
        order.shipped = order.shipped or row.shipped
        order.statuses.add(row.status.lower())
        order.lines.append((row.sku, row.qty))
        country = country or str(ws.cell(r, col['country']).value or '').strip()
    return rows, orders, country


def read_ads_usd(ws):
    """广告费 sheet：「广告费」标签下方一格的美元金额。"""
    for row in ws.iter_rows():
        for cell in row:
            if str(cell.value or '').strip() == '广告费':
                return to_float(ws.cell(cell.row + 1, cell.column).value)
    print('   [警告] 广告费 sheet 里找不到「广告费」标签，广告费按 0 计')
    return 0.0


# ================= 成本表兜底 =================

def load_cost_table(path):
    """万山成本价（A:SKU  B:仓库  C:成本）→ {(SKU, 仓库): 单位成本}。"""
    wb = openpyxl.load_workbook(path, read_only=True)
    cost_map = {}
    for sku, warehouse, price in wb.worksheets[0].iter_rows(min_row=2, max_col=3, values_only=True):
        if sku is not None and warehouse is not None:
            cost_map[(str(sku).strip(), str(warehouse).strip())] = to_float(price)
    wb.close()
    return cost_map


def sku_unit_cost(sku, warehouse, cost_map):
    """单个 SKU 的单位成本，查不到返回 None。依次尝试：
    整个名字（有些正式 SKU 名里带 *，如 Coffee-*2）→ 「基础SKU*N」按基础成本×N → KP 兜底。"""
    if (sku, warehouse) in cost_map:
        return cost_map[(sku, warehouse)]
    base, times = sku, 1
    head, star, tail = sku.rpartition('*')
    if star and tail.strip().isdigit():
        base, times = head.strip(), int(tail)
        if (base, warehouse) in cost_map:
            return cost_map[(base, warehouse)] * times
    prefix = re.match(r'([A-Za-z]+)-', base)
    if prefix and prefix.group(1).upper() in KP_FALLBACK_PREFIXES:
        kp_cost = cost_map.get((prefix.group(1).upper() + '-KP', warehouse))
        if kp_cost is not None:
            return kp_cost * times
    return None


def order_cost_from_table(order, warehouse, cost_map):
    """按订单明细行算成本表成本 = Σ 单位成本 × 数量。返回 (成本, 查不到的 SKU 列表)。
    一个单元格里可能有多个 SKU（换行 / 逗号 / 分号 / 顿号分隔）；
    不按空格、+、* 拆，因为 SKU 名本身会含这些字符。"""
    total, missing = 0.0, []
    for sku_cell, qty in order.lines:
        for sku in filter(None, (s.strip() for s in re.split(r'[\r\n,，;；、]+', sku_cell))):
            unit = sku_unit_cost(sku, warehouse, cost_map)
            if unit is None:
                missing.append(sku)
            else:
                total += unit * qty
    return round(total, 6), missing


# ================= 计算 =================

@dataclass
class SalesRow:
    date: object
    oid: str
    sku: object
    qty: object
    sales: float
    disc: float
    refund: float
    cost: float
    cost_missing: bool      # 该计成本但大卖家没有这单 → 标黄
    not_shipped: bool
    cross_month: bool


@dataclass
class UnmatchedRow:
    detail: DetailRow
    label: str
    cost: object = None     # None = 不写成本（未发货，或同一订单的后续行）
    cost_missing: bool = False


@dataclass
class Report:
    month: int
    sales: list
    unmatched: list
    net_sales: float        # Σ(销售单价 + 优惠金额 + 退款金额)，原币
    refund: float           # Σ 退款金额，原币
    fees: float             # Σ 总费用，原币
    withholding: float      # 目标月预扣税，原币
    product_cost: float     # 销售成本 + 未匹配已发货成本，CNY
    cross_month_rows: int
    not_shipped_orders: int
    shipped_unmatched_orders: int
    label_counts: dict      # 未匹配各分类的订单数


def drop_last_newline(value):
    """大卖家导出值以 \\r\\n 结尾，去掉最后的 \\n（与历史表保持一致）。"""
    return value[:-1] if isinstance(value, str) and value.endswith('\n') else value


def clean_quantity(value):
    """单个数量写成数字；一单多个商品时保留大卖家的多行文本。"""
    value = drop_last_newline(value)
    try:
        return tidy(float(str(value).strip()))
    except ValueError:
        return value if isinstance(value, str) else 0


def sku_qty_from_details(order):
    """大卖家没有这单时，SKU / 数量从订单明细补：多行则 SKU 换行合并、数量相加。"""
    skus = [sku for sku, _ in order.lines if sku]
    return '\n'.join(skus) or None, sum(qty for _, qty in order.lines)


def unmatched_label(order):
    if order.statuses & IN_TRANSIT_STATUSES:
        return IN_TRANSIT
    return CANCELLED_SHIPPED if order.shipped else NOT_SHIPPED


def build_report(bill, detail_rows, detail_orders, bigseller, pending, warehouse, cost_map):
    bill_orders = [b for b in bill if b.kind in ORDER_TYPES]
    matched = [b for b in bill_orders if b.oid in detail_orders]
    if not matched:
        raise ValueError('账单与订单明细没有任何相同的订单号')

    months = Counter(m for m in (month_of(b.created) for b in matched) if m)
    if not months:
        raise ValueError('账单的下单日期认不出月份')
    month = pending or months.most_common(1)[0][0]
    print(f'   目标月份: {month}月（订单下单月份分布 {dict(months)}）')

    # ---- 销售明细：按下单日期排序，同一天保持账单原顺序 ----
    sales, costed = [], set()
    for b in sorted(matched, key=lambda b: (str(b.created or ''), b.row)):
        order = bigseller.get(b.oid)
        cross_month = month_of(b.created) not in (None, month)
        shipped = detail_orders[b.oid].shipped
        bears_cost = b.sub != 0 and not cross_month and shipped
        cost = 0
        if order is None:
            sku, qty = sku_qty_from_details(detail_orders[b.oid])
            if bears_cost:
                print(f'   [警告] 销售订单 {b.oid}（{sku}）不在大卖家表，成本填 0 并标黄')
        else:
            sku, qty = order.sku, order.qty
            if bears_cost and b.oid not in costed:      # 同一订单在账单里有多行时只计一次
                cost = order.cost
                costed.add(b.oid)
        sales.append(SalesRow(
            date=b.created, oid=b.oid, sku=drop_last_newline(sku), qty=clean_quantity(qty),
            sales=tidy(b.sub), disc=tidy(b.disc), refund=tidy(b.refund), cost=tidy(cost),
            cost_missing=order is None and bears_cost,
            not_shipped=not shipped, cross_month=cross_month))

    # ---- 未匹配：订单明细里有、账单订单行里没有 ----
    billed = {b.oid for b in bill_orders}
    unmatched, costed, unmatched_cost = [], set(), 0.0
    label_orders = {NOT_SHIPPED: set(), CANCELLED_SHIPPED: set(), IN_TRANSIT: set()}
    for d in detail_rows:
        if d.oid in billed:
            continue
        item = UnmatchedRow(d, unmatched_label(detail_orders[d.oid]))
        unmatched.append(item)
        label_orders[item.label].add(d.oid)
        if item.label == NOT_SHIPPED or d.oid in costed:
            continue
        costed.add(d.oid)
        if month_of(d.created) not in (None, month):    # 跨月：上月已计
            item.cost = 0
        elif d.oid in bigseller:
            item.cost = round(bigseller[d.oid].cost, 6)
        else:
            item.cost, missing = order_cost_from_table(detail_orders[d.oid], warehouse, cost_map)
            if missing:
                item.cost, item.cost_missing = 0, True
                print(f'   [警告] 未匹配订单 {d.oid}：成本表({warehouse})查不到 SKU {missing}，已标黄')
        unmatched_cost += item.cost

    report = Report(
        month=month, sales=sales, unmatched=unmatched,
        net_sales=sum(b.sub + b.disc + b.refund for b in matched),
        refund=sum(b.refund for b in matched),
        fees=sum(b.fees for b in matched),
        withholding=sum(b.adj for b in bill
                        if b.kind in WITHHOLDING_TYPES and month_of(b.created) == month),
        product_cost=sum(s.cost for s in sales) + unmatched_cost,
        cross_month_rows=sum(s.cross_month for s in sales),
        not_shipped_orders=len({s.oid for s in sales if s.not_shipped}),
        shipped_unmatched_orders=len(label_orders[CANCELLED_SHIPPED] | label_orders[IN_TRANSIT]),
        label_counts={label: len(oids) for label, oids in label_orders.items()})

    print(f'   销售明细: {len(sales)} 行（跨月 {report.cross_month_rows} 行，'
          f'未发货 {report.not_shipped_orders} 单）| 未匹配: '
          + '，'.join(f'{label} {count}' for label, count in report.label_counts.items())
          + f' | 实际产品成本: {fmt_num(report.product_cost)}')
    return report


# ================= 写表 =================

def write_sales_sheet(ws, sales, rate):
    """销售明细表：每行沿用模板第 2 行的样式，末尾加合计行。"""
    styles = [copy(ws.cell(2, c)._style) for c in range(1, 12)]
    last_used = max(ws.max_row, len(sales) + 2)
    for row in ws.iter_rows(min_row=2, max_row=last_used, max_col=max(12, ws.max_column)):
        for cell in row:
            cell.value = None

    for r, item in enumerate(sales, start=2):
        values = [item.date, item.oid, item.sku, item.qty, item.sales, item.disc, item.refund,
                  f'=E{r}+F{r}+G{r}', f'=H{r}*{rate}', item.cost, f'=I{r}-J{r}']
        for c, value in enumerate(values, start=1):
            cell = ws.cell(r, c)
            cell._style = copy(styles[c - 1])
            cell.value = value
        if item.cost_missing:
            ws.cell(r, 10).fill = YELLOW_FILL
        if item.not_shipped:
            mark = ws.cell(r, 12)
            mark.value = '未发货'
            mark.font = Font(name='宋体', size=11, bold=True, charset=134, color='FF000000')
            mark.alignment = Alignment(vertical='center')

    last = len(sales) + 1
    for c in range(5, 12):
        letter = get_column_letter(c)
        ws.cell(last + 1, c).value = f'=SUM({letter}2:{letter}{last})'


def write_unmatched_sheet(ws, ws_detail, unmatched, id_width):
    """未匹配表：从订单明细表原样复制表头两行和未匹配的行，右侧追加「分类」「成本」两列。"""
    last_col = header_columns(ws_detail).get(DETAIL_LAST_COLUMN.lower(), ws_detail.max_column)
    font = Font(name='宋体', size=11, charset=134)
    v_center = Alignment(vertical='center')

    def copy_row(src_row, dst_row):
        for c in range(1, last_col + 1):
            source, target = ws_detail.cell(src_row, c), ws.cell(dst_row, c)
            target.value = source.value
            target._style = copy(source._style)
            plain_font = copy(target.font)          # 统一成中文字符集、默认字色
            plain_font.charset = 134
            plain_font.color = None
            target.font = plain_font
        ws.row_dimensions[dst_row].height = ws_detail.row_dimensions[src_row].height or 15

    ws.delete_rows(1, ws.max_row)
    ws.column_dimensions['A'].width = id_width
    copy_row(1, 1)
    copy_row(2, 2)
    for r, item in enumerate(unmatched, start=3):
        copy_row(item.detail.row, r)
        label_cell = ws.cell(r, last_col + 1)
        label_cell.value, label_cell.font, label_cell.alignment = item.label, font, v_center
        if item.label != NOT_SHIPPED:
            cost_cell = ws.cell(r, last_col + 2)
            cost_cell.font, cost_cell.alignment = font, v_center
            if item.cost is not None:
                cost_cell.value = tidy(item.cost)
            if item.cost_missing:
                cost_cell.fill = YELLOW_FILL


def pnl_layout(ws):
    """利润表定位 → ({A 列项目名(去空格): 行号}, {月份: 列号}, 月份表头所在行)。"""
    rows = {}
    for cell in ws['A']:
        if cell.value is not None:
            rows[re.sub(r'[ \t]', '', str(cell.value))] = cell.row
    month_cols, header_row = {}, None
    for row in ws.iter_rows(min_row=1, max_row=5):
        for cell in row:
            match = re.fullmatch(r'(\d{1,2})\s*月', str(cell.value or '').strip())
            if match:
                month_cols[int(match.group(1))] = cell.column
                header_row = cell.row
    if not month_cols:
        raise ValueError('利润表前 5 行找不到「N月」表头')
    return rows, month_cols, header_row


def pnl_row(rows, keyword):
    """项目名包含关键字的第一行。"""
    for label, row in rows.items():
        if keyword in label:
            return row
    raise ValueError(f'利润表 A 列找不到「{keyword}」')


def pending_month(ws):
    """利润表里已建好月份列、但「实际产品成本」还空着的最新月份；没有返回 None。
    月初模板会预建当月列，以它为准可以避免少量跨月订单把结果写进上个月。"""
    rows, month_cols, _ = pnl_layout(ws)
    cost_row = pnl_row(rows, '实际产品成本')
    pending = [m for m, col in month_cols.items() if ws.cell(cost_row, col).value in (None, '')]
    return max(pending) if pending else None


def add_month_column(ws, rows, month_cols, header_row, month):
    """在最后一个月份列右侧新建一列：样式和公式照抄上一列，公式里的列字母右移一位。"""
    prev_col = max(month_cols.values())
    col = prev_col + 1
    prev_letter, letter = get_column_letter(prev_col), get_column_letter(col)
    for r in range(header_row, max(rows.values()) + 1):
        source, target = ws.cell(r, prev_col), ws.cell(r, col)
        target._style = copy(source._style)
        if isinstance(source.value, str) and source.value.startswith('='):
            target.value = re.sub(rf'\b{prev_letter}(\d+)', rf'{letter}\1', source.value)
    ws.cell(header_row, col).value = f'{month}月'
    ws.cell(pnl_row(rows, '团队运营费用'), col).value = 0
    return col


def pnl_title():
    """A1 标题：「利润表」宋体 12 号，第二行红色 8 号「单位：CNY」。"""
    def block(size, color, text):
        return TextBlock(InlineFont(rFont='宋体', charset=134, scheme='minor',
                                    b=True, sz=size, color=color), text)
    return CellRichText(block(12, Color(theme=1), '利润表'),
                        block(8, Color(rgb='FFFF0000'), '\n' + ' ' * 10 + '单位：CNY'))


def write_pnl_sheet(ws, report, rate, unit_fee, ads_usd):
    """利润表：填目标月份列，并在「本期利润」下一行写未匹配备注。"""
    ws['A1'] = pnl_title()
    rows, month_cols, header_row = pnl_layout(ws)
    col = month_cols.get(report.month)
    if col is None:
        col = add_month_column(ws, rows, month_cols, header_row, report.month)
        print(f'   利润表新建 {report.month}月 列（{get_column_letter(col)} 列）')

    refund_cny = abs(report.refund) * rate
    dispatch_count = (f'{len(report.sales)}-{report.cross_month_rows}'
                      f'-{report.not_shipped_orders}+{report.shipped_unmatched_orders}')
    values = {
        '本期销售': round(report.net_sales * rate + refund_cny, 10),     # 退款前的销售额
        '销售退款': round(refund_cny, 10),
        '实际产品成本': round(report.product_cost, 4),
        '发单费': f'=({dispatch_count})*{fmt_num(unit_fee)}',
        '广告费': f'={fmt_num(ads_usd)}*{fmt_num(USD_RATE)}',
        '平台费用': round(abs(report.fees) * rate, 10),
        '调整金额': round(abs(report.withholding) * rate, 10),
    }
    for keyword, value in values.items():
        ws.cell(pnl_row(rows, keyword), col).value = value

    note = ws.cell(pnl_row(rows, '本期利润') + 1, 1)
    parts = [f'{label}{count}' for label, count in report.label_counts.items() if count]
    if not parts:
        note.value = None
        return
    note.value = f'{report.month}月未匹配：' + '，'.join(parts)
    note.font = Font(name='宋体', size=10, bold=True, charset=134, color='FFFF0000')
    note.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)


# ================= 保存后的修补 =================
# openpyxl 保存时会把批注的字体、框大小重置成默认值。批注内容没变的 sheet，
# 直接把原文件里的批注部件（内容 + 外观）覆盖回去。

NS_MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NS_REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def comment_parts(zf):
    """xlsx 压缩包 → {sheet 名: {'comments': 批注内容部件, 'vmlDrawing': 批注外观部件}}"""
    def resolve(folder, target):
        if target.startswith('/'):
            return target.lstrip('/')
        return posixpath.normpath(posixpath.join(folder, target))

    targets = {rel.get('Id'): rel.get('Target')
               for rel in ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))}
    names, result = set(zf.namelist()), {}
    for sheet in ET.fromstring(zf.read('xl/workbook.xml')).iter(NS_MAIN + 'sheet'):
        sheet_part = resolve('xl', targets[sheet.get(NS_REL + 'id')])
        folder, file_name = posixpath.split(sheet_part)
        rels_part = f'{folder}/_rels/{file_name}.rels'
        if rels_part not in names:
            continue
        parts = {}
        for rel in ET.fromstring(zf.read(rels_part)):
            kind = rel.get('Type').rsplit('/', 1)[-1]
            if kind in ('comments', 'vmlDrawing'):
                parts[kind] = resolve(folder, rel.get('Target'))
        if len(parts) == 2:
            result[sheet.get('name')] = parts
    return result


def restore_comment_styles(source_path, output_path):
    def commented_cells(xml):
        return sorted(re.findall(rb'<comment [^>]*ref="([^"]+)"', xml))

    replacements = {}
    with ZipFile(source_path) as src, ZipFile(output_path) as out:
        originals = comment_parts(src)
        for sheet, parts in comment_parts(out).items():
            origin = originals.get(sheet)
            if origin and (commented_cells(src.read(origin['comments']))
                           == commented_cells(out.read(parts['comments']))):
                for kind, part in parts.items():
                    replacements[part] = src.read(origin[kind])
        if not replacements:
            return
        temp_path = output_path + '.tmp'
        with ZipFile(temp_path, 'w', ZIP_DEFLATED) as fixed:
            for info in out.infolist():
                fixed.writestr(info, replacements.get(info.filename) or out.read(info.filename))
    os.replace(temp_path, output_path)


def warn_if_images_lost(source_path, output_path):
    """广告费截图等图片正常会原样保留；万一 openpyxl 认不出某种图片格式，这里会提醒。"""
    def image_count(path):
        with ZipFile(path) as zf:
            return sum(name.startswith('xl/media/') and not name.endswith('/')
                       for name in zf.namelist())

    lost = image_count(source_path) - image_count(output_path)
    if lost > 0:
        print(f'   [警告] 有 {lost} 张图片没能保留到输出文件，请从原表复制')


# ================= 主流程 =================

def find_sheets(wb):
    sheets = {key: next((wb[name] for name in wb.sheetnames if keyword in name), None)
              for key, keyword in SHEET_KEYWORDS.items()}
    if sheets['unmatched'] is None:
        sheets['unmatched'] = wb.create_sheet('未匹配')
    missing = [SHEET_KEYWORDS[key] for key, ws in sheets.items() if ws is None]
    if missing:
        raise ValueError(f'缺少 sheet: {missing}')
    return sheets


def process_file(path, cost_map):
    name = os.path.basename(path)
    print(f"\n{'=' * 60}\n>>> {name}")
    wb = openpyxl.load_workbook(path, rich_text=True)    # rich_text: 保留利润表标题等富文本
    sheets = find_sheets(wb)

    bill, currency = read_bill(sheets['bill'])
    if currency not in CURRENCY_RATES:
        raise ValueError(f'币种 {currency!r} 没有配置汇率（CURRENCY_RATES）')
    rate = CURRENCY_RATES[currency]
    bigseller = read_bigseller(sheets['bigseller'])
    detail_rows, detail_orders, country = read_details(sheets['detail'])
    ads_usd = read_ads_usd(sheets['ads'])
    print(f'   币种 {currency}（汇率 {rate}）| 账单 {len(bill)} 行 | 订单明细 {len(detail_orders)} 单'
          f' | 大卖家 {len(bigseller)} 单 | 广告费 {fmt_num(ads_usd)} USD')

    report = build_report(bill, detail_rows, detail_orders, bigseller,
                          pending_month(sheets['pnl']),
                          COUNTRY_WAREHOUSES.get(country, country), cost_map)

    write_sales_sheet(sheets['sales'], report.sales, rate)
    write_unmatched_sheet(sheets['unmatched'], sheets['detail'], report.unmatched,
                          UNMATCHED_ID_WIDTHS.get(currency, UNMATCHED_ID_WIDTH_DEFAULT))
    write_pnl_sheet(sheets['pnl'], report, rate, COUNTRY_UNIT_FEES[currency], ads_usd)

    output_path = os.path.join(OUTPUT_FOLDER, f'新_{name}')
    try:
        wb.save(output_path)
    except PermissionError:
        raise ValueError(f'无法写入 {output_path}，请先在 Excel/WPS 里关闭它') from None
    restore_comment_styles(path, output_path)
    warn_if_images_lost(path, output_path)
    print(f'>>> [成功] 新_{name}')


def main():
    if not os.path.isdir(INPUT_FOLDER):
        print(f'找不到输入目录: {INPUT_FOLDER}')
        return
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    cost_path = next((path for path in (os.path.join(folder, COST_FILE_NAME)
                                        for folder in COST_FILE_FOLDERS)
                      if os.path.exists(path)), None)
    cost_map = load_cost_table(cost_path) if cost_path else {}
    print('>>> 万山成本价: ' + (f'{len(cost_map)} 条（{cost_path}）' if cost_path else '未提供'))

    files = [path for path in sorted(glob.glob(os.path.join(INPUT_FOLDER, '*.xlsx')))
             if not os.path.basename(path).startswith(('新_', '~$'))
             and os.path.basename(path) != COST_FILE_NAME]
    print(f'>>> 待处理 {len(files)} 个文件')

    failed = []
    for path in files:
        try:
            process_file(path, cost_map)
        except Exception as error:      # 单个文件出错不影响其它文件
            failed.append(os.path.basename(path))
            print(f'>>> [失败] {os.path.basename(path)}: {error}')
            if not isinstance(error, ValueError):
                traceback.print_exc()
    print(f'\n全部执行完毕：成功 {len(files) - len(failed)} 个，失败 {len(failed)} 个 {failed or ""}')


if __name__ == '__main__':
    main()
    input('按回车键退出...')
