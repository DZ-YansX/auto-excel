# -*- coding: utf-8 -*-
"""Shopee 万山（大卖家成本口径）月度利润表生成脚本。

用法：把各店铺 xlsx 放进 ShopeeDeal/（文件名以国家开头，如「印尼39店.xlsx」），
改好本月汇率后运行，结果写到 output_wanshan/新_{原文件名}.xlsx（原文件不动）。

每个店铺文件需要的 sheet：
    账单明细   Shopee 结算账单（印尼/马来表头在第 1 行，菲律宾/泰国在第 6 行，自动识别）
    订单明细   Shopee 订单导出（英文 / 泰文 / 印尼文表头）
    大卖家     BigSeller 订单导出，按「订单号」取「总成本价」
    广告       广告交易记录，D 列为金额
    销售明细 / 利润表 / 未匹配   模板 sheet，由脚本填写

口径：
    - 销售明细：账单里能在订单明细找到的订单，一行账单一行；按下单日期排序。
      E 销售单价、F 卖家优惠、G Shopee 补贴、H 退款，四项之和为本币销售额。
      各国哪些账单列算进 F/G、平台费用怎么算，见 BILL_RULES。
    - 成本：只取大卖家「总成本价」，同一订单只计一次；大卖家没有的订单记 0 并标黄。
    - 跨月订单（下单月 ≠ 目标月）：成本和发单费已在下单月计过，本月成本记 0，发单费扣除。
    - 未匹配：订单明细里有、账单里没有的订单，分为
        发货未到账（已发货未取消）、取消已发货、预售单（待发货且预计发货在下月之后）、取消未发货。
      前两类计成本（取大卖家）和发单费，后两类不计。
    - 发单费 = (销售订单数 - 跨月订单数 + 未匹配已发货订单数) × 各国单价。
    - 广告费 = 广告表 D 列合计取正 × 各国税率系数 × 汇率。
    - 利润表：写入「{月}月」列（模板已建好的空月份列优先，否则按订单最多的月份新建一列），
      历史月份不动；本期利润下一行写未匹配统计备注。
    - 不处理补单：补单明细有数据时只提示，需手工填写。
"""

import glob
import os
import posixpath
import re
import shutil
import stat
import sys
import warnings
import xml.etree.ElementTree as ET
from collections import Counter
from copy import copy
from datetime import datetime
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')


# ================= 配置 =================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER_PATH = r"ShopeeDeal"
OUTPUT_FOLDER_PATH = r"output_wanshan"

# 汇率：1 本币 = ? CNY，每月处理前更新。文件名以国家开头来选用。
CURRENCY_RATES = {
    '菲律宾': 0.105,
    '印尼': 0.00037053,
    '马来': 1.616,
    '泰国': 0.1982,
}
DISPATCH_FEES = {'菲律宾': 3.5, '印尼': 3.0, '马来': 3.5, '泰国': 3.5}      # 发单费 CNY/单
ADS_TAX_FACTORS = {'菲律宾': 1.0, '印尼': 1.11, '马来': 1.08, '泰国': 1.07}  # 广告费含税系数

SELLER_VOUCHERS = ('卖家券', '合资卖家券', '卖家币回扣', '合资卖家币回扣')
# 各国账单口径：
#   vouchers       哪些卖家券计入 F 卖家优惠（其余券留在平台费用里）
#   shopee_rebate  Shopee 补贴是否计入 G
#   fee_range      平台费用 = 这段列逐列相加；为 None 时 = 拨款 - E - F - G - H - 广告充值
BILL_RULES = {
    '菲律宾': dict(vouchers=SELLER_VOUCHERS, shopee_rebate=True, fee_range=None),
    '马来': dict(vouchers=SELLER_VOUCHERS, shopee_rebate=True, fee_range=None),
    '印尼': dict(vouchers=('卖家券',), shopee_rebate=True, fee_range=None),
    '泰国': dict(vouchers=('卖家券',), shopee_rebate=False,
               fee_range=('Shipping Fee Paid by Buyer', 'Transaction Fee')),
}

# 账单表头别名（不区分大小写；先全名匹配，再按前缀匹配，如「Total Released Amount (₱)」）
BILL_COLUMNS = {
    '订单号': ('Order ID', '订单编号'),
    '下单时间': ('Order Creation Date', '订单成立时间'),
    '商品价': ('Original Product Price', 'Product Price'),
    '卖家折扣': ('Seller Product Promotion', 'Your Seller Product Promotion'),
    '退款': ('Refund Amount', '退款金額', '退款金额'),
    'Shopee补贴': ('Rebate Provided by Shopee', 'Shopee回扣金额'),
    '卖家券': ('Voucher Sponsored by Seller', '卖家赞助的优惠券'),
    '合资卖家券': ('Cofund Voucher Sponsored by Seller', '卖家赞助的合资优惠券'),
    '卖家币回扣': ('Coin Cashback Sponsored by Seller', '卖家赞助的 Shopee 币回扣'),
    '合资卖家币回扣': ('Cofund Coin Cashback Sponsored by Seller', '卖家赞助的合资 Shopee 币回扣'),
    '广告充值': ('Ads Sales Top Up Fee', 'Ads Escrow Top Up Fee', '广告托管充值费用'),
    '拨款': ('Total Released Amount', '拨款金额'),
}
BILL_REQUIRED = ('订单号', '下单时间', '商品价', '退款', '拨款')

# 订单明细表头别名（英文 / 泰文 / 印尼文）
ORDER_COLUMNS = {
    '订单号': ('Order ID', 'หมายเลขคำสั่งซื้อ', 'No. Pesanan'),
    '状态': ('Order Status', 'สถานะการสั่งซื้อ', 'Status Pesanan'),
    '发货时间': ('Ship Time', 'เวลาส่งสินค้า', 'Waktu Pengiriman Diatur'),
    '下单时间': ('Order Creation Date', 'วันที่ทำการสั่งซื้อ', 'Waktu Pesanan Dibuat'),
    'SKU': ('SKU Reference No.', 'เลขอ้างอิง SKU (SKU Reference No.)', 'Nomor Referensi SKU'),
    '数量': ('Quantity', 'จำนวน', 'Jumlah'),
    '预计发货': ('Estimated Ship Out Date', 'วันที่คาดว่าจะทำการจัดส่งสินค้า'),
    # 未匹配表复制到这一列为止，后面接「未匹配类型」「成本」。每月会把上月未匹配的跨月单
    # 连同这两列粘回订单明细并标红，所以列位置要保持固定。
    '未匹配末列': ('Zip Code', 'Postal Code', 'ประเภทคำสั่งซื้อ', 'Waktu Pesanan Selesai'),
}
ORDER_REQUIRED = ('订单号', '状态', '发货时间', '下单时间', 'SKU', '数量')
CANCELLED_STATUSES = {'cancelled', 'cancellation requested', 'batal', 'ยกเลิกแล้ว'}
TO_SHIP_STATUSES = {'to ship', 'perlu dikirim', 'ที่ต้องจัดส่ง'}

SHEET_NAMES = {
    'pnl': ('利润表',),
    'sales': ('销售明细', '销售明细表'),
    'bill': ('账单明细', '收入明细表', '收入明细'),
    'order': ('订单明细', '订单明细原表', '订单明细表'),
    'ads': ('广告', '广告费'),
    'supplement': ('补单明细', '补单'),
    'unmatched': ('未匹配', '未匹配订单'),
    'dmj': ('大卖家',),
}
REQUIRED_SHEETS = ('pnl', 'sales', 'bill', 'order', 'dmj')

SHIPPED_LABELS = ('取消已发货', '发货未到账')   # 计成本、计发单费的未匹配类别
YELLOW_FILL = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')


# ================= 通用工具 =================

def to_number(value):
    if isinstance(value, bool) or value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(',', '').strip()
    try:
        return float(text)
    except ValueError:
        return 0.0


def sum_numbers_in_cell(value):
    """大卖家多商品订单的成本单元格形如 '6.58\\n6.58'，把其中数字相加；无数字返回 None。"""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    numbers = re.findall(r'-?\d+(?:\.\d+)?', str(value or '').replace(',', ''))
    return sum(float(n) for n in numbers) if numbers else None


def clean_number(value):
    value = round(float(value), 6)
    return int(value) if value.is_integer() else value


def fmt_num(value):
    value = round(float(value), 4)
    return str(int(value)) if value.is_integer() else f'{value:f}'.rstrip('0').rstrip('.')


def month_of(value):
    """从日期（datetime 或 'YYYY-MM-DD…' 文本）取月份，取不到返回 None。"""
    if isinstance(value, datetime):
        return value.month
    match = re.match(r'\s*\d{4}[/-](\d{1,2})', str(value or ''))
    return int(match.group(1)) if match else None


def normalize_header(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip().lower()


def find_columns(header_cells, aliases):
    """按别名表定位列号：{字段: 列号}。别名按先后优先，先全名后前缀；找不到的字段不出现在结果里。"""
    headers = [(cell.column, normalize_header(cell.value)) for cell in header_cells
               if cell.value not in (None, '')]
    found = {}
    for field, names in aliases.items():
        names = [normalize_header(n) for n in names]
        column = (next((c for n in names for c, h in headers if h == n), None)
                  or next((c for n in names for c, h in headers if h.startswith(n)), None))
        if column is not None:
            found[field] = column
    return found


def get_sheet(wb, names):
    for title in wb.sheetnames:
        if title.strip() in names:
            return wb[title]
    return None


def copy_cell(source, target):
    target.value = source.value
    if source.has_style:
        target._style = copy(source._style)


# ================= 读取 =================

def find_bill_header_row(ws):
    order_id_names = {normalize_header(n) for n in BILL_COLUMNS['订单号']}
    for row in ws.iter_rows(min_row=1, max_row=10):
        if any(normalize_header(cell.value) in order_id_names for cell in row):
            return row[0].row
    raise ValueError('账单明细前 10 行找不到「Order ID / 订单编号」表头')


def read_bill(ws, country):
    """读取账单，每行返回 E/F/G/H 四项本币金额和平台费用。"""
    header_row = find_bill_header_row(ws)
    header = ws[header_row]
    cols = find_columns(header, BILL_COLUMNS)
    missing = [f for f in BILL_REQUIRED if f not in cols]
    if missing:
        raise ValueError(f'账单明细缺少列: {missing}')

    rule = BILL_RULES[country]
    fee_cols = None
    if rule['fee_range']:
        start, end = (find_columns(header, {'x': (name,)}).get('x') for name in rule['fee_range'])
        if not start or not end:
            raise ValueError(f'账单明细找不到平台费用列范围: {rule["fee_range"]}')
        fee_cols = range(start, end + 1)

    def value(row, field):
        return to_number(row[cols[field] - 1].value) if field in cols else 0.0

    rows = []
    for row in ws.iter_rows(min_row=header_row + 1):
        oid = str(row[cols['订单号'] - 1].value or '').strip()
        if not oid:
            continue
        sales = value(row, '商品价')
        seller_discount = value(row, '卖家折扣') + sum(value(row, v) for v in rule['vouchers'])
        shopee_rebate = value(row, 'Shopee补贴') if rule['shopee_rebate'] else 0.0
        refund = value(row, '退款')
        if fee_cols:
            fee = sum(to_number(row[c - 1].value) for c in fee_cols)
        else:
            # 拨款 = 销售四项 + 平台费用 + 广告充值；广告充值单独计入广告费。
            fee = (value(row, '拨款') - sales - seller_discount - shopee_rebate
                   - refund - value(row, '广告充值'))
        rows.append({
            'oid': oid, 'date': row[cols['下单时间'] - 1].value,
            'sales': sales, 'disc': seller_discount, 'rebate': shopee_rebate,
            'refund': refund, 'fee': fee,
        })
    return rows


def read_orders(ws):
    """读取订单明细：返回 (按订单汇总的 dict, 未匹配表要复制到的最后一列)。"""
    cols = find_columns(ws[1], ORDER_COLUMNS)
    missing = [f for f in ORDER_REQUIRED if f not in cols]
    if missing:
        raise ValueError(f'订单明细缺少列: {missing}')
    last_col = cols.get('未匹配末列') or max(c.column for c in ws[1] if c.value not in (None, ''))

    orders = {}
    for row in ws.iter_rows(min_row=2):
        oid = str(row[cols['订单号'] - 1].value or '').strip()
        if not oid:
            continue
        order = orders.setdefault(oid, {
            'rows': [], 'skus': [], 'qty': 0.0, 'status': '', 'shipped': False,
            'created': row[cols['下单时间'] - 1].value,
            'estimated': row[cols['预计发货'] - 1].value if '预计发货' in cols else None,
        })
        order['rows'].append(row[0].row)
        sku = str(row[cols['SKU'] - 1].value or '').strip()
        if sku:
            order['skus'].append(sku)
        order['qty'] += to_number(row[cols['数量'] - 1].value)
        order['status'] = order['status'] or str(row[cols['状态'] - 1].value or '').strip()
        ship_time = str(row[cols['发货时间'] - 1].value or '').strip()
        order['shipped'] = order['shipped'] or ship_time not in ('', '-')
    return orders, last_col


def read_dmj_costs(ws):
    """大卖家：订单号 -> 总成本价。同一订单多行（如已完成+平台处理中重复导出）取最大值。"""
    if ws.max_row < 2:
        return {}
    cols = find_columns(ws[1], {'订单号': ('订单号',), '成本': ('总成本价', '商品成本')})
    if len(cols) < 2:
        raise ValueError('大卖家 sheet 缺少「订单号」或「总成本价」列')
    costs = {}
    for row in ws.iter_rows(min_row=2):
        oid = str(row[cols['订单号'] - 1].value or '').strip()
        cost = sum_numbers_in_cell(row[cols['成本'] - 1].value)
        if oid and cost is not None:
            costs[oid] = max(cost, costs.get(oid, cost))
    return costs


def read_ads_total(ws):
    """广告表 D 列（表头 Amount / 总额 等）下方数值之和，取正数。"""
    if ws is None:
        return 0.0
    names = {'amount', '总额', '總額', 'expense', '花费', '費用', 'ค่าใช้จ่าย'}
    header_row = next((r for r in range(1, min(ws.max_row, 20) + 1)
                       if normalize_header(ws.cell(r, 4).value) in names), None)
    if header_row is None:
        if any(cell.value not in (None, '') for row in ws.iter_rows() for cell in row):
            raise ValueError('广告表 D 列找不到 Amount/总额 表头，请检查广告导出格式')
        return 0.0
    values = (ws.cell(r, 4).value for r in range(header_row + 1, ws.max_row + 1))
    return abs(sum(v for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)))


def has_supplements(ws):
    if ws is None or ws.max_row < 2:
        return False
    col = find_columns(ws[1], {'oid': ('线上单号',)}).get('oid')
    return bool(col) and any(ws.cell(r, col).value not in (None, '')
                             for r in range(2, ws.max_row + 1))


# ================= 计算 =================

def find_pending_month(ws_pnl):
    """模板里已建好、但「实际产品成本」仍为空的最新月份列。"""
    cost_row = find_label_row(ws_pnl, '实际产品成本')
    header_row, month_cols = find_month_columns(ws_pnl)
    pending = [m for m, c in month_cols.items() if ws_pnl.cell(cost_row, c).value in (None, '')]
    return max(pending) if pending else None


def classify_unmatched(order, target_month):
    status = order['status'].lower()
    if order['shipped']:
        return '取消已发货' if status in CANCELLED_STATUSES else '发货未到账'
    estimated_month = month_of(order['estimated'])
    if status in TO_SHIP_STATUSES and estimated_month not in (None, target_month):
        return '预售单'
    return '取消未发货'


def compute(bill_rows, orders, dmj_costs, target_month):
    """生成销售明细行、未匹配订单和利润表需要的各项合计（本币 / 订单数）。"""
    def is_cross_month(created):
        month = month_of(created)
        return month is not None and month != target_month

    sales_rows = []
    costed = set()
    for bill in sorted((b for b in bill_rows if b['oid'] in orders),
                       key=lambda b: str(b['date'] or '')):
        oid = bill['oid']
        order = orders[oid]
        cross = is_cross_month(bill['date'])
        cost, cost_missing = 0.0, False
        if not cross and oid not in costed:
            costed.add(oid)
            if oid in dmj_costs:
                cost = dmj_costs[oid]
            else:
                cost_missing = True
        sales_rows.append(dict(bill, sku='\n'.join(order['skus']) or None,
                               qty=order['qty'], cost=cost,
                               cost_missing=cost_missing, cross=cross))

    bill_oids = {b['oid'] for b in bill_rows}
    unmatched = []
    for oid, order in orders.items():
        if oid in bill_oids:
            continue
        label = classify_unmatched(order, target_month)
        item = {'oid': oid, 'rows': order['rows'], 'label': label,
                'cross': is_cross_month(order['created']), 'cost': None, 'cost_missing': False}
        if label in SHIPPED_LABELS:
            if item['cross']:
                item['cost'] = 0.0
            elif oid in dmj_costs:
                item['cost'] = dmj_costs[oid]
            else:
                item['cost'], item['cost_missing'] = 0.0, True
        unmatched.append(item)

    shipped_unmatched = [u for u in unmatched if u['label'] in SHIPPED_LABELS]
    cross_orders = ({r['oid'] for r in sales_rows if r['cross']}
                    | {u['oid'] for u in shipped_unmatched if u['cross']})
    return {
        'sales_rows': sales_rows,
        'unmatched': unmatched,
        'sales_local': sum(r['sales'] + r['disc'] + r['rebate'] for r in sales_rows),
        'refund_local': sum(r['refund'] for r in sales_rows),
        'fee_local': abs(sum(r['fee'] for r in sales_rows)),
        'product_cost': (sum(r['cost'] for r in sales_rows)
                         + sum(u['cost'] for u in shipped_unmatched)),
        'sales_orders': len({r['oid'] for r in sales_rows}),
        'cross_orders': len(cross_orders),
        'shipped_unmatched_orders': len(shipped_unmatched),
        'label_counts': Counter(u['label'] for u in unmatched),
        'missing_cost_orders': sorted({r['oid'] for r in sales_rows if r['cost_missing']}
                                      | {u['oid'] for u in unmatched if u['cost_missing']}),
    }


# ================= 写表 =================

def write_sales_sheet(ws, sales_rows, rate):
    """A 日期、B 订单号、C SKU、D 数量、E~H 四项本币金额、I 本币合计、J 人民币、K 成本、L 毛利。"""
    styles = [copy(ws.cell(2, c)._style) for c in range(1, 13)]
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    for r, item in enumerate(sales_rows, start=2):
        values = [item['date'], item['oid'], item['sku'], clean_number(item['qty']),
                  clean_number(item['sales']), clean_number(item['disc']),
                  clean_number(item['rebate']), clean_number(item['refund']),
                  f'=E{r}+F{r}+G{r}+H{r}', f'=I{r}*{rate}',
                  clean_number(item['cost']), f'=J{r}-K{r}']
        for c, value in enumerate(values, start=1):
            cell = ws.cell(r, c, value)
            cell._style = copy(styles[c - 1])
        if item['cost_missing']:
            ws.cell(r, 11).fill = YELLOW_FILL

    last_row = len(sales_rows) + 1
    for c in range(4, 13):
        letter = get_column_letter(c)
        cell = ws.cell(last_row + 1, c, f'=SUM({letter}2:{letter}{last_row})')
        cell.font = Font(name='宋体', size=11, bold=True, charset=134)


def write_unmatched_sheet(ws, ws_order, unmatched, last_col):
    """复制未匹配订单的订单明细原行，后面加「未匹配类型」「成本(¥)」两列。"""
    if ws.max_row:
        ws.delete_rows(1, ws.max_row)
    label_col, cost_col = last_col + 1, last_col + 2
    for c in range(1, last_col + 1):
        copy_cell(ws_order.cell(1, c), ws.cell(1, c))
        letter = get_column_letter(c)
        ws.column_dimensions[letter].width = ws_order.column_dimensions[letter].width
    for c, title in ((label_col, '未匹配类型'), (cost_col, '成本(¥)')):
        ws.cell(1, c, title)._style = copy(ws_order.cell(1, last_col)._style)
        ws.column_dimensions[get_column_letter(c)].width = 12

    out_row = 2
    for item in unmatched:
        for i, source_row in enumerate(item['rows']):
            for c in range(1, last_col + 1):
                copy_cell(ws_order.cell(source_row, c), ws.cell(out_row, c))
            ws.cell(out_row, label_col, item['label'])
            if i == 0 and item['cost'] is not None:
                cost_cell = ws.cell(out_row, cost_col, clean_number(item['cost']))
                if item['cost_missing']:
                    cost_cell.fill = YELLOW_FILL
            out_row += 1


def find_label_row(ws, key):
    for cell in ws['A']:
        if key in re.sub(r'\s', '', str(cell.value or '')):
            return cell.row
    raise ValueError(f'利润表找不到行: {key}')


def find_month_columns(ws):
    """返回 (表头行, {月份: 列号})，表头在前 5 行内形如「10月」。"""
    for row in ws.iter_rows(min_row=1, max_row=5):
        months = {int(m.group(1)): cell.column for cell in row
                  if (m := re.fullmatch(r'(\d{1,2})\s*月', str(cell.value or '').strip()))}
        if months:
            return row[0].row, months
    raise ValueError('利润表前 5 行找不到「N月」表头')


def month_column(ws, month):
    """取目标月份列；没有就在最后一个月份列右侧新建一列（复制样式和公式）。"""
    header_row, month_cols = find_month_columns(ws)
    if month in month_cols:
        return month_cols[month]
    prev_col = max(month_cols.values())
    col = prev_col + 1
    prev_letter, letter = get_column_letter(prev_col), get_column_letter(col)
    for r in range(header_row, ws.max_row + 1):
        source, target = ws.cell(r, prev_col), ws.cell(r, col)
        target._style = copy(source._style)
        if isinstance(source.value, str) and source.value.startswith('='):
            target.value = re.sub(rf'\b{prev_letter}(\d+)', rf'{letter}\1', source.value)
    ws.cell(header_row, col, f'{month}月')
    return col


def write_pnl_sheet(ws, month, values, note):
    col = month_column(ws, month)
    letter = get_column_letter(col)

    def ref(key):
        return f'{letter}{find_label_row(ws, key)}'

    for key, value in values.items():
        ws[ref(key)] = value if isinstance(value, str) else clean_number(value)

    settlement_key = next(re.sub(r'\s', '', str(c.value)).lstrip('0123456789.、')
                          for c in ws['A'] if str(c.value or '').strip().endswith('结算'))
    formulas = {
        '实际主营收入': f'={ref("本期销售")}-{ref("补单金额")}-{ref("销售退款")}',
        '实际主营成本': f'={ref("实际产品成本")}+{ref("发单费")}',
        '推广费用': f'={ref("广告费")}+{ref("补单费")}',
        settlement_key: f'={ref("平台费用")}',
        '营业费用': f'={ref("推广费用")}+{ref(settlement_key)}',
        '本期利润': f'={ref("实际主营收入")}-{ref("实际主营成本")}-{ref("营业费用")}',
    }
    for key, formula in formulas.items():
        ws[ref(key)] = formula

    ws['A1'] = CellRichText([
        TextBlock(InlineFont(rFont='宋体', sz=12, b=True, charset=134,
                             color=Color(rgb='FF000000')), '利润表'),
        TextBlock(InlineFont(rFont='宋体', sz=8, b=True, charset=134,
                             color=Color(rgb='FFFF0000')), '\n单位：CNY'),
    ])
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    note_cell = ws.cell(find_label_row(ws, '本期利润') + 1, 1, note)
    note_cell.font = Font(name='宋体', size=10, bold=True, charset=134, color='FFFF0000')
    note_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)


# ================= 保存后修补 =================
# openpyxl 保存时会重建利润表批注，Excel 打开会提示修复，所以保存后从原文件把批注原样拷回。

PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
DOC_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
# 关系类型后缀 -> 工作表里引用它的节点名（批注本身没有节点，挂在 vmlDrawing 上显示）
SHEET_REL_NODES = {'/vmlDrawing': 'legacyDrawing', '/comments': None}
# 工作表 XML 里排在 legacyDrawing 之后的节点，新节点插在它们前面
LATER_NODES = (b'<legacyDrawingHF', b'<picture', b'<oleObjects', b'<controls',
               b'<webPublishItems', b'<tableParts', b'<extLst', b'</worksheet>')


def rels_path(part):
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')


def resolve_target(part, target):
    if target.startswith('/'):
        return target.lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def worksheet_part(read, sheet_name):
    workbook = ET.fromstring(read('xl/workbook.xml'))
    targets = {r.attrib['Id']: r.attrib['Target']
               for r in ET.fromstring(read('xl/_rels/workbook.xml.rels'))}
    for sheet in workbook.iter(f'{{{MAIN_NS}}}sheet'):
        if sheet.attrib['name'] == sheet_name:
            return resolve_target('xl/workbook.xml', targets[sheet.attrib[f'{{{DOC_REL_NS}}}id']])
    raise KeyError(sheet_name)


def xml_bytes(element, default_ns):
    ET.register_namespace('', default_ns)
    return ET.tostring(element, encoding='utf-8', xml_declaration=True)


def copy_sheet_relations(src, parts, sheet_name, suffixes):
    """把原文件中某 sheet 指定类型的关系（及其引用的文件）复制到输出包 parts。"""
    src_sheet = worksheet_part(src.read, sheet_name)
    if rels_path(src_sheet) not in src.namelist():
        return 0
    src_rels = [r for r in ET.fromstring(src.read(rels_path(src_sheet)))
                if r.attrib['Type'].endswith(suffixes)]
    if not src_rels:
        return 0

    out_sheet = worksheet_part(lambda name: parts[name], sheet_name)
    out_rels_path = rels_path(out_sheet)
    out_rels = (ET.fromstring(parts[out_rels_path]) if out_rels_path in parts
                else ET.Element(f'{{{PKG_REL_NS}}}Relationships'))
    types = ET.fromstring(parts['[Content_Types].xml'])

    # 先删掉 openpyxl 写出的同类关系和文件
    for rel in [r for r in out_rels if r.attrib['Type'].endswith(suffixes)]:
        removed = resolve_target(out_sheet, rel.attrib['Target'])
        parts.pop(removed, None)
        for node in [n for n in types if n.attrib.get('PartName') == '/' + removed]:
            types.remove(node)
        out_rels.remove(rel)

    sheet_xml = parts[out_sheet]
    for node in SHEET_REL_NODES.values():
        if node:
            sheet_xml = re.sub(rb'<%s\s[^>]*/>' % node.encode(), b'', sheet_xml)

    used_ids = {r.attrib['Id'] for r in out_rels}
    copied, new_nodes = set(), []
    for rel in src_rels:
        rel_id = next(f'rId{n}' for n in range(1, 1000) if f'rId{n}' not in used_ids)
        used_ids.add(rel_id)
        ET.SubElement(out_rels, f'{{{PKG_REL_NS}}}Relationship', dict(rel.attrib, Id=rel_id))
        part = resolve_target(src_sheet, rel.attrib['Target'])
        copied.add(part)
        if rels_path(part) in src.namelist():          # vml 引用的图片等
            copied.add(rels_path(part))
            for sub in ET.fromstring(src.read(rels_path(part))):
                copied.add(resolve_target(part, sub.attrib['Target']))
        node = next(n for s, n in SHEET_REL_NODES.items() if rel.attrib['Type'].endswith(s))
        if node:
            new_nodes.append(f'<{node} r:id="{rel_id}"/>'.encode())
    for part in copied:
        if part in src.namelist():
            parts[part] = src.read(part)

    if new_nodes:
        if b'xmlns:r=' not in sheet_xml[:2000]:
            sheet_xml = sheet_xml.replace(b'<worksheet ', b'<worksheet xmlns:r="%s" '
                                          % DOC_REL_NS.encode(), 1)
        cut = min(i for i in (sheet_xml.find(tag) for tag in LATER_NODES) if i >= 0)
        sheet_xml = sheet_xml[:cut] + b''.join(new_nodes) + sheet_xml[cut:]
    parts[out_sheet] = sheet_xml
    parts[out_rels_path] = xml_bytes(out_rels, PKG_REL_NS)

    # 补上 Content Types 声明
    src_types = ET.fromstring(src.read('[Content_Types].xml'))
    have = {n.attrib.get('Extension') or n.attrib.get('PartName') for n in types}
    extensions = {posixpath.splitext(p)[1].lstrip('.').lower() for p in copied}
    for node in src_types:
        key = node.attrib.get('Extension') or node.attrib.get('PartName')
        wanted = (key.lower() in extensions if 'Extension' in node.attrib
                  else key.lstrip('/') in copied)
        if wanted and key not in have:
            types.append(node)
            have.add(key)
    parts['[Content_Types].xml'] = xml_bytes(types, CONTENT_TYPES_NS)
    return len(src_rels)


def patch_saved_workbook(source_path, output_path, pnl_name):
    with ZipFile(output_path) as out:
        infos = out.infolist()
        parts = {info.filename: out.read(info.filename) for info in infos}
    with ZipFile(source_path) as src:
        copy_sheet_relations(src, parts, pnl_name, tuple(SHEET_REL_NODES))
    # 纯空格的富文本片段必须带 xml:space，否则 Excel 会提示修复
    for name in parts:
        if name.startswith('xl/worksheets/') and name.endswith('.xml'):
            parts[name] = re.sub(rb'<t>([ \t\r\n]+)</t>',
                                 rb'<t xml:space="preserve">\1</t>', parts[name])

    temp_path = output_path + '.tmp'
    with ZipFile(temp_path, 'w', ZIP_DEFLATED) as dst:
        written = set()
        for info in infos:
            if info.filename in parts:
                dst.writestr(info, parts[info.filename])
                written.add(info.filename)
        for name in parts.keys() - written:
            dst.writestr(name, parts[name])
    os.replace(temp_path, output_path)


# ================= 流程 =================

def detect_country(file_name):
    return next((c for c in CURRENCY_RATES if file_name.startswith(c)), None)


def prepare_output(file_path, out_dir):
    """复制原文件到输出目录；同名输出正被 Excel 打开时改名为 …_修正版N。"""
    stem, ext = os.path.splitext(os.path.basename(file_path))
    output_path = os.path.join(out_dir, f'新_{stem}{ext}')
    suffix = 1
    while os.path.exists(output_path):
        try:
            os.chmod(output_path, stat.S_IWRITE)
            os.remove(output_path)
        except PermissionError:
            output_path = os.path.join(out_dir, f'新_{stem}_修正版{suffix}{ext}')
            suffix += 1
    shutil.copy2(file_path, output_path)
    os.chmod(output_path, stat.S_IWRITE)
    return output_path


def process_file(file_path, out_dir):
    file_name = os.path.basename(file_path)
    country = detect_country(file_name)
    if country is None:
        print(f'[跳过] 文件名无法识别国家: {file_name}')
        return
    rate = CURRENCY_RATES[country]
    print(f"\n{'=' * 60}\n>>> {file_name} | {country} 汇率 {rate}")

    output_path = prepare_output(file_path, out_dir)
    wb = openpyxl.load_workbook(output_path, rich_text=True)
    try:
        sheets = {key: get_sheet(wb, names) for key, names in SHEET_NAMES.items()}
        missing = [SHEET_NAMES[k][0] for k in REQUIRED_SHEETS if sheets[k] is None]
        if missing:
            raise ValueError(f'缺少 sheet: {missing}')
        if sheets['unmatched'] is None:
            sheets['unmatched'] = wb.create_sheet('未匹配')

        bill_rows = read_bill(sheets['bill'], country)
        orders, last_order_col = read_orders(sheets['order'])
        dmj_costs = read_dmj_costs(sheets['dmj'])
        matched_months = Counter(month_of(b['date']) for b in bill_rows if b['oid'] in orders)
        matched_months.pop(None, None)
        if not matched_months:
            raise ValueError('账单与订单明细没有任何匹配订单')
        month = find_pending_month(sheets['pnl']) or matched_months.most_common(1)[0][0]

        result = compute(bill_rows, orders, dmj_costs, month)
        write_sales_sheet(sheets['sales'], result['sales_rows'], rate)
        write_unmatched_sheet(sheets['unmatched'], sheets['order'], result['unmatched'],
                              last_order_col)

        dispatch = str(result['sales_orders'])
        if result['cross_orders']:
            dispatch += f"-{result['cross_orders']}"
        if result['shipped_unmatched_orders']:
            dispatch += f"+{result['shipped_unmatched_orders']}"
        counts = result['label_counts']
        refund = result['refund_local']
        note = f"{month}月未匹配：取消未发货{counts['取消未发货']}，取消已发货{counts['取消已发货']}"
        for label in ('发货未到账', '预售单'):
            if counts[label]:
                note += f'，{label}{counts[label]}'
        values = {
            '本期销售': (result['sales_local'] + refund + abs(refund)) * rate,
            '补单金额': 0,
            '销售退款': abs(refund) * rate,
            '实际产品成本': result['product_cost'],
            '发单费': f'=({dispatch})*{fmt_num(DISPATCH_FEES[country])}',
            '广告费': read_ads_total(sheets['ads']) * ADS_TAX_FACTORS[country] * rate,
            '补单费': 0,
            '平台费用': result['fee_local'] * rate,
        }
        write_pnl_sheet(sheets['pnl'], month, values, note)

        supplements = has_supplements(sheets['supplement'])
        wb.calculation.fullCalcOnLoad = True
        wb.save(output_path)
        wb.close()
        patch_saved_workbook(file_path, output_path, sheets['pnl'].title)
    except Exception:
        wb.close()
        os.remove(output_path)
        raise

    print(f"   {month}月 | 销售 {result['sales_orders']} 单"
          f"（跨月 {result['cross_orders']}）| 未匹配 {dict(counts)}")
    print(f"   本期销售 {values['本期销售']:.2f} | 产品成本 {values['实际产品成本']:.2f} | "
          f"发单费 {values['发单费']} | 广告费 {values['广告费']:.2f} | "
          f"平台费用 {values['平台费用']:.2f}")
    if result['missing_cost_orders']:
        print(f"   [注意] 大卖家没有成本、已记 0 标黄的订单: {result['missing_cost_orders']}")
    if supplements:
        print('   [注意] 补单明细有数据，本脚本不计算补单，请手工填写补单金额/补单费')
    print(f'>>> 完成: {output_path}')


def resolve_path(path):
    return path if os.path.isabs(path) else os.path.join(SCRIPT_DIR, path)


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else resolve_path(INPUT_FOLDER_PATH)
    if not os.path.isdir(folder):
        print(f'找不到输入文件夹: {folder}')
        return 1
    out_dir = resolve_path(OUTPUT_FOLDER_PATH)
    os.makedirs(out_dir, exist_ok=True)

    files = [p for p in sorted(glob.glob(os.path.join(folder, '*.xlsx')))
             if not os.path.basename(p).startswith(('新_', '~$'))]
    print(f'找到 {len(files)} 个店铺文件')
    failures = 0
    for path in files:
        try:
            process_file(path, out_dir)
        except Exception as exc:
            failures += 1
            print(f'[失败] {os.path.basename(path)}: {exc}')
    print(f'\n全部执行完毕：成功 {len(files) - failures}，失败 {failures}')
    return 1 if failures else 0


if __name__ == '__main__':
    raise SystemExit(main())
