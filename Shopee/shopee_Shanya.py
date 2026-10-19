# -*- coding: utf-8 -*-
"""
Shopee 利润表自动化批量填写程序（AURA SKIN 店铺 / 泰国站）

用法：把待处理的「原表」xlsx 放进脚本同目录下的 Shopee_Data_Shanya 文件夹，
并放入唯一成本表 山亚产品成本.xlsx，运行本脚本，
即可批量输出「新_xxx.xlsx」到 output_Shanya 文件夹。

销售明细、利润表和未匹配统一使用山亚产品成本.xlsx；输出文件中
不保留“销售明细1”sheet。

本版本相对旧版仅重构【填写逻辑】，源表读取/复制/格式等机制保持不变。
核心计算规则（已与“正确表”逐项核对一致）：
  销售明细：以收入明细表为主顺序遍历，保留能在订单明细原表匹配到的订单（378 单）
     A 日期      = 收入表 E (Order Creation Date)
     C 商品SKU   = 订单表 T 列该订单所有行 SKU，用 "," 连接
     D 销售数量  = 订单表 X 列该订单所有行求和
     E 销售单价  = 收入表 L (Original product price) 该订单求和
     F 优惠金额  = 收入表 M (Your Seller product promotion) 该订单求和
     G 优惠劵金额= 收入表 P (Voucher Sponsored by Seller) 该订单求和
     H 退款金额  = 收入表 N (Refund Amount) 该订单求和
     I/J/L 为表内公式；K 商品成本总额 = Σ(unit_cost(SKU)*数量)
  利润表（单位 ¥，汇率 D2=D5=EXCHANGE_RATE）：
     B4 本期销售     = SUM(销售明细 J) + B6
     B5 其中:补单金额= 补单订单在销售明细中 J 列求和
     B6 销售退款     = -SUM(销售明细 H) * 汇率
     B8 实际产品成本 = SUM(销售明细 K) + 未匹配已发货订单成本
     B9 发单费       = (销售明细订单数 + 未匹配已发货数 - 补单数) * UNIT_FEE
     跨月订单（下单月≠本期月）：成本记 0、不计入发单费（上月已计）
     B12 广告费      = ABS(SUM(广告费 Amount 列)) * 汇率
     B13 补单佣金    = Σ 补单明细 佣金(P)
     B14 补单平台费用= 补单订单在收入表 T:AJ 整段求和（改正数=取正）* 汇率
     B15 补单汇率差额= Σ补单垫付小计(O) - Σ补单垫付成本(K) * 汇率（文本转数值后）
     B16 补单空包费  = 补单数 * UNIT_FEE
     B18 平台费用    = 销售明细订单在收入表 T:AJ 整段求和（改正数=取正）* 汇率 - B14
  未匹配：订单明细原表中订单号在收入表查不到的订单，整行连表头连格式复制，
         末尾加 BH 成本价(¥) = 该行 unit_cost(SKU)*数量；
         未匹配已发货 = Q 列(发货时间)非空的订单。
"""

import os
import re
import shutil
import time
import glob
import stat
import warnings
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import Border, Side, Alignment
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

# ================= 配置区域 =================

# 1. 待处理文件所在的文件夹（脚本同目录下的 Shopee_Data_Shanya）
TARGET_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Shopee_Data_Shanya')

# 1.1 输出文件夹（脚本同目录下的 output_Shanya）
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output_Shanya')

# 2. 唯一成本表文件名（销售明细、利润表和未匹配统一使用）
COST_FILE_PATTERN = '山亚*成本*.xlsx'  # 兼容 山亚产品成本 / 山亚商品成本 / 山亚成本价

# 3. 参数配置
EXCHANGE_RATE = 0.2011    # 1 THB = 0.2011 CNY（2026-09 口径）
UNIT_FEE = 3.5              # 发单费 / 补单空包费 单价

# 4. 月份列控制
#    TARGET_MONTH = 0 → 自动按销售明细日期识别月份；非 0 → 强制指定（如 5 表示“5月”）
#    利润表会写入对应“{月}月”列：已存在该列则更新，不存在则新增到首个空列，
#    从而保留之前几个月的数据（多月累计）。
TARGET_MONTH = 0

# Sheet 名称映射（容错匹配）
SHEET_NAMES = {
    'order':     ['订单明细原表', '订单明细', '订单明细表'],
    'income':    ['收入明细表', '收入明细'],
    'brush':     ['补单明细 ', '补单明细', '补单'],
    'ads':       ['广告费', '广告'],
    'sales':     ['销售明细', '销售明细表'],
    'pnl':       ['利润表', 'Profit'],
    'unmatched': ['未匹配', '未匹配订单'],
}

# ---- 订单明细原表（泰文表头；优先按第1行表头名定位列，找不到再回退下面的固定列号）----
#      注：新店铺（敖广 SP 系列）导出多了列，SKU 在 U、数量在 Y、下单日期在 H、发货时间在 R，
#          与旧格式固定列号不同，故必须按表头定位。
ORDER_HEADER_ROW = 1
ORDER_COL_ID   = 1    # A  订单号
ORDER_COL_DATE = 7    # G  下单日期
ORDER_COL_SHIP = 17   # Q  发货时间
ORDER_COL_SKU  = 20   # T  SKU
ORDER_COL_QTY  = 24   # X  数量

# ---- 收入明细表（英文表头，第6行为表头，数据第7行起）----
INCOME_HEADER_ROW = 6

# ---- 补单明细（表头第1行，数据第2行起，遇空行停止）----
BRUSH_COL_ID    = 6    # F  线上单号
BRUSH_COL_COST  = 11   # K  垫付成本
BRUSH_COL_SUB   = 15   # O  垫付小计
BRUSH_COL_COMM  = 16   # P  佣金

# ---- 广告费（表头第7行，数据第8行起）----
ADS_HEADER_ROW = 7


# ================= 1. 基础工具函数 =================

def to_num(x):
    """把单元格值转成数值；自动清洗 ฿ ¥ , 空格 等文本符号。文本型数字也能转。"""
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace('฿', '').replace('¥', '').replace('PHP', '').replace('RM', '')
    s = s.replace('₱', '').replace(',', '').replace(' ', '').strip()
    if s == '' or s == '-':
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date_dt(x):
    """把日期字符串/对象转成 datetime（仅保留到日）。失败返回 None。"""
    if x is None or x == '':
        return None
    if isinstance(x, datetime):
        return datetime(x.year, x.month, x.day)
    s = str(x).strip()
    for fmt in ('%Y-%m-%d', '%Y/%m/%d'):
        try:
            d = datetime.strptime(s[:10], fmt)
            return datetime(d.year, d.month, d.day)
        except ValueError:
            continue
    return None


def get_sheet(wb, candidates):
    for name in candidates:
        if name in wb.sheetnames:
            return wb[name]
    return None


def build_unit_cost(cost_map):
    """SKU -> 单件成本。精确匹配优先，并兼容倍数与组合 SKU。"""
    def unit_cost(sku):
        if sku is None:
            return 0.0
        s = str(sku).strip().upper()
        if s == '' or s == 'NAN':
            return 0.0

        # 成本表中若已配置完整 SKU（如 SIRI-2001+2），必须以该配置为准。
        if s in cost_map:
            return cost_map[s]

        total = 0.0
        for part in s.split('+'):
            part = part.strip()
            if part in cost_map:
                total += cost_map[part]
                continue

            # A*2 表示两件 A；A*B 表示 A 与 B 的组合。
            m = re.fullmatch(r'(.+?)\*\s*(\d+)', part)
            if m:
                total += cost_map.get(m.group(1).strip(), 0.0) * int(m.group(2))
            elif '*' in part:
                total += sum(cost_map.get(p.strip(), 0.0) for p in part.split('*'))
            else:
                total += cost_map.get(part, 0.0)
        return total
    return unit_cost


def missing_cost_parts(sku, cost_map):
    """返回一个订单 SKU 中无法从唯一成本表解析出的基础 SKU。"""
    if sku is None:
        return []
    s = str(sku).strip().upper()
    if not s or s == 'NAN' or s in cost_map:
        return []

    missing = []
    for part in s.split('+'):
        part = part.strip()
        if part in cost_map:
            continue
        m = re.fullmatch(r'(.+?)\*\s*(\d+)', part)
        bases = [m.group(1).strip()] if m else [p.strip() for p in part.split('*')]
        missing.extend(base for base in bases if base and base not in cost_map)
    return missing


def thin_border_cell(ws, r, c):
    thin = Side(border_style="thin", color="000000")
    cell = ws.cell(r, c)
    cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
    cell.alignment = Alignment(horizontal='center', vertical='center')


def clean_label(s):
    """去掉 利润表 项目名的序号前缀，便于精确匹配。"""
    if s is None:
        return ''
    s = str(s).strip()
    s = re.sub(r'^[一二三四五六七八九十]+、', '', s)   # 去“一、”
    s = re.sub(r'^\d+\.', '', s)                       # 去“1.”
    s = s.replace('其中:', '').replace('其中：', '')    # 去“其中:”
    s = s.rstrip(':：').strip()
    return s


# ================= 2. 订单表 / 收入表列定位 =================

def locate_order_cols(ws):
    """根据订单明细原表第1行泰文表头定位关键列，找不到时回退旧版固定列号。"""
    names = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(ORDER_HEADER_ROW, c).value
        if v is not None:
            names[str(v).strip()] = c

    def find(pred, default):
        for nm, c in names.items():
            if pred(nm):
                return c
        return default

    return {
        'id':   find(lambda n: 'หมายเลขคำสั่งซื้อ' in n, ORDER_COL_ID),          # 订单号
        'date': find(lambda n: n in ('วันที่ทำการสั่งซื้อ', 'วันที่สั่งซื้อ'), ORDER_COL_DATE),  # 下单日期
        'ship': find(lambda n: n == 'เวลาส่งสินค้า', ORDER_COL_SHIP),            # 实际发货时间
        'sku':  find(lambda n: 'sku reference' in n.lower(), ORDER_COL_SKU),      # SKU 编码
        'qty':  find(lambda n: n == 'จำนวน', ORDER_COL_QTY),                      # 数量
    }



def locate_income_cols(ws):
    """根据第6行英文表头定位收入表关键列，返回列号(1-based)字典。失败回退固定列。"""
    names = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(INCOME_HEADER_ROW, c).value
        if v is not None:
            names[str(v).strip().lower()] = c

    def find(pred, default):
        for nm, c in names.items():
            if pred(nm):
                return c
        return default

    cols = {
        'id':     find(lambda n: n == 'order id', 2),                                   # B
        'date':   find(lambda n: 'order creation date' in n, 5),                        # E
        'price':  find(lambda n: 'original product price' in n, 12),                    # L
        'promo':  find(lambda n: 'seller product promotion' in n, 13),                 # M
        'refund': find(lambda n: 'refund amount' in n, 14),                            # N
        'voucher': find(lambda n: 'voucher sponsored by seller' in n and 'cofund' not in n, 16),  # P
        'taj_start': find(lambda n: 'shipping fee paid by buyer' in n, 20),            # T
        'taj_end':   find(lambda n: 'trade-in bonus by seller' in n, 36),             # AJ
    }
    return cols


# ================= 3. 单文件处理 =================

def process_single_file(source_path, cost_map):
    file_name = os.path.basename(source_path)
    print(f"\n{'=' * 60}\n>>> 正在处理: {file_name}")

    unit_cost = build_unit_cost(cost_map)

    # ---- 复制源文件到输出文件夹，得到可编辑副本 ----
    new_file_name = f"新_{file_name.replace('.xlsx', '')}.xlsx"
    file_output = os.path.join(OUTPUT_FOLDER, new_file_name)
    print(f"   生成副本: {new_file_name}")
    try:
        shutil.copy(source_path, file_output)
        os.chmod(file_output, stat.S_IRUSR | stat.S_IWUSR | stat.S_IRGRP | stat.S_IROTH)
        time.sleep(0.3)
    except Exception as e:
        print(f"   [错误] 复制失败: {e}")
        return

    try:
        wb = openpyxl.load_workbook(file_output)
    except Exception as e:
        print(f"   [错误] 打开副本失败: {e}")
        return

    # 输出不再需要“销售明细1”；源文件中即使已有也一并删除。
    for ws in list(wb.worksheets):
        if ws.title.strip() == '销售明细1':
            wb.remove(ws)

    ws_income = get_sheet(wb, SHEET_NAMES['income'])
    ws_order  = get_sheet(wb, SHEET_NAMES['order'])
    ws_brush  = get_sheet(wb, SHEET_NAMES['brush'])
    ws_ads    = get_sheet(wb, SHEET_NAMES['ads'])
    if ws_income is None or ws_order is None:
        print("   [跳过] 缺少收入明细表或订单明细原表")
        wb.close()
        return

    # ---------- A. 读取收入表：按订单号聚合 + 记录首次出现顺序 ----------
    ic = locate_income_cols(ws_income)
    income_agg = {}      # oid -> {date,price,promo,refund,voucher,taj}
    income_seq = []      # 订单号首次出现顺序
    for row in ws_income.iter_rows(min_row=INCOME_HEADER_ROW + 1):
        oid = row[ic['id'] - 1].value
        if oid is None or str(oid).strip() == '':
            continue
        oid = str(oid).strip()
        taj = sum(to_num(row[c - 1].value) for c in range(ic['taj_start'], ic['taj_end'] + 1))
        if oid not in income_agg:
            income_agg[oid] = {'date': row[ic['date'] - 1].value, 'price': 0.0,
                               'promo': 0.0, 'refund': 0.0, 'voucher': 0.0, 'taj': 0.0}
            income_seq.append(oid)
        d = income_agg[oid]
        d['price']   += to_num(row[ic['price'] - 1].value)
        d['promo']   += to_num(row[ic['promo'] - 1].value)
        d['refund']  += to_num(row[ic['refund'] - 1].value)
        d['voucher'] += to_num(row[ic['voucher'] - 1].value)
        d['taj']     += taj

    # ---------- B. 读取订单表：按订单号聚合 + 保留每行原始数据 ----------
    oc = locate_order_cols(ws_order)
    order_agg = {}       # oid -> {skus,qty,cost,ship,date}
    order_seq = []
    order_rows = []      # (oid, row_values_list, line_cost) 供未匹配整行复制
    missing_skus = set()
    n_order_cols = ws_order.max_column
    for row in ws_order.iter_rows(min_row=ORDER_HEADER_ROW + 1):
        oid = row[oc['id'] - 1].value
        if oid is None or str(oid).strip() == '':
            continue
        oid = str(oid).strip()
        sku = row[oc['sku'] - 1].value
        qty = to_num(row[oc['qty'] - 1].value)
        line_cost = unit_cost(sku) * qty
        missing_skus.update(missing_cost_parts(sku, cost_map))
        if oid not in order_agg:
            order_agg[oid] = {'skus': [], 'qty': 0.0, 'cost': 0.0,
                              'ship': row[oc['ship'] - 1].value, 'date': row[oc['date'] - 1].value}
            order_seq.append(oid)
        o = order_agg[oid]
        o['skus'].append(str(sku).strip() if sku is not None else '')
        o['qty'] += qty
        o['cost'] += line_cost
        if o['ship'] in (None, '', '-'):
            o['ship'] = row[oc['ship'] - 1].value
        order_rows.append((oid, [c.value for c in row[:n_order_cols]], line_cost,
                           [c for c in row[:n_order_cols]]))

    # 跨月订单（下单月≠本期月）：成本记 0、不计发单费（上月已计）
    if TARGET_MONTH:
        cur_month = TARGET_MONTH
    else:
        mc = {}
        for o in order_agg.values():
            dt = parse_date_dt(o['date'])
            if dt is not None:
                mc[dt.month] = mc.get(dt.month, 0) + 1
        cur_month = max(mc.items(), key=lambda kv: kv[1])[0] if mc else None
    cross_ids = set()
    if cur_month:
        for oid, o in order_agg.items():
            dt = parse_date_dt(o['date'])
            if dt is not None and dt.month != cur_month:
                cross_ids.add(oid)
                o['cost'] = 0.0
        order_rows = [(oid, v, 0.0 if oid in cross_ids else lc, c) for oid, v, lc, c in order_rows]
    if cross_ids:
        print(f"   [提示] 跨月订单 {len(cross_ids)} 单（成本记 0、不计发单费）: {', '.join(sorted(cross_ids))}")

    income_ids = set(income_agg)
    order_ids = set(order_agg)

    # 销售明细：以收入表顺序遍历，保留订单表中存在的订单
    sales_ids = [oid for oid in income_seq if oid in order_ids]
    # 未匹配：订单表有、收入表无
    unmatched_ids = [oid for oid in order_seq if oid not in income_ids]
    print(f"   收入订单 {len(income_ids)}  订单表订单 {len(order_ids)}  "
          f"销售明细 {len(sales_ids)}  未匹配订单 {len(unmatched_ids)}")
    if missing_skus:
        print(f"   [警告] 唯一成本表缺少 SKU: {', '.join(sorted(missing_skus))}（对应成本按 0 计算）")

    # ---------- C. 补单明细 ----------
    brush_ids = []
    brush_sub_sum = 0.0    # Σ垫付小计 O
    brush_cost_sum = 0.0   # Σ垫付成本 K
    brush_comm_sum = 0.0   # Σ佣金 P
    if ws_brush is not None:
        for row in ws_brush.iter_rows(min_row=2):
            bid = row[BRUSH_COL_ID - 1].value
            if bid is None or str(bid).strip() == '':
                break  # 数据区结束
            brush_ids.append(str(bid).strip())
            brush_sub_sum  += to_num(row[BRUSH_COL_SUB - 1].value)
            brush_cost_sum += to_num(row[BRUSH_COL_COST - 1].value)
            brush_comm_sum += to_num(row[BRUSH_COL_COMM - 1].value)
    n_brush = len(brush_ids)

    # ---------- D. 广告费：Amount 列求和后转为正数 ----------
    ads_amount_sum = 0.0
    if ws_ads is not None:
        ads_col = None
        for c in range(1, ws_ads.max_column + 1):
            v = ws_ads.cell(ADS_HEADER_ROW, c).value
            if v is not None and str(v).strip().lower() == 'amount':
                ads_col = c
                break
        if ads_col:
            for row in ws_ads.iter_rows(min_row=ADS_HEADER_ROW + 1):
                ads_amount_sum += to_num(row[ads_col - 1].value)
        else:
            print("   [警告] 广告费表找不到 Amount 列，广告费按 0 计算")

    # ---------- E. 写「销售明细」----------
    def fill_sales_sheet():
        ws_s = get_sheet(wb, SHEET_NAMES['sales'])
        if ws_s is None:
            ws_s = wb.create_sheet('销售明细')
            ws_s.append(['日期', '订单号', '商品SKU', '销售数量', '销售单价(฿)', '优惠金额(฿)',
                         '优惠劵金额(฿)', '退款金额(฿)', '销售金额(฿)', '销售金额(¥)',
                         '商品成本总额(¥)', '毛利(¥)'])
        if ws_s.max_row >= 2:
            ws_s.delete_rows(2, ws_s.max_row - 1)

        agg = {'sum_J': 0.0,        # Σ销售金额(¥)
               'sum_H': 0.0,        # Σ退款金额(฿)
               'sum_K': 0.0,        # Σ商品成本(¥)
               'J_by_id': {},       # oid -> J
               'taj_sum': 0.0}      # Σ销售订单 T:AJ
        for i, oid in enumerate(sales_ids):
            r = i + 2
            o = order_agg[oid]
            d = income_agg[oid]
            E, F, G, H = d['price'], d['promo'], d['voucher'], d['refund']
            I = E + F + G + H
            J = I * EXCHANGE_RATE
            K = o['cost']
            ws_s.cell(r, 1).value = d['date']   # 直接复制收入明细表 E 列原值
            ws_s.cell(r, 2).value = oid
            ws_s.cell(r, 3).value = ','.join(s for s in o['skus'] if s)
            ws_s.cell(r, 4).value = o['qty']
            ws_s.cell(r, 5).value = E
            ws_s.cell(r, 6).value = F
            ws_s.cell(r, 7).value = G
            ws_s.cell(r, 8).value = H
            ws_s.cell(r, 9).value  = f'=E{r}+F{r}+G{r}+H{r}'
            ws_s.cell(r, 10).value = f'=I{r}*{EXCHANGE_RATE}'
            ws_s.cell(r, 11).value = K
            ws_s.cell(r, 12).value = f'=J{r}-K{r}'
            for c in range(1, 13):
                thin_border_cell(ws_s, r, c)
            agg['sum_J'] += J
            agg['sum_H'] += H
            agg['sum_K'] += K
            agg['J_by_id'][oid] = J
            agg['taj_sum'] += d['taj']
        return agg

    sales_agg = fill_sales_sheet()

    # 利润表等后续计算仍基于「销售明细」（山亚产品成本.xlsx 口径）
    sum_sales_J = sales_agg['sum_J']
    sum_sales_H = sales_agg['sum_H']
    sum_sales_K = sales_agg['sum_K']
    sales_J_by_id = sales_agg['J_by_id']
    sales_taj_sum = sales_agg['taj_sum']
    last_sales_row = len(sales_ids) + 1

    # ---------- F. 写「未匹配」（整行连表头连格式复制 + BH 成本价）----------
    ws_un = get_sheet(wb, SHEET_NAMES['unmatched'])
    if ws_un is None:
        ws_un = wb.create_sheet('未匹配')
    if ws_un.max_row >= 1:
        ws_un.delete_rows(1, ws_un.max_row)

    bh_col = n_order_cols + 1
    # 表头：复制订单表第1行（含格式）+ BH 成本价(¥)
    for c in range(1, n_order_cols + 1):
        src = ws_order.cell(ORDER_HEADER_ROW, c)
        dst = ws_un.cell(1, c)
        dst.value = src.value
        if src.has_style:
            dst._style = copy(src._style)
    ws_un.cell(1, bh_col).value = '成本价(¥)'
    try:
        ws_un.row_dimensions[1].height = ws_order.row_dimensions[ORDER_HEADER_ROW].height or 15
    except Exception:
        ws_un.row_dimensions[1].height = 15

    unmatched_set = set(unmatched_ids)
    out_r = 1
    shipped_orders = set()       # 未匹配已发货订单（Q 非空）
    shipped_cost_total = 0.0     # 未匹配已发货订单成本
    seen_shipped_cost = set()
    for oid, vals, line_cost, cells in order_rows:
        if oid not in unmatched_set:
            continue
        out_r += 1
        for c in range(1, n_order_cols + 1):
            dst = ws_un.cell(out_r, c)
            dst.value = vals[c - 1]
            src = cells[c - 1]
            if getattr(src, 'has_style', False):
                dst._style = copy(src._style)
        ws_un.cell(out_r, bh_col).value = round(line_cost, 10)
        ws_un.row_dimensions[out_r].height = 15
        ship = order_agg[oid]['ship']
        if ship not in (None, '', '-') and oid not in cross_ids:
            shipped_orders.add(oid)
            if oid not in seen_shipped_cost:
                shipped_cost_total += order_agg[oid]['cost']
                seen_shipped_cost.add(oid)
    n_shipped = len(shipped_orders)

    # ---------- G. 计算利润表各明细值 ----------
    B6 = -sum_sales_H * EXCHANGE_RATE                       # 销售退款
    B4 = sum_sales_J + B6                                   # 本期销售
    B5 = sum(sales_J_by_id.get(b, 0.0) for b in brush_ids)  # 其中:补单金额
    B8 = sum_sales_K + shipped_cost_total                   # 实际产品成本
    n_sales_fee = sum(1 for oid in sales_ids if oid not in cross_ids)
    B9 = (n_sales_fee + n_shipped - n_brush) * UNIT_FEE      # 发单费
    B12 = abs(ads_amount_sum) * EXCHANGE_RATE                # 广告费：Amount 合计取正数 × 汇率
    B13 = brush_comm_sum                                    # 补单佣金
    brush_taj_sum = sum(income_agg.get(b, {}).get('taj', 0.0) for b in brush_ids)
    B14 = -brush_taj_sum * EXCHANGE_RATE                    # 补单平台费用（改正数=取正）
    B15 = brush_sub_sum - brush_cost_sum * EXCHANGE_RATE    # 补单汇率差额
    B16 = n_brush * UNIT_FEE                                # 补单空包费
    B18 = -sales_taj_sum * EXCHANGE_RATE - B14              # 平台费用（改正数=取正）- 补单部分

    detail_vals = {
        '本期销售': B4, '补单金额': B5, '销售退款': B6,
        '实际产品成本': B8, '发单费': B9, '广告费': B12,
        '补单佣金': B13, '补单平台费用': B14, '补单汇率差额': B15,
        '补单空包费': B16, '平台费用': B18, '预扣税': 0,
    }

    # ---------- H. 写「利润表」----------
    # 注：汇率不再写入表内单元格，统一使用程序开头的常量 EXCHANGE_RATE。
    ws_pnl = get_sheet(wb, SHEET_NAMES['pnl'])
    if ws_pnl is not None:
        target_header = f"{cur_month}月" if cur_month else None

        # 定位本月数据列（多月累计，保留其它月份）：
        #   第2行已存在“{月}月”表头 → 写入该列（更新本月）；
        #   不存在 → 写入首个空列并补表头（新增本月，旧列不动）。
        target_col = None
        first_empty = None
        for ci in range(2, ws_pnl.max_column + 2):
            cv = ws_pnl.cell(2, ci).value
            cvs = str(cv).strip() if cv is not None else ''
            if target_header and cvs == target_header:
                target_col = ci
                break
            if cvs == '' and first_empty is None:
                first_empty = ci
        if target_col is None:
            target_col = first_empty if first_empty is not None else (ws_pnl.max_column + 1)
            if target_header:
                ws_pnl.cell(2, target_col).value = target_header
        tcl = get_column_letter(target_col)

        # 扫描 A 列，定位每个项目所在行（精确匹配清洗后的名称）
        row_idx = {}
        for r in range(3, ws_pnl.max_row + 1):
            lab = clean_label(ws_pnl.cell(r, 1).value)
            if lab == '':
                continue
            key = lab.lower() if lab.lower() == 'shopee结算' else lab
            row_idx[key] = r

        # 写明细值
        for key, val in detail_vals.items():
            if key in row_idx:
                ws_pnl.cell(row_idx[key], target_col).value = val

        def ref(key):
            return f"{tcl}{row_idx[key]}" if key in row_idx else None

        def write_agg(key, plus=None, minus=None):
            if key not in row_idx:
                return
            plus = plus or []
            minus = minus or []
            terms = []
            for k in plus:
                rk = ref(k)
                if rk:
                    terms.append('+' + rk)
            for k in minus:
                rk = ref(k)
                if rk:
                    terms.append('-' + rk)
            if not terms:
                return
            formula = ''.join(terms)
            if formula[0] == '+':
                formula = formula[1:]
            ws_pnl.cell(row_idx[key], target_col).value = '=' + formula

        write_agg('实际主营收入', plus=['本期销售'], minus=['补单金额', '销售退款'])
        write_agg('实际主营成本', plus=['实际产品成本', '发单费'])
        write_agg('推广费用', plus=['广告费', '补单佣金', '补单平台费用', '补单汇率差额', '补单空包费'])
        write_agg('shopee结算', plus=['平台费用', '预扣税'])
        if '货损费用' in row_idx:
            if ws_pnl.cell(row_idx['货损费用'], target_col).value is None:
                ws_pnl.cell(row_idx['货损费用'], target_col).value = 0
            write_agg('团队运营费用', plus=['货损费用'])
        elif '团队运营费用' in row_idx:
            ws_pnl.cell(row_idx['团队运营费用'], target_col).value = 0
        write_agg('营业费用', plus=['推广费用', 'shopee结算', '团队运营费用'])
        profit_key = '本期利润' if '本期利润' in row_idx else ('毛利润' if '毛利润' in row_idx else None)
        if profit_key:
            write_agg(profit_key, plus=['实际主营收入'], minus=['实际主营成本', '营业费用'])

        # 备注：按月追加“{月}月未匹配”，保留以前月份的备注
        n_unshipped = sum(1 for oid in unmatched_ids if order_agg[oid]['ship'] in (None, '', '-'))
        n_shipped_all = len(unmatched_ids) - n_unshipped
        note = f"{cur_month}月未匹配：取消未发货{n_unshipped}，取消已发货{n_shipped_all}"
        start = (row_idx[profit_key] + 1) if profit_key else 3
        note_row = None
        for r in range(start, ws_pnl.max_row + 2):
            lab = str(ws_pnl.cell(r, 1).value or '').strip()
            if lab.startswith(f"{cur_month}月未匹配") or lab.startswith('备注'):
                note_row = r
                break
            if lab == '':
                note_row = r
                break
        ws_pnl.cell(note_row, 1).value = note

    # ---------- 保存 ----------
    try:
        wb.save(file_output)
        wb.close()
        print(f"   [成功] 本期销售={B4:.2f}  实际产品成本={B8:.2f}  发单费={B9:.0f}  "
              f"平台费用={B18:.2f}")
    except Exception as e:
        print(f"   [失败] 保存出错: {e}")


# ================= 4. 主入口 =================

def load_cost_map(cost_path):
    wb = openpyxl.load_workbook(cost_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows())
    if not rows:
        wb.close()
        return {}
    header = [str(c.value).strip().lower() if c.value is not None else '' for c in rows[0]]
    sku_i = next((i for i, h in enumerate(header) if 'sku' in h), 0)
    cost_i = next((i for i, h in enumerate(header) if '成本' in h or 'cost' in h), 1)
    cost_map = {}
    for row in rows[1:]:
        if sku_i >= len(row):
            continue
        sku = row[sku_i].value
        if sku is None or str(sku).strip() == '':
            continue
        cost_map[str(sku).strip().upper()] = to_num(row[cost_i].value if cost_i < len(row) else 0)
    wb.close()
    return cost_map


def main():
    print(f"启动批量处理... 目标文件夹: {TARGET_FOLDER}")
    if not os.path.exists(TARGET_FOLDER):
        print(f"错误: 找不到文件夹 '{TARGET_FOLDER}'")
        return

    cost_files = [p for p in glob.glob(os.path.join(TARGET_FOLDER, COST_FILE_PATTERN))
                  if not os.path.basename(p).startswith('~$')]
    if not cost_files:
        print(f"错误: 找不到成本表 '{COST_FILE_PATTERN}'")
        return
    cost_path = cost_files[0]
    cost_name = os.path.basename(cost_path)

    print(">>> 加载商品成本表...")
    try:
        cost_map = load_cost_map(cost_path)
        print(f"   成本表 SKU 数: {len(cost_map)}（全部成本统一使用）")
    except Exception as e:
        print(f"   读取成本表失败: {e}")
        return

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    xlsx_files = glob.glob(os.path.join(TARGET_FOLDER, "*.xlsx"))
    count = 0
    for file_path in xlsx_files:
        fn = os.path.basename(file_path)
        if fn == cost_name or fn.startswith("新_") or fn.startswith("~$"):
            continue
        process_single_file(file_path, cost_map)
        count += 1

    print(f"\n{'=' * 60}\n批量处理完成！共 {count} 个文件。\n新文件存放于: {OUTPUT_FOLDER}")
    try:
        input("按回车键退出...")
    except EOFError:
        pass


if __name__ == '__main__':
    main()
