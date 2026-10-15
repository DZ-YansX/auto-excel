# auto-excel

跨境电商（Shopee / TikTok Shop）店铺月度报表自动化脚本集合。把平台导出的账单、订单明细、广告费，以及 BigSeller（大卖家）导出的成本数据，自动填进各店铺的利润表模板，生成 **销售明细 / 利润表 / 未匹配** 等工作表。

> 仓库只保存脚本。店铺报表、成本表等业务数据都不入库（见 `.gitignore`），使用时自己放进对应的输入目录。

## 目录

| 目录 | 脚本 | 用途 |
| --- | --- | --- |
| `Shopee/` | `shopee_Shanya.py` | 虾皮·山亚（泰国站 THB）利润表 |
| `Shopee/` | `shopee_Shunchang.py` | 虾皮·顺昌（菲律宾 / 印尼 / 马来）利润表 |
| `Shopee_bigseller/Shopee_wanshan/` | `Shopee_BigSeller.py` | 虾皮·万山（大卖家成本口径，印尼 / 菲 / 马 / 泰）利润表 |
| `TK no horse/` | `TK_NoHorse_Shanya.py` | TikTok·山亚利润表（成本取山亚成本表） |
| `TK no horse/` | `TK_NoHorse_Shunchang.py` | TikTok·顺昌利润表 |
| `TK_Bigseller/` | `TK_Bigseller.py` | TikTok·大卖家（BigSeller）利润表，支持中/英/印尼语账单表头 |
| `填表/TK/` | `fill_tk.py` | 按店铺拆分 TK 大卖家订单，写入各店铺工作簿的「大卖家」sheet |
| `填表/虾皮/` | `fill_shopee.py` | 按店铺拆分虾皮大卖家 + 补单（测评）明细，写入各店铺模板 |
| `补成本/` | `补成本.py` | 批量补齐大卖家表的「商品成本」「总成本价」 |
| `格式化/` | `format_workbooks.py` | 批量清理 / 格式化店铺月度工作簿，为下个月做准备 |

`Myexcel/` 是早期版本，单独维护。

## 环境

- Python 3.9+
- 依赖：

```bash
pip install pandas numpy openpyxl
```

## 月度流程

一般顺序：

1. **补成本**：大卖家导出表放进 `补成本/`，和 `万山成本价.xlsx` 一起运行 `补成本.py`，得到 `<原名>_已填成本.xlsx`。
2. **填表**：把大卖家表（及虾皮的测评补单表）和各店铺模板放进 `填表/TK/数据源/` 或 `填表/虾皮/数据源/`，运行 `fill_tk.py` / `fill_shopee.py`，结果在 `输出/`，同时生成 `订单核对报告.txt`。
3. **算利润表**：把填好的店铺表放进对应脚本的输入目录，改好本月汇率后运行，结果输出为 `新_<原名>.xlsx`。
4. **格式化**（可选）：用 `format_workbooks.py` 清空上月原始数据 sheet，生成下月模板。

## 各脚本说明

### 利润表脚本

| 脚本 | 输入目录 | 成本表 | 输出目录 |
| --- | --- | --- | --- |
| `Shopee/shopee_Shanya.py` | `Shopee_Data_Shanya/` | 同目录内 `山亚*成本.xlsx` | `output_Shanya/` |
| `Shopee/shopee_Shunchang.py` | `Shopee_Data_Shunchang/` | `顺昌商品成本.xlsx` | `output_Shunchang/` |
| `Shopee_bigseller/Shopee_wanshan/Shopee_BigSeller.py` | `ShopeeDeal/` | 只用店铺表内「大卖家」sheet 的总成本价 | `output_wanshan/` |
| `TK no horse/TK_NoHorse_Shanya.py` | `TKdeal_Shanya/` | `TKdeal_Shanya/山亚成本价.xlsx` | `output_Shanya/` |
| `TK no horse/TK_NoHorse_Shunchang.py` | `TKdeal_Shunchang/` | 脚本同目录 `商品成本.xlsx` | `output_Shunchang/` |
| `TK_Bigseller/TK_Bigseller.py` | `TKdeal/` | 大卖家优先；`万山成本价.xlsx`（脚本目录或 `../补成本/`）兜底未匹配已发货 | `output/` |

输入目录、输出目录都在脚本所在目录下，可以直接双击或命令行运行：

```bash
python TK_Bigseller/TK_Bigseller.py
```

部分脚本运行结束会等待「按回车键退出」，批量 / 自动运行时可用 `echo | python 脚本.py`。

**每月运行前先检查配置区**（文件顶部 `配置区域`）：

- 汇率：`EXCHANGE_RATE` / `CURRENCY_RATES` / `EXCHANGE_RATE_MAP` / `CURRENCY_RATE_MAP` 等，每月按财务给的汇率更新；山亚和大卖家的汇率可能不同，要分开确认。
- 发单费单价：`UNIT_FEE`、`DISPATCH_FEES` / `COUNTRY_UNIT_FEES`（印尼 3，其余 3.5，顺昌 3）。
- 目标月份：`TARGET_MONTH = 0` 表示自动识别；需要时可强制指定。

共同口径：

- **销售明细**：账单里的订单行，且能在订单明细表找到的订单。
- **未匹配**：订单明细里有、账单里没有的订单，按状态分为取消未发货 / 取消已发货 / 发货未到账。
- **跨月订单**（下单月 ≠ 本期月）：成本记 0、不计发单费（上月已计）。
- **利润表多月累计**：按「N月」表头写入对应列，没有该列就在最后一个月份列右侧新建，历史月份不动。
- 查不到成本的单元格填 0 并**标黄**，控制台也会提示，需要手工补。

### `填表/fill_tk.py`、`填表/fill_shopee.py`

```bash
python 填表/TK/fill_tk.py                 # 处理 数据源/ 下所有店铺
python 填表/TK/fill_tk.py <店铺文件名>      # 只处理一个店铺
python 填表/虾皮/fill_shopee.py --exclude <模板名> --skip-review
```

- `--output-dir`：输出目录，默认 `输出/`
- `--exclude`（虾皮）：排除属于其他平台的模板，可重复
- `--skip-review`（虾皮）：只填大卖家，不动补单明细（确认本月没有补单时用）

### `补成本/补成本.py`

```bash
cd 补成本
python 补成本.py                                   # 处理当前目录所有大卖家表
python 补成本.py 表1.xlsx 表2.xlsx 万山成本价.xlsx
```

注意：不带参数时扫描的是**当前工作目录**，要先 `cd` 到 `补成本/`。

按「店铺SKU」匹配成本：先用本表同国家同 SKU 的已有成本反推，再查万山成本价；YY / XY / THS / WSS 前缀查不到时按 `*-KP` 兜底。两处都查不到的标黄。

### `格式化/format_workbooks.py`

```bash
cd 格式化
python format_workbooks.py                    # 默认 --input deal --output output（相对当前目录）
```

清空账单 / 订单明细 / 广告费 / 补单明细 / 大卖家等原始数据 sheet，整理未匹配和利润表公式，用于生成下个月的空白模板。

## 注意事项

- 账单表头可能是中文、英文或印尼语，列名也会随平台改版变化；如果销售额 / 成本整列变成 0，先检查表头别名（`BILL_COL_ALIASES` 等）是否需要补充。
- 菲律宾 TK 订单明细日期是美式 `MM/DD/YYYY h:mm:ss AM/PM`，其它国家为 `DD/MM/YYYY`。
- 当地时间月末最后几个小时下的订单，大卖家按北京时间导出时会被划到下个月，导致这些订单没有成本，需要手工补。
- 跑完后建议核对：销售明细每行 E+F+G 应等于账单「总收入 / Total Revenue」。
