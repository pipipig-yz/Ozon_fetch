# Ozon Finance Report（Ozon 财务应计报告）

从 [Ozon Seller API](https://docs.ozon.ru/api/seller/zh/) 拉取财务应计数据（accruals），生成多 Sheet 的 Excel 报告。

## 功能

- 按日期范围拉取 Ozon 应计记录（v1 LTS 接口，长期稳定）
- 生成 6 个 Sheet 的 Excel 报告：
  - **按订单号排列** — 按基础单号聚合，适合对账
  - **按操作排列** — 每条应计记录展开为一行
  - **应计记录** — 原始应计数据概览
  - **费用明细** — 所有费用项展开（item_fees + non_item_fee + 国际配送）
  - **商品佣金** — 每 SKU 的销售佣金详情
  - **费用类型** — type_id → 名称映射表
- 自动汇总：各费用类别的笔数和总金额
- 支持按天数快捷查询，也支持指定起止日期
- 可选的 JSON 原始数据导出

## 前置条件

1. Ozon 卖家账号，并在后台生成 API Key（**Finance → 财务权限**必须开启）
2. Python 3.10+

## 安装

```bash
git clone <repo-url> && cd ozon

# 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

## 配置

在项目根目录创建 `.env` 文件：

```ini
OZON_CLIENT_ID=你的Client-Id
OZON_API_KEY=你的Api-Key
```

或者直接设置环境变量：

```bash
export OZON_CLIENT_ID=你的Client-Id
export OZON_API_KEY=你的Api-Key
```

> `.env` 已加入 `.gitignore`，不会被提交到 Git。

## 用法

```bash
# 拉取最近 7 天 → report.xlsx
python fetch_transactions.py

# 拉取最近 30 天
python fetch_transactions.py --days 30

# 指定日期范围
python fetch_transactions.py --from 2026-05-01 --to 2026-07-05

# 自定义输出路径
python fetch_transactions.py --output my_report.xlsx

# 同时导出原始 JSON（方便调试）
python fetch_transactions.py --json raw_data.json

# 终端打印汇总表（不打开 Excel 也能快速对账）
python fetch_transactions.py -s
python fetch_transactions.py --days 30 -s
```

### 完整参数

| 参数 | 简写 | 默认值 | 说明 |
|------|------|--------|------|
| `--from` | — | 7 天前 | 开始日期 `YYYY-MM-DD` |
| `--to` | — | 今天 | 结束日期 `YYYY-MM-DD` |
| `--days` | — | `7` | 往回查几天（和 `--from`/`--to` 互斥） |
| `--output` | `-o` | `report.xlsx` | Excel 输出路径 |
| `--summary` | `-s` | 否 | 在终端打印按订单号聚合的汇总表 |
| `--json` | — | 无 | 追加导出原始 JSON 到此文件 |
| `--max-pages` | — | `100` | 每天游标翻页上限 |

## 验证 API 权限

如果不确定你的 API Key 是否有 finance 权限，可以先跑：

```bash
python test_roles.py
```

它会依次调用：
1. `POST /v1/roles` — 列出当前 Key 的所有权限路径
2. `POST /v1/finance/cash-flow-statement/list` — 临时测试 v1 现金流接口
3. `POST /v3/finance/transaction/list` — 对比测试 v3 交易接口（旧版）

如果输出里出现 `✅ Finance paths found`，说明权限正确。

## Excel 报告说明

### Sheet 1 — 按订单号排列

按基础单号（去掉末尾 `-N` 后缀）聚合。同一基础单号下的所有 SKU 的金额会合并为一行。

| 列 | 说明 |
|----|------|
| 基础单号 | 去掉 `-1`/`-2` 后缀的单号 |
| 日期 | 该单号涉及的所有日期范围 |
| SKU | 涉及的 SKU 列表 |
| 收入 | `sale_amount` 合计 |
| 收单 | Acquiring 费用（支付手续费） |
| 销售佣金 | `sale_commission` |
| 配送佣金 | 国内配送相关费用 |
| 国际配送 | 国际物流费用 |
| 关联单号 | 组成该基础单号的所有原始单号 |

### Sheet 2 — 按操作排列

每条应计记录按 SKU 展开，一个 (应计ID, SKU) 对应一行。

### Sheet 3 — 应计记录

每个应计 ID 一行，包含投递模式、应计类别、总金额、非商品费用等。

### Sheet 4 — 费用明细

每条费用的详细展开，标注来源（`item_fees` / `non_item_fee` / `delivery_service`）。

### Sheet 5 — 商品佣金

每 SKU 的佣金明细：销售金额、售价、销售佣金、平台佣金、佣金比率、奖金、共同投资等。

### Sheet 6 — 费用类型

Ozon 系统中定义的费用类型 ID → 名称 → 描述的映射表，方便对照。

## 项目结构

```
.
├── fetch_transactions.py   # 主脚本：拉取数据 + 生成 Excel
├── test_roles.py           # 辅助：验证 API Key 权限
├── requirements.txt        # Python 依赖
├── .env                    # API 凭据（不入 Git）
├── .gitignore
├── json_keys_mapping.md    # v3 接口字段参考（已弃用，保留供参考）
└── README.md
```

## 常见问题

### 报错 "Missing credentials"

脚本找不到 `OZON_CLIENT_ID` 和 `OZON_API_KEY`。确保 `.env` 文件在脚本同目录下，或已设置环境变量。

### 返回空数据

- 检查开始/结束日期是否正确（注意时区，Ozon 用莫斯科时间 MSK，UTC+3）
- 用 `test_roles.py` 确认 API Key 有 finance 权限
- 检查 Ozon 后台该时间段是否确实有订单

### 日期范围太大导致超时

每天一个 API 请求 + 游标翻页，大范围会很慢。建议每次查不超过 90 天，或用 cron 每天增量拉取。

## License

Internal use.
