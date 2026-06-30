# Ozon Transaction JSON 字段映射说明

> 基于 `/v3/finance/transaction/list` API 返回的数据结构

---

## 字段详解

| JSON Key | 中文含义 | 说明 |
|---|---|---|
| `operation_id` | 操作ID | Ozon 系统中该笔交易的唯一编号 |
| `operation_type` | 操作类型（代码） | 内部类型标识，如 `OperationAgentDeliveredToCustomer` |
| `operation_date` | 操作日期 | 该笔交易发生的日期 |
| `operation_type_name` | 操作类型名称（俄文） | 人类可读的操作描述，如 "Доставка покупателю"（送达买家） |
| `delivery_charge` | 配送费 | 商家承担的配送费用（通常为 0，配送费通过其他条目扣减） |
| `return_delivery_charge` | 退货配送费 | 退货产生的配送费用 |
| `accruals_for_sale` | 销售应计收入 | 商品销售给买家后计入商家的收入（正数） |
| `sale_commission` | 销售佣金 | Ozon 平台从销售额中扣除的佣金（负数） |
| `amount` | 最终金额 | 该笔操作实际计入商家账户的金额（正=收入，负=支出） |
| `type` | 交易大类 | `orders`（订单）/ `services`（服务费）/ `other`（其他） |
| `posting.delivery_schema` | 配送模式 | 如 `RFBS`（商家自发货）/ `FBS` / `FBO` |
| `posting.order_date` | 下单日期 | 买家下单的时间 |
| `posting.posting_number` | 发货单号 | Ozon 系统内的发货编号 |
| `posting.warehouse_id` | 仓库ID | 商品所在仓库的标识 |
| `items[].name` | 商品名称 | 销售的商品全名（俄文） |
| `items[].sku` | 商品 SKU | Ozon 系统中的商品 SKU 编号 |
| `services[].name` | 服务项名称 | 扣费服务的名称 |
| `services[].price` | 服务项金额 | 单项服务的扣费金额 |

---

## 金额字段关系

```
accruals_for_sale (销售收入) + sale_commission (销售佣金) = amount 的基础
此外 delivery_charge / return_delivery_charge / services 等会进一步影响 amount
```

- **正数 `amount`** → 商家收入（如订单送达）
- **负数 `amount`** → 商家支出（如佣金、罚款、配送费）

---

## `type` 字段取值说明

| type 值 | 含义 |
|---|---|
| `orders` | 订单相关交易（收入/退款/配送费重结算） |
| `services` | 服务费（佣金、物流代理费等） |
| `other` | 其他费用（如收单/支付手续费） |

---

## 嵌套对象结构示意

```
transaction
├── operation_id          (int)
├── operation_type        (string)
├── operation_date        (string, datetime)
├── operation_type_name   (string)
├── delivery_charge       (float)
├── return_delivery_charge(float)
├── accruals_for_sale     (float)
├── sale_commission       (float)
├── amount                (float)
├── type                  (string: orders | services | other)
├── posting
│   ├── delivery_schema   (string)
│   ├── order_date        (string, datetime)
│   ├── posting_number    (string)
│   └── warehouse_id      (int)
├── items[]               (array)
│   ├── name              (string)
│   └── sku               (int)
└── services[]            (array)
    ├── name              (string)
    └── price             (float)
```
