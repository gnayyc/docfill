# DocFill 進階功能指南

DocFill 基於 [docxtpl](https://docxtpl.readthedocs.io/) 套件，完全支援 Jinja2 模板語法。

## 🎯 重要提醒

**無需任何額外參數或設定！**所有功能都是自動支援的。你在模板中使用什麼語法，DocFill 就會自動處理。

---

## 基本語法（你已經在用的）

### 簡單變數替換
```yaml
# config.yaml
name: "張三"
date: "2024-10-23"
```

```
模板：姓名：{{name}}，日期：{{date}}
結果：姓名：張三，日期：2024-10-23
```

### 巢狀資料存取
```yaml
# config.yaml
company:
  name: "ABC公司"
  contact:
    phone: "02-1234-5678"
    email: "info@abc.com"
```

```
模板：公司：{{company.name}}
      電話：{{company.contact.phone}}
結果：公司：ABC公司
      電話：02-1234-5678
```

---

## 🚀 進階語法（現在也可以用！）

### 1. 條件判斷（If/Else）

#### 設定檔
```yaml
# config.yaml
employee:
  name: "李四"
  salary: 50000
  performance: "優良"
```

#### Word 模板
```
員工：{{employee.name}}

{% if employee.salary > 45000 %}
薪資等級：高
獎金：10%
{% else %}
薪資等級：一般
獎金：5%
{% endif %}

{% if employee.performance == "優良" %}
🎉 恭喜！您獲得年度績效獎勵
{% endif %}
```

#### 輸出結果
```
員工：李四

薪資等級：高
獎金：10%

🎉 恭喜！您獲得年度績效獎勵
```

---

### 2. 迴圈（For Loop）

#### 設定檔
```yaml
# config.yaml
projects:
  - name: "專案A"
    budget: 100000
    status: "進行中"
  - name: "專案B"
    budget: 200000
    status: "已完成"
  - name: "專案C"
    budget: 150000
    status: "規劃中"

total_budget: 450000
```

#### Word 模板
```
專案列表：

{% for project in projects %}
{{ loop.index }}. {{project.name}}
   預算：{{project.budget}} 元
   狀態：{{project.status}}
{% endfor %}

總預算：{{total_budget}} 元
```

#### 輸出結果
```
專案列表：

1. 專案A
   預算：100000 元
   狀態：進行中
2. 專案B
   預算：200000 元
   狀態：已完成
3. 專案C
   預算：150000 元
   狀態：規劃中

總預算：450000 元
```

---

### 3. 迴圈變數

Jinja2 在迴圈中提供特殊變數：

| 變數 | 說明 |
|------|------|
| `loop.index` | 當前迭代索引（從 1 開始） |
| `loop.index0` | 當前迭代索引（從 0 開始） |
| `loop.first` | 是否為第一次迭代 |
| `loop.last` | 是否為最後一次迭代 |
| `loop.length` | 迴圈總長度 |

#### 範例
```
{% for item in items %}
{% if loop.first %}=== 開始 ==={% endif %}
第 {{loop.index}}/{{loop.length}} 項：{{item}}
{% if loop.last %}=== 結束 ==={% endif %}
{% endfor %}
```

---

### 4. 過濾器（Filters）

Jinja2 提供豐富的過濾器來處理資料：

#### 文字處理
```yaml
name: "john doe"
text: "  hello world  "
```

```
大寫：{{ name | upper }}           # JOHN DOE
首字母大寫：{{ name | title }}       # John Doe
去空白：{{ text | trim }}           # hello world
長度：{{ name | length }}           # 8
```

#### 數字處理
```yaml
price: 1234.567
```

```
四捨五入：{{ price | round(2) }}     # 1234.57
整數：{{ price | int }}             # 1234
```

#### 日期處理
```yaml
date: "2024-10-23"
```

```
格式化：{{ date }}
```

#### 預設值
```yaml
description: ""  # 空值
```

```
{{ description | default("無說明") }}  # 輸出：無說明
```

---

### 5. 複雜範例：組合使用

#### 設定檔
```yaml
# invoice.yaml
invoice_number: "INV-2024-001"
date: "2024-10-23"
customer:
  name: "ABC公司"
  address: "台北市信義區"

items:
  - name: "產品A"
    quantity: 10
    price: 1000
    total: 10000
  - name: "產品B"
    quantity: 5
    price: 2000
    total: 10000
  - name: "產品C"
    quantity: 3
    price: 5000
    total: 15000

subtotal: 35000
tax_rate: 0.05
tax: 1750
total: 36750
```

#### Word 模板
```
                        發票

發票編號：{{invoice_number}}
日期：{{date}}

客戶資訊：
  {{customer.name}}
  {{customer.address}}

─────────────────────────────────────────
項目                 數量    單價      金額
─────────────────────────────────────────
{% for item in items %}
{{item.name | ljust(20)}} {{item.quantity}} {{item.price}} {{item.total}}
{% endfor %}
─────────────────────────────────────────

小計：$ {{subtotal}}
稅金 ({{(tax_rate * 100)|int}}%)：$ {{tax}}
─────────────────────────────────────────
總計：$ {{total}}

{% if total > 30000 %}
⚠️ 此發票金額超過 30,000 元，需主管核准
{% endif %}
```

---

## 📚 更多進階功能

### 表格內的迴圈

在 Word 表格中也可以使用迴圈！使用特殊標籤 `{%tr %}` 來標記要重複的表格列。

#### 設定檔
```yaml
employees:
  - name: "張三"
    department: "工程部"
    salary: 50000
  - name: "李四"
    department: "行銷部"
    salary: 45000
```

#### 在 Word 中建立表格
| 姓名 | 部門 | 薪資 |
|------|------|------|
| {%tr for emp in employees %} | | |
| {{emp.name}} | {{emp.department}} | {{emp.salary}} |
| {%tr endfor %} | | |

**注意**：`{%tr %}` 標籤要放在表格的儲存格中。

---

## 💡 使用技巧

### 1. 保持模板可讀性
```
# ✅ 好的做法
{% if condition %}
內容
{% endif %}

# ❌ 避免
{% if condition %}內容{% endif %}
```

### 2. 使用註解
```
{# 這是註解，不會出現在最終文件中 #}
```

### 3. 空白行控制
```
# 去除前後空白
{%- if condition -%}
內容
{%- endif -%}
```

---

## 🔗 更多資源

- [Jinja2 官方文件](https://jinja.palletsprojects.com/)
- [docxtpl 官方文件](https://docxtpl.readthedocs.io/)

---

## ❓ 常見問題

**Q: 需要特別的命令列參數來啟用進階功能嗎？**
A: 不需要！直接在模板中使用進階語法即可，DocFill 會自動處理。

**Q: 簡單的 {{key}} 語法還能用嗎？**
A: 當然可以！所有原有功能都完全相容。

**Q: 我可以混合使用簡單和進階語法嗎？**
A: 可以！在同一份模板中可以同時使用 {{key}} 和 {% if %} 等語法。

**Q: 測試我的模板語法正確嗎？**
A: 使用 `--check-placeholders` 參數檢查基本的變數替換，或直接執行看結果。

---

**享受更強大的文件生成能力！** 🚀
