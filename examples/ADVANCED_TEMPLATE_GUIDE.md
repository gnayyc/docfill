# 如何建立進階模板

這份指南說明如何在 Microsoft Word 中建立使用進階 Jinja2 語法的模板。

## 範例：專案報告模板

使用 `advanced_example.yaml` 作為設定檔，在 Word 中建立以下模板：

---

## 模板內容（在 Word 中直接輸入）

```
                    專案進度報告

專案名稱：{{project_name}}
專案編號：{{project_code}}
報告日期：{{date}}
專案經理：{{manager}}

═══════════════════════════════════════════

## 一、團隊成員

{% for member in team_members %}
{{loop.index}}. {{member.name}} - {{member.role}}
   部門：{{member.department}}
   年資：{{member.experience}} 年
   {% if member.experience >= 7 %}
   ⭐ 資深成員
   {% endif %}
{% endfor %}

═══════════════════════════════════════════

## 二、專案進度

{% for milestone in milestones %}
【{{milestone.phase}}】
進度：{{milestone.completion}}%
狀態：{{milestone.status}}
預定日期：{{milestone.date}}

{% if milestone.completion == 100 %}
✅ 已完成
{% elif milestone.completion > 0 %}
🔄 進行中 - 已完成 {{milestone.completion}}%
{% else %}
⏳ 待開始
{% endif %}

{% endfor %}

整體進度：{{overall_progress}}%

{% if overall_progress >= 80 %}
📊 專案進度良好
{% elif overall_progress >= 50 %}
📊 專案進度正常
{% else %}
⚠️ 專案進度落後
{% endif %}

═══════════════════════════════════════════

## 三、預算使用狀況

總預算：$ {{budget.total | format_number}}
已使用：$ {{budget.spent | format_number}}
剩餘：$ {{budget.remaining | format_number}}

使用率：{{ (budget.spent / budget.total * 100) | round(1) }}%

{% if budget.remaining < budget.total * 0.2 %}
⚠️ 警告：預算所剩不多，請注意控制開支
{% endif %}

═══════════════════════════════════════════

## 四、風險管理

{% for risk in risks %}
• {{risk.risk}}
  風險等級：{% if risk.level == "高" %}🔴{% elif risk.level == "中" %}🟡{% else %}🟢{% endif %} {{risk.level}}
  應對措施：{{risk.mitigation}}

{% endfor %}

═══════════════════════════════════════════

## 五、專案狀態總結

當前狀態：{{status}}

{% if is_delayed %}
⚠️ 專案進度延遲
{% else %}
✅ 專案進度正常
{% endif %}

{% if needs_attention %}
⚡ 需要管理層關注的事項：
- 請參閱風險管理章節
- 建議進行專案審查會議
{% endif %}

═══════════════════════════════════════════

報告人：{{manager}}
報告日期：{{date}}
```

---

## 使用方法

### 1. 在 Word 中建立模板

1. 開啟 Microsoft Word
2. 複製上面的模板內容（包含所有 `{{}}` 和 `{% %}` 標籤）
3. 貼上到 Word 文件中
4. 調整格式、字體、顏色等（保持 `{{}}` 和 `{% %}` 標籤不變）
5. 儲存為 `advanced_template.docx`

### 2. 執行 DocFill

```bash
# 基本用法
docfill advanced_example.yaml advanced_template.docx -o report.docx

# 含詳細輸出
docfill advanced_example.yaml advanced_template.docx -o report.docx -v

# 同時產生 PDF
docfill advanced_example.yaml advanced_template.docx -o report.docx --pdf
```

### 3. 查看結果

開啟 `report.docx`，你會看到：
- 團隊成員清單已自動展開（3 位成員）
- 每個里程碑都有對應的狀態圖示
- 預算使用率已自動計算
- 風險等級用不同顏色標示
- 根據條件顯示不同的警告訊息

---

## 進階技巧

### 1. 表格中使用迴圈

在 Word 表格中使用 `{%tr %}` 標籤：

| 姓名 | 角色 | 部門 |
|------|------|------|
| {%tr for member in team_members %} | | |
| {{member.name}} | {{member.role}} | {{member.department}} |
| {%tr endfor %} | | |

### 2. 條件格式化

```
預算狀態：
{% if budget.remaining > budget.total * 0.5 %}
🟢 充足
{% elif budget.remaining > budget.total * 0.2 %}
🟡 適中
{% else %}
🔴 緊張
{% endif %}
```

### 3. 數學運算

```
使用率：{{ (budget.spent / budget.total * 100) | round(2) }}%
平均年資：{{ (team_members | sum(attribute='experience') / team_members | length) | round(1) }} 年
```

### 4. 文字處理

```
大寫：{{ project_name | upper }}
首字母大寫：{{ project_name | title }}
截斷：{{ project_name | truncate(20) }}
```

---

## 注意事項

1. **保持標籤完整**：確保 `{{` `}}` 和 `{%` `%}` 成對出現
2. **縮排一致**：`{% for %}` 和 `{% endfor %}` 要對齊
3. **變數名稱**：必須與 YAML 檔案中的 key 完全一致
4. **Word 格式**：可以在 Word 中設定字體、顏色、大小等，不影響模板功能

---

## 測試範例

執行測試：

```bash
# 進入 examples 目錄
cd examples

# 建立你的進階模板（在 Word 中）
# 儲存為 advanced_template.docx

# 執行 DocFill
docfill advanced_example.yaml advanced_template.docx -v
```

這會產生 `advanced_template_filled.docx`，包含所有動態生成的內容！
