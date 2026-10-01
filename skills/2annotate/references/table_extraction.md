# 表格擷取規則（Table Extraction Reference）

> 本文件是 2annotate skill 的內部 reference，描述如何從論文 PDF
> 擷取表格並輸出為 `tables/table_X.json` 檔案。
>
> 主 SKILL.md 在 Step 3 會引用本文件。

## 概覽

主 skill 在處理論文時，必須為每個出現在論文中的 Table 同時產出一個獨立的 JSON 檔案，
存到 `tables/table_X.json`。這些檔案會被 `generate.py` 在產生 HTML 時自動載入，
渲染為左欄的 sticky 表格區塊。

## 與 paper.json 的關係

```
paper.json                              tables/table_X.json
─────────                              ────────────────────
table_load block  ──table_id 對應──→  id 欄位
table_ref block   ──row/col 索引──→   rows[row][col]
```

paper.json 中的 `table_load` block 只記錄 `table_id`（如 `"table_1"`），
實際的表格數據存在 `tables/table_1.json` 中。

paper.json 中的 `table_ref` block 用 0-based 的 `row`/`col` 索引
引用 `table_X.json` 的 `rows` 陣列。兩者**必須**對應一致。

---

## 輸出格式

每個表格一個 JSON 檔案：

```
tables/
  table_1.json
  table_2.json
  ...
```

### table_X.json 結構

```json
{
  "id": "table_1",
  "caption_en": "Table 1: Comparison with state-of-the-art on COCO val2017.",
  "caption_zh": "表 1：在 COCO val2017 上與現有最佳方法的比較。",
  "headers": ["Method", "Backbone", "AP", "AP50", "AP75"],
  "rows": [
    ["Faster R-CNN", "ResNet-50", "37.4", "58.1", "40.4"],
    ["DETR",         "ResNet-50", "42.0", "62.4", "44.2"],
    ["Ours",         "ResNet-50", "45.8", "66.3", "49.7"]
  ],
  "annotation": {
    "argument": "本表顯示本文方法在所有指標上均優於同等 backbone 的基線方法。",
    "reading_tip": "重點觀察 AP75（高 IoU 門檻），本文提升幅度最大，說明預測框更精準。"
  }
}
```

---

## 欄位填寫規則

### id
- 格式：`table_1`、`table_2`、`table_a` 等
- 必須與主 paper.json 中 `table_load` 和 `table_ref` 的 `table_id` 完全一致
- 命名按論文中表格編號：Table 3 → `table_3`

### caption_en / caption_zh
- `caption_en`：保留論文原始英文標題（含 "Table X:" 前綴）
- `caption_zh`：翻譯為中文

### headers
- 直接對應論文表格的欄位標題
- 複雜的雙層標題：用斜線合併，例如 `"AP / AR"`
- 保留原文英文，不翻譯

### rows
- 每個 row 為陣列，元素順序對應 headers
- 所有值均為**字串**，包括數字
- 缺值填 `"—"` 或 `"N/A"`
- 上標/特殊符號用 Unicode，例如 `"1.0 × 10²⁰"`
- 保持論文原始排列順序

### rows 索引說明

`table_ref` 用 0-based 索引引用此表：

| 索引 | 對應 |
|------|------|
| row=0, col=0 | rows[0][0]（第一行第一欄） |
| row=2, col=3 | rows[2][3]（第三行第四欄） |

**注意**：headers 不算在 row 索引內。row=0 是 rows 陣列的第一個元素（即表格第一行數據）。

### annotation
- `argument`（畫面標籤「這張表在說什麼」）：站在論文的立場，用 1–2 句白話說明此表想證明什麼
- `reading_tip`（畫面標籤「怎麼看」）：先看哪一欄／哪一列、關鍵對比點在哪；表中出現不常見的指標或縮寫時，
  順便用一句話解釋（例如「AP75＝框和答案重疊 75% 以上才算對，數字越高代表框越準」）
- 兩者都是導讀，不寫對實驗設計的批評（例如「缺少某某基線」「樣本太少」），除非論文自己在侷限討論中提到

---

## 常見表格類型

### 效能比較表（最常見）
```json
{
  "id": "table_1",
  "caption_en": "Table 1: Results on COCO val2017.",
  "caption_zh": "表 1：COCO val2017 結果。",
  "headers": ["Method", "Dataset", "Metric1", "Metric2"],
  "rows": [
    ["Baseline A", "COCO", "35.2", "56.8"],
    ["Baseline B", "COCO", "38.1", "59.3"],
    ["Ours",       "COCO", "42.5", "63.7"]
  ],
  "annotation": {
    "argument": "本方法在所有指標上超越基線。",
    "reading_tip": "注意 Metric2 的提升幅度（+4.4）為最大。"
  }
}
```

### 消融實驗表
```json
{
  "id": "table_3",
  "caption_en": "Table 3: Ablation study.",
  "caption_zh": "表 3：消融實驗。",
  "headers": ["Component", "Setting", "AP drop"],
  "rows": [
    ["Full model",        "✓ ✓ ✓", "—"],
    ["w/o cross-attn",   "✗ ✓ ✓", "-3.2"],
    ["w/o pos encoding", "✓ ✗ ✓", "-1.8"]
  ],
  "annotation": {
    "argument": "每個元件均對最終效能有正向貢獻。",
    "reading_tip": "cross-attn 移除後掉 3.2 AP，為最關鍵元件。"
  }
}
```

### 超參數敏感性表
```json
{
  "id": "table_4",
  "caption_en": "Table 4: Hyperparameter sensitivity.",
  "caption_zh": "表 4：超參數敏感性分析。",
  "headers": ["Learning Rate", "Batch Size", "Val Acc"],
  "rows": [
    ["1e-3", "32",  "82.1"],
    ["1e-4", "32",  "84.6"],
    ["1e-4", "128", "85.3"]
  ],
  "annotation": {
    "argument": "較小的學習率與較大的 batch size 組合效果最佳。",
    "reading_tip": "學習率從 1e-3 降至 1e-4 帶來最顯著提升（+2.5）。"
  }
}
```

---

## 執行流程

```
1. 閱讀論文 PDF，識別所有表格
2. 為每個表格產出對應的 table_X.json
3. 如使用者有提供 paper.json，確認 table_id 命名一致
4. 將所有 table_X.json 輸出至 /mnt/user-data/outputs/tables/ 供下載
5. 告知使用者將 tables/ 資料夾放到與 paper.json 同層目錄
```

## 輸出時附上的使用說明

```
📊 表格 JSON 已產出！

請將 tables/ 資料夾放到與 paper.json 同層目錄：
my-paper/
├── paper.json
├── generate.py
├── style.css
├── figures/
└── tables/          ← 把下載的 JSON 放這裡
    ├── table_1.json
    └── table_2.json

然後重新執行：python generate.py paper.json
```

## 品質檢核

- [ ] 每個 JSON 都有 `id`、`caption_en`、`caption_zh`、`headers`、`rows`、`annotation`
- [ ] `id` 與主 paper.json 的 `table_load` / `table_ref` 中的 `table_id` 一致
- [ ] `rows` 每行的元素數量與 `headers` 長度一致
- [ ] 所有值均為字串
- [ ] `annotation` 有 `argument` 和 `reading_tip`
