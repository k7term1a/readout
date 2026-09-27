---
name: ai-paper-annotator
description: >
  AI 學術論文雙欄批註器。當使用者上傳 AI / ML / NLP 論文 PDF，或要求「摘取論文」
  「製作論文批註」「產生論文 JSON」「整理論文」「論文摘要 HTML」時，務必使用此 skill。
  本 skill 一次產出 paper.json + tables/table_X.json + generate.py + style.css，
  讓使用者在本機執行 generate.py 產生雙欄對照、技術高亮、表格 sticky 的批註網頁。
  即使使用者只說「幫我整理這篇論文」「幫我看看這篇 paper」「翻譯這篇論文」也應觸發此 skill。
---

# AI Paper Annotator Skill

## 概覽

本 skill 將 AI 學術論文轉換為結構化 JSON，搭配 `generate.py`
在使用者本機產生雙欄批註 HTML，包含：

- 英中雙語對照 + 技術高亮（左欄）
- 逐段技術批註（右欄）
- 圖表佔位卡片（使用者將截圖放入 `figures/`）
- 表格完整呈現，Sticky 固定於段落上方，捲動時自動 highlight 被提及的列
- 技術演化路徑圖（Evolutionary Path）或魚骨圖（Fishbone Diagram）

## 完整系統流程

本 skill 一次產出 `paper.json`（主文批註）與 `tables/table_X.json`（表格數據）。
**圖片**由使用者自行截圖處理。

```
                   使用者上傳論文 PDF
                         │
            ┌────────────┼────────────┐
            ▼            ▼            ▼
     ┌──────────┐  ┌───────────┐  ┌──────────────┐
     │ 本 skill  │  │ 本 skill  │  │ 圖片：使用者  │
     │ 產出      │  │ 同時產出   │  │ 自行截圖      │
     │ paper.json│  │ tables/   │  │ 放入          │
     │ +generate │  │ table_X   │  │ figures/      │
     │ +style.css│  │ .json     │  └───────────────┘
     └────┬──────┘  └────┬──────┘        │
          │              │               │
          ▼              ▼               ▼
     ┌──────────────────────────────────────┐
     │ 使用者本機資料夾 my-paper/            │
     │  ├ paper.json                        │
     │  ├ generate.py                       │
     │  ├ style.css                         │
     │  ├ figures/figure_1.png ...          │
     │  └ tables/table_1.json ...           │
     │                                      │
     │  $ python generate.py paper.json     │
     │  → paper.html                        │
     └──────────────────────────────────────┘
```

### 圖片 vs 表格的處理差異

| 資源 | 來源 | 處理方式 | paper.json 中的表示 |
|------|------|---------|---------------------|
| 圖片 | 使用者從論文 PDF 截圖 | 手動存為 `figures/figure_X.png` | `figure` block（含 annotation） |
| 表格 | 本 skill 自動解析產出 | 自動產出 `tables/table_X.json` | `table_load` block（佔位）+ `table_ref` block（索引） |

## 執行流程

本 skill 安裝後，所有相關檔案都在同一個 skill 資料夾內。Claude 應依序：

```
1. 讀取本 skill 內的 references/paper_schema.json
   → 取得 paper.json 完整 JSON Schema 定義
2. 讀取本 skill 內的 references/table_extraction.md
   → 取得表格擷取的詳細規則與範例
3. 閱讀使用者上傳的論文 PDF
4. 依照本文件 + paper_schema.json 產生 paper.json
5. 依照 table_extraction.md 為論文中每個表格產出 tables/table_X.json
6. 將以下檔案複製到 /mnt/user-data/outputs/ 供使用者下載：
   a. paper.json                              ← 本次產出
   b. tables/table_1.json, table_2.json, ...  ← 本次產出（每個表格一個檔案）
   c. generate.py    ← 從本 skill 的 scripts/generate.py 複製
   d. style.css      ← 從本 skill 的 assets/style.css 複製
7. 用 present_files 工具呈現所有產出檔案
8. 告知使用者本機後續步驟（見下方）
```

> **檔案路徑提示**：本 skill 安裝後位於 `/mnt/skills/user/ai-paper-annotator/`，
> 內含：
> - `SKILL.md`（本文件）
> - `references/paper_schema.json`
> - `references/table_extraction.md`
> - `scripts/generate.py`
> - `assets/style.css`

## 輸出時附上的使用說明

完成所有 JSON 後，向使用者說明：

```
📁 請將下載的檔案放在同一資料夾：

my-paper/
├── paper.json          ← 主批註檔
├── generate.py         ← HTML 產生器
├── style.css           ← 樣式表
├── figures/            ← 從論文 PDF 截圖，命名對應 figure id
│   ├── figure_1.png
│   └── figure_2.jpg
└── tables/             ← 已自動產出，直接放入即可
    ├── table_1.json
    └── table_2.json

🖼️ 圖片處理：
   從論文 PDF 截圖，命名為 figure_1.png / figure_2.jpg 等，放入 figures/ 資料夾。

📊 表格已自動產出：
   tables/ 資料夾中的 JSON 檔案已一併下載，直接放入即可。
   若需重新擷取某個表格，可在新對話中要求 Claude「重新擷取 Table X」即可。

▶️ 產生 HTML：
   python generate.py paper.json

🌐 開啟 paper.html 即可瀏覽。
```

---

## Step 1：讀取論文

完整閱讀 PDF，理解：
- 研究問題與動機
- 提出的方法與架構
- 實驗設計與結果
- 與前人方法的比較
- 侷限與未來方向

---

## Step 2：產生 paper.json

嚴格依照 `references/paper_schema.json` 的結構輸出。

### 2-1. `meta` 欄位

```json
{
  "meta": {
    "title": "論文完整標題",
    "authors": ["Author One", "Author Two"],
    "venue": "NeurIPS 2024",
    "year": 2024,
    "abstract_zh": "完整摘要中譯，流暢中文學術語氣。"
  }
}
```

### 2-2. `sections` 欄位

每個 section 對應一個章節，`blocks` 陣列按論文順序排列，包含以下類型：

#### `paragraph` block

```json
{
  "block_type": "paragraph",
  "en": [
    "The proposed model uses ",
    {"text": "cross-attention", "type": "concept"},
    " to align visual and textual representations."
  ],
  "zh": "所提模型採用交叉注意力機制，對齊視覺與文字表徵。",
  "annotation": {
    "role": "模型元件描述",
    "logic": "交叉注意力允許不同模態間的資訊流動，相較於早期特徵拼接更精細。",
    "note": "注意此處未說明注意力頭數，實作時需參閱附錄。"
  }
}
```

**高亮類型：**
| type | 顏色 | 使用時機 |
|------|------|---------|
| `concept` | 紫 | 核心概念、模型名稱、重要術語 |
| `problem` | 紅 | 現有方法缺陷、研究缺口 |
| `solution` | 綠 | 本文提出的解法、技術貢獻 |
| `method` | 黃 | 方法細節、損失函數、超參數 |
| `result` | 藍 | 數值結果、benchmark 名稱 |

**annotation.role** 必須從以下選項擇一：
`問題定義` `研究動機` `文獻回顧` `方法提出` `模型元件描述`
`實作細節` `實驗設計` `結果分析` `消融實驗` `侷限討論` `結論`

**zh 翻譯原則：**
- 意譯為主，不逐字直譯
- 專有術語首次出現可加括號保留英文
- 高亮標記規則與 `en` 相同，使用 `{"text": "...", "type": "..."}` 格式標註對應的中文術語
- `zh` 可以是純字串或含高亮的陣列，格式與 `en` 一致

#### `figure` block

論文出現 Figure X 時插入（**放在提及該圖的段落之前**）。

```json
{
  "block_type": "figure",
  "id": "figure_1",
  "type": "模型架構圖",
  "caption_en": "Figure 1: Overview of the proposed architecture.",
  "caption_zh": "圖 1：所提架構總覽。",
  "aspect_ratio": "square",
  "annotation": {
    "argument": "此圖說明整體系統的資料流向與模組分工。",
    "reading_tip": "注意 Encoder 與 Decoder 之間的跳躍連接。",
    "data_flow": "影像 → CNN → Cross-Attention → Decoder → 輸出"
  }
}
```

**aspect_ratio 判斷：**
- `wide`：橫向比例（比較圖、訓練曲線）→ 預設，全欄寬
- `square`：接近 1:1（架構圖、混淆矩陣）→ 向右浮動
- `tall`：縱向（流程圖、樹狀結構）→ 置中，固定高度

**type 選項：**
`模型架構圖` `消融實驗表` `效能比較圖` `損失曲線`
`潛在空間視覺化` `資料流程圖` `注意力熱圖` `訓練流程圖` `其他`

**圖片命名規則（輸出時告知使用者）：**
`figure_1.png`、`figure_2.jpg` 放入 `figures/` 資料夾

#### `table_load` block

論文出現 Table X 時，在討論該表格的段落**之前**插入佔位。
對應的表格數據存在 `tables/table_X.json`（由本 skill 同時產出）。

```json
{
  "block_type": "table_load",
  "table_id": "table_1"
}
```

`table_id` 必須與同時產出的 `tables/table_1.json` 檔名及其內部 `id` 欄位一致。

#### `table_ref` block

段落討論特定表格數值時，緊接在對應 `paragraph` block 之**後**插入。
只列出被提及的儲存格，不重複整個表格。

```json
{
  "block_type": "table_ref",
  "table_id": "table_1",
  "cells": [
    {"row": 2, "col": 2, "label": "Our AP"},
    {"row": 2, "col": 3, "label": "Our AP50"}
  ],
  "note": "在 COCO val 上超越 DETR 3.8 AP。"
}
```

**重要：**
- `row`/`col` 為 **0-based**，對應 `table_X.json` 的 `rows` 陣列索引。
- 此 block 在 `tables/table_X.json` 不存在時仍可正常產生 HTML（會退化為連結）。
- `table_ref` 的 row/col 必須與同時產出的 `table_X.json` 的 rows 陣列對應一致。

---

### 2-3. `summary` 欄位

```json
{
  "summary": {
    "problem_gap": "現有方法缺口與研究動機（2–4 句）",
    "contributions": ["貢獻一", "貢獻二"],
    "one_line": "方法一句話摘要，含類比說明",
    "experiment_conclusion": "實驗結論（2–4 句）",
    "applications": ["應用場景一", "延伸方向二"],
    "evolutionary_path": { ... },
    "fishbone": { ... }
  }
}
```

#### evolutionary_path

技術演化有向圖。本論文節點設 `"is_this_paper": true`。

```json
{
  "nodes": [
    {"id": "prev", "label": "Prev Method", "year": 2020, "note": "一行說明"},
    {"id": "ours", "label": "Our Method",  "year": 2024, "note": "本文貢獻", "is_this_paper": true}
  ],
  "edges": [
    {"from": "prev", "to": "ours", "label": "改進"}
  ]
}
```

#### fishbone

問題根因分析。`effect` 為主要問題（魚頭），`causes` 為各骨幹類別。

```json
{
  "effect": "現有方法的核心問題",
  "causes": [
    {"category": "資料問題", "items": ["標註稀缺", "分布偏移"]},
    {"category": "模型限制", "items": ["感受野受限", "計算複雜度高"]}
  ]
}
```

**兩者擇一或同時填入均可。**

---

## Step 3：產出表格 JSON

依照 `references/table_extraction.md` 的規則，為論文中每個表格產出對應的 `table_X.json`。

要點：
- 每個表格一個獨立 JSON 檔案
- `id` 必須與 paper.json 中 `table_load` / `table_ref` 的 `table_id` 一致
- `rows` 陣列的索引必須與 `table_ref` 的 `row`/`col` 對應
- 詳細格式與範例見 `references/table_extraction.md`

---

## Step 4：輸出檔案

將以下檔案複製到 `/mnt/user-data/outputs/` 供使用者下載：

具體操作：
1. 將產出的 paper.json 寫入 `/mnt/user-data/outputs/paper.json`
2. 將產出的 table JSON 寫入 `/mnt/user-data/outputs/tables/table_X.json`
3. 複製本 skill 的 `scripts/generate.py` → `/mnt/user-data/outputs/generate.py`
4. 複製本 skill 的 `assets/style.css` → `/mnt/user-data/outputs/style.css`
5. 用 `present_files` 工具呈現所有檔案，paper.json 放第一個

---

## 品質檢核

產生 JSON 前確認：
- [ ] 每個 block 都有 `block_type` 欄位
- [ ] `figure` block 放在提及該圖的段落**之前**
- [ ] `table_load` block 放在討論該表格的段落**之前**
- [ ] `table_ref` block 緊接在對應 `paragraph` 之**後**
- [ ] `table_ref` 的 row/col 為 0-based，且與對應 table JSON 的 rows 陣列一致
- [ ] 每個 `table_load` 都有對應的 `tables/table_X.json` 產出
- [ ] `annotation.role` 使用規定清單中的選項
- [ ] `zh` 高亮標記與 `en` 對應
- [ ] `evolutionary_path` 的 edges from/to 對應實際 node id
- [ ] 輸出時包含 generate.py、style.css 和 tables/ 資料夾
