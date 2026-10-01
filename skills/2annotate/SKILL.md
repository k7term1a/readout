---
name: 2annotate
description: >
  AI 學術論文雙欄導讀器。當使用者上傳 AI / ML / NLP 論文 PDF，或要求「摘取論文」
  「製作論文批註」「製作論文導讀」「論文導讀」「產生論文 JSON」「整理論文」「論文摘要 HTML」時，務必使用此 skill。
  本 skill 一次產出 paper.json + tables/table_X.json + generate.py + style.css，
  讓使用者在本機執行 generate.py 產生雙欄對照、技術高亮、表格 sticky 的導讀網頁
  （右欄站在論文的立場，幫讀者讀懂每一段）。
  即使使用者只說「幫我整理這篇論文」「幫我看看這篇 paper」「翻譯這篇論文」也應觸發此 skill。
---

# 2annotate：論文雙欄導讀

## 概覽

本 skill 將 AI 學術論文轉換為結構化 JSON，搭配 `generate.py`
在使用者本機產生雙欄導讀 HTML，包含：

- 英中雙語對照 + 技術高亮（左欄）
- 逐段導讀（右欄）：這段在說什麼 → 讀懂需要的背景 → 跟前後文的關係
- 圖片可以點一下放大
- 圖表佔位卡片（使用者將截圖放入 `figures/`）
- 表格完整呈現，Sticky 固定於段落上方，捲動時自動 highlight 被提及的列
- 技術演化路徑圖（Evolutionary Path）或魚骨圖（Fishbone Diagram）

## 完整系統流程

本 skill 一次產出 `paper.json`（主文與導讀）與 `tables/table_X.json`（表格數據）。
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

> **檔案路徑提示**：以下路徑都相對於本 skill 的資料夾（即本 SKILL.md 所在的資料夾；
> 安裝位置依環境而不同，例如 claude.ai 在 `/mnt/skills/user/2annotate/`，Claude Code plugin 則在 plugin 的 `skills/2annotate/`）。
> 輸出位置 `/mnt/user-data/outputs/` 是 claude.ai 的下載資料夾；在其他環境（例如 Claude Code）改寫到使用者目前的工作資料夾。
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
├── paper.json          ← 主文與導讀
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
    "logic": "...",
    "note": "..."
  }
}
```

#### 右欄是「導讀」，不是「批註」

右欄站在**論文的立場**，幫讀者讀懂這一段。順序是：這段在說什麼 → 讀懂需要的背景 → 跟前後文的關係。
批評意見不是預設內容。

| 欄位 | 畫面標籤 | 要寫什麼 |
|---|---|---|
| `logic` | 這段在說什麼 | 用白話重述這段的主張，**優先用具體例子說明**。不要評論作者的寫作策略（例如「開場在替方法鋪路」）。 |
| `note` | 背景補充 | 補讀懂這段需要的知識：段落裡的術語、前人模型（例如 DDPM、STGCN 是什麼）、公式在算什麼，以及這段和前後文的關係。每個補充的術語都要有一句白話解釋或例子。 |

**禁止**：除非 `role` 是 `侷限討論`，`note` 和 `logic` 不得出現對論文的評論，例如「沒有驗證」「用詞誇大」
「審稿時會被抓」「有點勉強」「應該要做某某實驗」。這類內容整段刪掉，不要換個說法保留。
`侷限討論` 段落也以整理作者自己承認的侷限為主。

**範例**（交叉注意力那一段，以都市 OD 流量預測為例）：

```json
"annotation": {
  "role": "模型元件描述",
  "logic": "模型在預測每個區域的流量時，會去「查」周邊的 POI（興趣點）資訊，再決定要參考多少。例如兩個區域過去的進出人數幾乎一樣，但一個周邊都是辦公大樓、另一個都是住宅，早上八點的流向就會相反；交叉注意力讓模型能用 POI 把這兩個區域分開。",
  "note": "交叉注意力（cross-attention）：Q（query，問題）來自流量特徵，代表「這個區域現在要找什麼資訊」；K（key，索引）和 V（value，內容）來自 POI 特徵，K 用來比對相關程度，V 是真正被取出的內容。算出來的權重越高，代表那類 POI 對這個區域越重要。這一段接續上一段的「流量編碼器」，下一段會說明融合後的特徵怎麼送進擴散模型去噪。"
}
```

對照：**不要**寫成「交叉注意力比特徵拼接更精細」（沒說它在做什麼）或「此處未說明注意力頭數」（這是批評，不是導讀）。

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
    "reading_tip": "從左往右看：先看影像如何被 CNN 轉成特徵，再看中間的 Cross-Attention 如何結合兩種輸入。Encoder 與 Decoder 之間的跳躍連接（skip connection）把淺層的細節直接送到後面，避免細節在壓縮過程中遺失。",
    "data_flow": "影像 → CNN → Cross-Attention → Decoder → 輸出"
  }
}
```

圖與表的 annotation 同樣是導讀：`argument`（畫面標籤「這張圖在說什麼」）站在論文的立場重述這張圖要表達的事；
`reading_tip`（畫面標籤「怎麼看」）告訴讀者先看哪裡、符號或座標軸代表什麼，遇到不常見的元件或指標順便用一句話解釋。

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
- [ ] `logic` 是在重述論文說了什麼，不是在評論作者
- [ ] `note` 裡每個被補充的術語都有一句白話解釋或舉例
- [ ] 非「侷限討論」的段落，`note` 沒有批評性字眼（沒有驗證、誇大、勉強、審稿、應該要做…）
- [ ] `zh` 高亮標記與 `en` 對應
- [ ] `evolutionary_path` 的 edges from/to 對應實際 node id
- [ ] 輸出時包含 generate.py、style.css 和 tables/ 資料夾
