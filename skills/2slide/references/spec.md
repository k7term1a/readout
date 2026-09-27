# deck.json 規格

`scripts/build_deck.py` 讀取的唯一輸入。完整範例：`assets/example-deck.json`。

## 目錄
1. 頂層欄位
2. 導覽列與章節
3. 內容頁（content）與三種密度
4. 其他頁型（mapping / split / refs / statement / diagram / timeline）
5. 圖片與表格
6. 概念色
7. 主題覆寫
8. 建置指令與警告

---

## 1. 頂層欄位

```json
{
  "style":    {"nav": "single", "density": "balanced", "conclusion_pos": "center"},
  "chapters": [{"name": "問題背景", "sections": ["資料限制", "既有方法"]}, "..."],
  "concepts": {"圖像層級": 1, "區域層級": 2},
  "cover":    {"title": "…", "subtitle": "…", "byline": "報告人．Lab Meeting 論文報告", "date": "2026 / 10 / 02",
               "venue": "IJCAI 2018", "paper_title": "…", "authors": ["…"], "affiliations": ["…"], "presenter": "〇〇〇"},
  "theme":    {"colors": {}, "fonts": {}},
  "slides":   [ ... ]
}
```

- `style.nav`：`single`（一排章節分頁）或 `double`（章節＋子章節兩排）。指令列 `--nav` 可覆寫。
- `style.density`：`text` / `balanced` / `visual`，整份簡報的預設密度。`--density` 可覆寫，單頁也可用 `density` 覆寫。
- `chapters`：導覽列的分頁，順序即顯示順序。可寫字串或 `{name, sections}`；`double` 模式需要 `sections`。
- `concepts`：概念名稱 → 概念色槽位（1–4），見第 6 節。
- 封面不顯示導覽列，見下方「封面」。

### 封面

`cover.style` 決定封面樣式；沒寫就依導覽列決定（`single` → `centered`、`double` → `paper`）。
兩種樣式用到的欄位都寫上，切換導覽列時封面就不必重寫。

| 樣式 | 版面 | 使用欄位 |
|---|---|---|
| `centered` | 標題置中 → 橘線 → `subtitle`、`byline`（作者．用途）、`date` | `title` `subtitle` `byline` `date` |
| `paper` | 靠左：會議名稱藍色徽章 → 襯線粗體論文標題 → 橘線 → 作者、機構分兩欄 → 左下日期、右下「報告者：〇〇〇」 | `venue` `paper_title` `authors` `affiliations` `presenter` `date` |

- `paper` 樣式沒有 `paper_title` 時用 `title`；沒有 `authors` 時，左欄改放 `byline`。
- 作者、機構使用 `Times New Roman`；中文仍為微軟正黑體。
- 報告者姓名不知道時寫「〇〇〇」，不要自己編。

## 2. 導覽列與章節

- 每頁的 `chapter` 決定哪個分頁反白；`double` 模式下 `section` 決定第二排哪一項反白。
- `double` 模式：未選分頁為灰色；子章節是嵌在橘線上的小膠囊（目前子章節藍底、其他灰底），靠左排列但不貼齊左端：左邊留一小段橘線，右邊的橘線延伸到頁面右側。
  該章節沒有 `sections` 時，橘線不斷開。
- `double` 模式在右下角顯示頁碼（封面除外），使用 PowerPoint 的頁碼欄位，調整頁序後會自動更新。`single` 模式沒有頁碼。
- 沒有章節分隔頁（`divider` 頁型已移除；舊檔中的 `divider` 會被略過並發出警告）。
- 分頁數建議 4–6 個；超過會自動縮窄。多篇論文導讀時，把每篇論文的簡稱當成一個 chapter 即可。

## 3. 內容頁（content）與三種密度

```json
{
  "type": "content",
  "chapter": "問題背景", "section": "資料限制",
  "tag": "一、只有快照，沒有軌跡", "kind": "problem",
  "title": "網格人口快照", "subtitle": "電信資料只提供各網格、各時段的人數",
  "full":     ["完整段落一。", "完整段落二。"],
  "short":    ["條列一", "條列二"],
  "one_line": "一句話重點",
  "figs":     [{"id": "Figure 1", "caption": "網格人口快照示意", "path": "figures/figure_1.png"}],
  "table":    null,
  "notes":    ["講者備註（選填）"],
  "conclusion": "頁面底部的結論橫條（選填）",
  "conclusion_kind": "problem",
  "conclusion_pos": "center",
  "callout":  "頁底小字，例如資料來源（選填）",
  "density":  "visual",
  "text_ratio": 0.45
}
```

- `tag`：`single` 模式左上角的「/ 標籤」。`kind` 決定顏色：`problem`（紅）、`solution`（綠）、`neutral`（藍）。
  問題與解法請用相同編號（一、二、三）讓聽眾對得起來。
- `title`：`double` 模式左上角的粗體黑字標題；`subtitle` 是標題下方一行灰色說明（選填）。
  以圖為主的頁面可以省略 `title`，該頁就不放標題。
- **`tag` 和 `title` 都寫**，同一份 deck.json 才能切換導覽列：`single` 只用 `tag`，`double` 只用 `title`。
  `mapping`、`split`、`refs` 也一樣。
- 三個文字欄位對應三種密度，**盡量三個都寫**，切換密度時才不用重寫：

| 密度 | 使用欄位 | 版面 | 講者備註 |
|---|---|---|---|
| `text` | `full`（段落） | 文字欄約 56%，下方「本頁重點」框放 `one_line` | 只放 `notes` |
| `balanced` | `short`（條列） | 文字欄約 38%，圖在右 | `notes`，沒有就自動放 `full` |
| `visual` | `one_line` | 一句話＋滿版圖；沒有圖就變成置中大字的宣言頁 | `notes`，沒有就自動放 `full` |

- `conclusion`：全寬白字的結論橫條，寫這一頁要點出的結論（一句話，約 30 字以內）。
  - **點一下才出現**：標題以下的內容被半透明白色方框（「結論遮罩」，不透明度 70%）蓋住淡化，同時淡入橫條（「結論橫條」），
    兩者都是 0.5 秒淡入。報告時先講內容，講完點一下再帶出結論。
  - `conclusion_kind`：橫條顏色，`problem`（紅，預設）、`solution`（綠）、`neutral`（藍）。
    紅色只給問題；正向的成果、結論要寫 `solution`，中性的整理寫 `neutral`。
  - `conclusion_pos`：橫條位置。
    - `center`（預設）：不預留空間，內容照常排滿；方框蓋住標題以下整個內容區，橫條置中在方框中間。
    - `bottom`：在內容下方預留一條空間，橫條放在頁面底部，方框只蓋到橫條上緣。內容較少、底下本來就有空白時使用。
    - `below`：在內容下方預留空間，橫條放在底部，**不加白色方框**。有截圖標註（`marks`）的頁面預設用這個，
      讓統整畫面裡的紅框、色塊保持清楚。
    - 可以在頂層 `style.conclusion_pos` 設定整份的預設，單頁再用 `conclusion_pos` 覆寫（明寫的值永遠優先）。
  - 三種密度、兩種導覽列都會顯示。文字太長時會縮小字級，縮到 14pt 仍放不下就發出警告。
  - 雙排風格用它來呈現問題；這些問題最後要在「研究動機」彙整成 `mapping` 對照頁。
  - 不能用動畫的場合（匯出 PDF、上傳 Google 簡報），建置時加 `--keyframes`：每個有結論的頁面拆成前後兩頁
    （沒有橫條／有遮罩與橫條），內容位置完全相同，翻頁時就像動畫。截圖標註的 `reveal` 也一樣，每一下拆成一頁。
- 缺少的欄位會依序遞補（例如沒寫 `short` 就用 `full`）。
- `==文字==` 會變成藍色粗體強調。
- `text_ratio`：單頁覆寫文字欄比例（0–1）。

## 4. 其他頁型

**mapping**：問題 → 解法對照頁（紅色標籤 → 綠色標籤）。最多約 4 組。
```json
{"type": "mapping", "chapter": "方法提案", "tag": "問題與解法對照",
 "pairs": [["只有快照，沒有軌跡", "從快照反推流量"], ["…", "…"]]}
```

**split**：左綠右紅的對照頁，常用於「可以借用的 vs 限制與疑問」、「應用 vs 限制」。`visual` 密度只顯示標題。
```json
{"type": "split", "chapter": "討論啟發", "tag": "對我研究的意義",
 "good_title": "可以借用的", "bad_title": "限制與疑問",
 "good": [["標題", "一句說明"]], "bad": [["標題", "一句說明"]]}
```

**refs**：參考文獻頁，每筆是 `{title, source, gist}`；`visual` 密度不顯示 `gist`。

**statement**：置中大字（例如一句結論）。`{"type": "statement", "text": "…", "chapter": "選填"}`

**diagram**：報告者自己整理的概念架構圖，一頁滿版（不受密度影響；`full`／`short` 會放進講者備註）。
架構寫在頂層 `diagrams`，各頁用名稱引用，同一張圖可以跨頁沿用：

```json
"diagrams": {
  "vlm": {
    "direction": "right",
    "nodes": [
      {"id": "img",  "label": "輸入影像", "col": 1, "row": 1, "image": "figures/input.png"},
      {"id": "txt",  "label": "輸入文字", "col": 1, "row": 2},
      {"id": "venc", "label": "Image Encoder", "col": 2, "row": 1, "concept": "圖像層級"},
      {"id": "tenc", "label": "Text Encoder",  "col": 2, "row": 2},
      {"id": "fuse", "label": "Fusion", "col": 3, "row": [1, 2]}
    ],
    "edges": [["img", "venc"], ["txt", "tenc"], ["venc", "fuse"], ["tenc", "fuse", "拼接"]],
    "modules": [{"id": "enc", "label": "Encoder", "nodes": ["venc", "tenc"]}]
  }
},
"slides": [
  {"type": "diagram", "chapter": "介紹", "section": "過往研究", "tag": "…", "title": "過往研究：晚期融合",
   "diagram": "vlm", "conclusion": "兩個模態只在最後才融合，互動不足"},
  {"type": "diagram", "chapter": "研究方法", "section": "核心概念", "tag": "…", "title": "本文方法",
   "diagram": "vlm", "highlight": ["fuse"], "relabel": {"fuse": "Cross-Attention Fusion"}}
]
```

- **節點**：`id`、`label`、`col`、`row`（從 1 開始；`[1, 2]` 表示跨兩列／兩欄）。方塊大小依頁面自動計算。
  `concept` 套用概念色；`image` 在方塊上半部放縮圖（找不到檔案時畫虛線佔位並警告）。`label` 可以用 `{{概念}}`。
- **連線**：`[from, to]` 或 `[from, to, 標籤]`，也可以寫成 `{"from", "to", "label", "dashed": true}`。
  一律是 PowerPoint 連接線、黏在方塊上：對齊時是直線，否則是直角折線（`direction: "right"` 先水平、
  `"down"` 先垂直）；轉折點自動避開模組框。在 PowerPoint 裡拖動方塊，箭頭會跟著走。
- **模組框**：`modules: [{"id", "label", "nodes"}]`，虛線圓角框住這些節點，左上角放標題。
- **每頁的變化**（位置永遠依整張圖計算，所以跨頁不會移動）：
  - `highlight`：節點或模組的 id，加綠色描邊（`6AA84F`）與淺綠底（`D9EAD3`），表示新增或有變化。
  - `hide`：先藏起還沒介紹的節點，連到它們的線也一起藏起。
  - `relabel`：`{"id": "新文字"}`，換掉某個節點的文字。
- 可以加 `conclusion`（點一下出現），`--keyframes` 同樣適用。
- 整張圖組成群組「架構圖」；節點命名 `架構圖/<id>`，連線 `架構圖/連線/<from>-<to>`，模組框 `架構圖/模組/<id>`。
- 欄列太多、方塊會太小時會警告，請拆成兩頁或減少節點。
- 色彩對照表會多一列「綠色描邊」，列出本份簡報標綠的元件。

**timeline**：領域演進的時間軸，一頁滿版（不受密度影響）。底部是年份軸，每個分類一條水平色帶，模型用小膠囊標在年份上。

```json
{"type": "timeline", "chapter": "介紹", "section": "過往研究", "tag": "…", "title": "視覺語言模型的演進",
 "lanes": [{"name": "圖像層級", "concept": "圖像層級"}, {"name": "區域層級", "concept": "區域層級"}],
 "items": [
   {"label": "CLIP", "year": 2021, "lane": "圖像層級", "note": "ICML"},
   {"label": "GLIP", "year": 2022, "lane": "區域層級"},
   {"label": "本文", "year": 2023, "lane": "區域層級", "highlight": true}
 ]}
```

- `year` **只用整數年份**；起訖依 `items` 自動決定，或用 `years: [2018, 2024]` 指定。
- `lanes`：由上往下排的分類帶。有 `concept` 時用該概念色；沒有就用淺灰（`EFEFEF`），**不要為了時間軸另外新增概念**。
  不寫 `lanes` 就只有一條不分類的色帶。
- 同一年同一類有多個項目時自動上下錯開；色帶高度依內容決定，太擠時會警告（拆成兩頁或拉長年份範圍）。
- `note`：膠囊下方一行小字（例如會議名稱）；`highlight`：綠色描邊，用來標出本文或要強調的項目。
- 年份一律要有來源（論文的 related work 或參考文獻），不確定的不要放。
- 不畫項目之間的關係箭頭；可以加 `conclusion`。整張圖組成群組「時間軸」，膠囊命名 `時間軸/<label>`。

### 漸進聚焦（`focus`）

`content`（有截圖的頁）與 `diagram` 頁可以加 `focus`，建置時展開成多頁：先一頁全貌，再每一步亮一塊、其餘淡化
（填色與線條透明度約 70%、文字 `C0C0C0`、截圖被半透明白框蓋住）。位置完全不動。

```json
{"type": "diagram", "diagram": "vlm", "title": "方法元件",
 "focus": [
   {"on": "enc",  "subtitle": "先把兩個模態各自編碼"},
   {"on": "fuse", "subtitle": "再用跨模態注意力融合", "conclusion": "兩邊在每一層都有互動", "conclusion_kind": "solution"}
 ]}
```

- `on`：**一步只亮一塊**。架構圖寫節點或模組框的 id（模組會連同框內節點一起亮；連到亮起節點的連線也保留）；
  截圖寫 `figs[].regions` 的區塊 id，或整張圖的 `id`（同頁有多張圖時只亮那一張）。
- 截圖先用 `regions` 劃出區塊，寫法和截圖標註的位置一樣（`x` `y` `w` `h` 或 `row`／`col`）：
  `"regions": [{"id": "a", "x": 0, "y": 0, "w": 0.33, "h": 1}, …]`。區塊外的標註會一起淡化。
- 順序：截圖照**論文描述這張圖的順序**；架構圖照講述順序。
- 每一步可以覆寫 `title` `subtitle` `tag` `one_line` `short` `full` `notes` `conclusion` `conclusion_kind` `conclusion_pos`。
  頁面本身的 `conclusion` 只出現在全貌頁，不會複製到每一步。
- `"focus_overview": false` 可以省略全貌頁。
- 每一步的說明建議寫在 `subtitle`：雙排顯示在標題下方；單排（只有「/ 標籤」）在聚焦頁會顯示在標籤右側。
- 聚焦頁不畫截圖標註（`marks`）：標註是單頁重點，聚焦是展開說明，兩者分開放。

## 5. 圖片與表格

- `figs` 最多 3 張並排，4 張排成 2×2；在文字欄旁邊時，2 張以內上下疊放。
- `path` 相對於 deck.json 所在資料夾，只支援 PNG / JPG。檔案不存在時會畫虛線佔位框並發出警告，所以可以先建置、之後再補圖。
- `table`：二維陣列，第一列是表頭。儲存格含「待填」會以橘色顯示，用來標示尚未填入的數值。
- 同一頁有 `table` 時，`figs` 只會在 `visual` 密度顯示（當作表格的圖像版）。

### 截圖標註（`figs[].marks`）

在截圖（或佔位框）上疊加原生圖形，之後可以在 PowerPoint 裡直接拖曳、修改。

```json
"figs": [{
  "id": "Table 2", "path": "figures/table_2.png",
  "reveal": "click",
  "marks": [
    {"type": "band", "row": [2, 5], "rows": 8, "concept": "圖像層級", "text": "相同參數量"},
    {"type": "box",  "row": 8, "rows": 8, "col": 3, "cols": 4},
    {"type": "note", "x": 0.86, "y": 0.94, "text": "用執行效率換來的", "with_previous": true}
  ]
}]
```

| type | 畫出來的樣子 | 欄位 |
|---|---|---|
| `box` | 紅框（`FF0000`，2pt，不填色），圈出重點數值 | 位置；`color` 可改色 |
| `band` | 半透明概念色塊（不透明度 45%），標出一組列 | 位置（沒寫 `x`／`w` 時橫跨整張圖）；`concept`；`text`；`label_side` |
| `note` | 一句說明＋箭頭，箭頭指到 (`x`, `y`) | `x` `y`；`text`；`side`；`color`（箭頭）、`text_color` |

**位置**：一律是相對於圖片的 0–1 比例（左上角是 0, 0），圖片換了解析度也不會跑位。
- `x` `y` `w` `h`：直接給比例。
- `row` + `rows`、`col` + `cols`：把圖片均分成 `rows` 列（`cols` 欄），`row` 寫第幾列（從 1 開始，含表頭），
  或 `[起, 迄]`。適合列高平均的表格截圖；列高不平均時改用 `y` `h`。
- 座標是看截圖估的，會有誤差，使用者可能需要在 PowerPoint 裡微調。

**說明文字的位置**（`note` 的 `side`、`band` 的 `label_side`）：
- `right`（預設）／`left`：放在截圖旁的欄位，截圖會縮小讓出空間；同一側有多個說明時自動往下排開。
- `top`／`bottom`：放在截圖上方或下方（下方會放在圖說之後）。
- `inside`：疊在圖上。`note` 用 `tx` `ty`（0–1）指定文字位置；`band` 的文字放在色塊內左側。
- 說明文字寫在截圖旁邊，不要蓋住數值；只有圖上有空白處時才用 `inside`。

**點一下出現**：圖加上 `"reveal": "click"`，每個標註依 `marks` 的順序各自點一下淡入；標註加 `"with_previous": true`
就和上一個一起出現。標註加 `"exit": true` 會在**下一次點擊時淡出**（和下一個標註的出現同一下）；
最後一個標註不會消失。

**統整畫面**：講完之後，**所有淡出過的標註會在最後一下一起回來**，讓頁面停在「全部重點都在」的狀態。
有 `conclusion` 時就和統整橫幅同一下出現（橫幅在截圖下方、不加白色方框）；沒有 `conclusion` 時自己多一下。
`--keyframes` 的最後一頁就是這個統整畫面。

**一頁最多 3 組標註**：每個標註算一組，`with_previous` 併入前一組，「出現 → 消失」也只算一組。超過會警告。
漸進聚焦（`focus`）的頁面不會畫標註（會警告），要標重點請另開一頁。

**可編輯性**：每個元件命名為 `標註/<圖號>/紅框1`、`色塊2`、`色塊2說明`、`說明3`、`說明3箭頭`。
沒有動畫時，截圖和所有標註組成一個群組 `標註/<圖號>`，可以整組移動；有 `reveal` 時不組群組
（PowerPoint 無法對群組內的個別元件設定動畫）。

截圖寬度小於 3.5 吋時會警告：標註在文字型密度下常常太小，改用 `balanced`／`visual`，或調小 `text_ratio`。

## 6. 概念色

頂層 `concepts` 指定論文裡需要區分的概念，每個概念一個槽位（1–4），整份簡報不換色。
**這是選用欄位**：論文沒有界線明確、需要反覆對照的概念時就不要寫（判斷標準見 `style.md` 第 4 節）。

| 槽位 | 底色 | 描邊／文字 |
|---|---|---|
| 1（藍） | `B6CFF5` | `3C78D8` |
| 2（黃） | `FFECB3` | `BF9000` |
| 3（紫） | `D9D2E9` | `674EA7` |
| 4（青灰） | `D0E0E3` | `45818E` |

標記方式：

- **文字**：`{{圖像層級}}`，或 `{{圖像層級|顯示的文字}}`（用別的字，但算同一個概念）。渲染成該概念的文字色粗體。
  所有文字欄位都可以用（`title`、`subtitle`、`full`、`short`、`one_line`、表格、`split` 等）；
  放在膠囊標籤或結論橫條裡時只顯示文字、不上色，但仍會記入色彩對照。
- **圖片**：`figs[].concept: "圖像層級"`。截圖外加一圈概念色描邊；佔位框改用概念色底與描邊。
- **表格列**：某一列的第一格以 `{{概念}}` 開頭，整列套用該概念的底色，例如 `["{{圖像層級}} CLIP", "76.2"]`。

- 同一個槽位不能給兩個概念；槽位超出 1–4、或用了沒宣告的概念，都會發出警告並以一般文字呈現。
- 紅色、綠色保留給問題／解法，不能當概念色（槽位本來就不包含這兩色）。

**色彩對照**：每次建置後會印出一張 Markdown 表，列出這份簡報實際用到的紅（問題）、綠（解法）與各個概念色，
附色碼、代表的內容與出現的頁碼，可以直接貼進回報。問題與解法的名稱優先取自 `mapping` 對照頁。

```
| 顏色 | 色碼 | 代表 | 出現在 |
|---|---|---|---|
| 紅 | `E06666` | 問題：只有快照，沒有軌跡／格間流動看不見 | 第 2–5 頁 |
| 概念色 1（藍） | `B6CFF5`／`3C78D8` | 出發機率 | 第 7 頁 |
```

## 7. 主題覆寫

預設色取自實驗室簡報：主色 `1085DE`、未選分頁 `B7DAF5`（單排）／`BFBFBF`（雙排）、分隔線 `FF9900`、問題 `E06666`、解法 `8CD96A`。
要換色只寫要改的鍵：

```json
"theme": {"colors": {"active": "2E7D32", "rule": "F9A825"}, "fonts": {"ea": "Noto Sans TC"}}
```

可改的顏色鍵：`active inactive inactive_double rule text dark problem solution neutral muted good_bg good_fg bad_bg bad_fg tint todo`，以及概念色 `concept1_bg`、`concept1_fg` … `concept4_fg`。

## 8. 建置指令與警告

```bash
python scripts/build_deck.py deck.json -o out.pptx
python scripts/build_deck.py deck.json -o out.pptx --nav double --density visual
python scripts/build_deck.py deck.json -o out.pptx --all-densities   # 一次輸出三種密度
python scripts/build_deck.py deck.json -o out.pptx --keyframes       # 不用動畫，每一下點擊拆成一頁
```

以 `⚠` 開頭的輸出是警告，常見的有：
- `text may overflow`：文字縮到最小字級仍放不下 → 刪減內容，或拆成兩頁。
- `conclusion bar text may overflow`：結論橫條文字太長 → 縮成一句話。
- `annotated … is only …in wide`：有標註的截圖太小 → 換密度或調 `text_ratio`。
- `mark N: …`：標註的類型、座標或 `side` 寫錯。
- `N annotations … on one slide`：標註超過 3 組 → 拆頁。
- `marks … are not drawn on progressive-focus pages`：聚焦頁上的標註被略過 → 另開一頁放標註。
- `focus '…' matches no …`：`focus` 的 `on` 在這頁找不到對應的節點、模組或截圖區塊。
- `image missing`：找不到圖 → 補圖，或保留佔位並在回報中列出。
- `only N figures fit`：圖太多 → 拆頁。
