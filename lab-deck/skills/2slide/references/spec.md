# deck.json 規格

`scripts/build_deck.py` 讀取的唯一輸入。完整範例：`assets/example-deck.json`。

## 目錄
1. 頂層欄位
2. 導覽列與章節
3. 內容頁（content）與三種密度
4. 其他頁型（mapping / split / refs / statement）
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
  - `conclusion_pos`：橫條位置。
    - `center`（預設）：不預留空間，內容照常排滿；方框蓋住標題以下整個內容區，橫條置中在方框中間。
    - `bottom`：在內容下方預留一條空間，橫條放在頁面底部，方框只蓋到橫條上緣。內容較少、底下本來就有空白時使用。
    - 可以在頂層 `style.conclusion_pos` 設定整份的預設，單頁再用 `conclusion_pos` 覆寫。
  - 三種密度、兩種導覽列都會顯示。文字太長時會縮小字級，縮到 14pt 仍放不下就發出警告。
  - 雙排風格用它來呈現問題；這些問題最後要在「研究動機」彙整成 `mapping` 對照頁。
  - 不能用動畫的場合（匯出 PDF、上傳 Google 簡報），建置時加 `--keyframes`：每個有結論的頁面拆成前後兩頁
    （沒有橫條／有遮罩與橫條），內容位置完全相同，翻頁時就像動畫。
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

## 5. 圖片與表格

- `figs` 最多 3 張並排，4 張排成 2×2；在文字欄旁邊時，2 張以內上下疊放。
- `path` 相對於 deck.json 所在資料夾，只支援 PNG / JPG。檔案不存在時會畫虛線佔位框並發出警告，所以可以先建置、之後再補圖。
- `table`：二維陣列，第一列是表頭。儲存格含「待填」會以橘色顯示，用來標示尚未填入的數值。
- 同一頁有 `table` 時，`figs` 只會在 `visual` 密度顯示（當作表格的圖像版）。

## 6. 概念色

頂層 `concepts` 指定論文裡需要區分的概念，每個概念一個槽位（1–4），整份簡報不換色：

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
python scripts/build_deck.py deck.json -o out.pptx --keyframes       # 不用動畫，結論橫條拆成前後兩頁
```

以 `⚠` 開頭的輸出是警告，常見的有：
- `text may overflow`：文字縮到最小字級仍放不下 → 刪減內容，或拆成兩頁。
- `conclusion bar text may overflow`：結論橫條文字太長 → 縮成一句話。
- `image missing`：找不到圖 → 補圖，或保留佔位並在回報中列出。
- `only N figures fit`：圖太多 → 拆頁。
