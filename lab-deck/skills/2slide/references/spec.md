# deck.json 規格

`scripts/build_deck.py` 讀取的唯一輸入。完整範例：`assets/example-deck.json`。

## 目錄
1. 頂層欄位
2. 導覽列與章節
3. 內容頁（content）與三種密度
4. 其他頁型（mapping / split / refs / statement）
5. 圖片與表格
6. 主題覆寫
7. 建置指令與警告

---

## 1. 頂層欄位

```json
{
  "style":    {"nav": "single", "density": "balanced"},
  "chapters": [{"name": "問題背景", "sections": ["資料限制", "既有方法"]}, "..."],
  "cover":    {"title": "…", "subtitle": "…", "byline": "報告人．Lab Meeting 論文報告", "date": "2026 / 10 / 02"},
  "theme":    {"colors": {}, "fonts": {}},
  "slides":   [ ... ]
}
```

- `style.nav`：`single`（一排章節分頁）或 `double`（章節＋子章節兩排）。指令列 `--nav` 可覆寫。
- `style.density`：`text` / `balanced` / `visual`，整份簡報的預設密度。`--density` 可覆寫，單頁也可用 `density` 覆寫。
- `chapters`：導覽列的分頁，順序即顯示順序。可寫字串或 `{name, sections}`；`double` 模式需要 `sections`。
- 封面不顯示導覽列。

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

## 6. 主題覆寫

預設色取自實驗室簡報：主色 `1085DE`、未選分頁 `B7DAF5`（單排）／`BFBFBF`（雙排）、分隔線 `FF9900`、問題 `E06666`、解法 `8CD96A`。
要換色只寫要改的鍵：

```json
"theme": {"colors": {"active": "2E7D32", "rule": "F9A825"}, "fonts": {"ea": "Noto Sans TC"}}
```

可改的顏色鍵：`active inactive inactive_double rule text dark problem solution neutral muted good_bg good_fg bad_bg bad_fg tint todo`。

## 7. 建置指令與警告

```bash
python scripts/build_deck.py deck.json -o out.pptx
python scripts/build_deck.py deck.json -o out.pptx --nav double --density visual
python scripts/build_deck.py deck.json -o out.pptx --all-densities   # 一次輸出三種密度
```

以 `⚠` 開頭的輸出是警告，常見的有：
- `text may overflow`：文字縮到最小字級仍放不下 → 刪減內容，或拆成兩頁。
- `image missing`：找不到圖 → 補圖，或保留佔位並在回報中列出。
- `only N figures fit`：圖太多 → 拆頁。
