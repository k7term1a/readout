# CLAUDE.md

給接手開發的 agent 看的專案背景。使用者是實驗室研究生，平常用繁體中文溝通。

## 這個專案在做什麼

一條「論文 → 標注檔＋簡報」的流程，由兩個 skill 組成：
- `skills/2annotate`：產出 paper.json、tables/、generate.py、style.css，讓使用者在本機產生雙欄批註網頁。這是使用者原本就寫好的 skill。
- `skills/2slide`：統一分派流程；把內容寫成 deck.json，再由 `scripts/build_deck.py`（python-pptx）渲染成 .pptx。

簡報是**初稿**，使用者會自己在 PowerPoint 裡再修改，所以**可編輯性比精緻度重要**：
文字放在文字方塊、架構圖用原生圖形繪製並組成群組、命名，不要把內容壓成圖片。

## 風格來源

使用者提供了兩份自己的簡報作為目標，兩種風格都會用到：
- **單排**：研究提案（Google 簡報）
- **雙排**：論文報告（PowerPoint，10 × 5.625 吋）

規範已整理在 `skills/2slide/references/style.md`，**以該檔為準**。原始簡報不在 repo 裡（避免外流論文截圖與研究內容）；
使用者如果放在 `private/`，可以拿來對照。

## 已定案的設計決策

- 導覽列（`single` / `double`）和密度（`text` / `balanced` / `visual`）是兩個獨立設定。
- 單排時，章節就是故事線；雙排時，章節用通用名稱（介紹／研究方法／實驗結果／結論），故事線放在子章節。
- **不加**章節分隔頁、大綱頁、Q&A 頁。
- 紅色、綠色只用於「問題／解法」；概念色（藍 → 黃 → 紫 → 青灰）依每份簡報的內容指定，指定後整份不換色。
- 同一份 deck.json 必須能用任意導覽列與密度的組合建置；每個內容頁都要寫 `full`、`short`、`one_line` 三種文字。
- 數字不能捏造：沒有來源的數值一律寫「— 待填 —」。
- 有產出簡報時，回報一定要附上色彩與概念對照表（格式見 2slide 的 SKILL.md 第 6 節）。
- 渲染器留在 skill 資料夾內（claude.ai 的 skill 必須能獨立運作），`pyproject.toml` 以 `package-dir` 指向它。

## 待辦（依序進行，每完成一項就更新 SKILL.md 的「實作狀態」段落）

1. **雙排外觀對齊**：未選分頁改為灰色 `BFBFBF`；子章節改成嵌在橘線中間的小膠囊（藍底／灰底白字，橘線在左右兩側斷開）；
   雙排模式改用黑字標題（`title` 欄位，可省略），取代「/ 彩色標籤」；右下角加頁碼（封面除外）；
   **移除雙排模式自動產生的 divider 頁**（`divider` 頁型直接拿掉，同時更新 spec.md 與範例）。
2. **兩種封面**：`cover.style = "centered"`（單排：標題＋橘線＋作者．用途）／`"paper"`（雙排：會議名稱徽章＋襯線標題＋
   橘線＋作者與機構兩欄＋右下角報告者）。預設依導覽列決定。
3. **紅色結論橫條**：內容頁新增 `conclusion` 欄位，渲染成頁面底部的全寬紅底白字橫條。
4. **概念色**：頂層新增 `concepts: {"圖像層級": 1, "區域層級": 2}`；內容可以用 `{{圖像層級}}` 標記文字或區塊，
   渲染時套用對應的底色或描邊。建置時輸出「色彩對照」清單（顏色、色碼、概念、出現的頁碼），
   供 agent 直接貼進回報。
5. **截圖標註**：`figs[].marks`，每個標記可以是 `box`（紅框）、`band`（半透明概念色塊＋文字）、`note`（箭頭＋說明）；
   座標用相對於圖片的 0–1 比例，圖片換了也不會跑位。
6. **原生架構圖**：新增 `diagram` 頁型，以節點與連線描述（`nodes` / `edges`，節點有 `id`、`label`、`col`、`row`），
   自動排版；`highlight` 用綠色描邊標出新元件；整張圖組成一個群組，每個節點命名為 `架構圖/<id>`。
7. **時間軸**：新增 `timeline` 頁型，橫軸為年份，縱向用概念色帶分類。
8. **漸進聚焦**：新增 `focus_sequence`：指定一個 base 頁（content、diagram 或帶標註的圖）與依序聚焦的元素，
   展開成多頁；每頁只有聚焦的元素保持原色，其餘淡化（透明度約 70%，文字改為 `C0C0C0`）。
9. **收尾**：更新 `spec.md`、範例與測試（每項功能至少一個測試）；刪除 SKILL.md 的「實作狀態」段落；打 tag 發佈。

另外兩件小事：
- `skills/2annotate/SKILL.md` 裡的路徑寫成 `/mnt/skills/user/ai-paper-annotator/`，實際安裝後的路徑會依環境而不同，
  改成相對於 skill 資料夾的寫法。
- 2annotate 的描述在「上傳論文」時也會觸發，和 2slide 重疊。可以在它的描述最後加一句
  「若使用者也要簡報，改由 2slide 主導」；這是使用者的 skill，**修改前先問過使用者**。

## 開發方式

```bash
pip install -e ".[dev]"
pytest -q
python skills/2slide/scripts/build_deck.py skills/2slide/assets/example-deck.json -o out/ex.pptx --all-densities
soffice --headless --convert-to pdf --outdir out out/ex-balanced.pptx && pdftoppm -jpeg -r 60 out/ex-balanced.pdf out/p
```

每做完一項，都要轉成圖片**實際看過**，並和 style.md 的規範對照後才算完成。改外觀時，同時檢查單排與雙排、三種密度，
不要只看其中一種。
