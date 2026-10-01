# CLAUDE.md

給接手開發的 agent 看的專案背景。使用者是實驗室研究生，平常用繁體中文溝通。

## 這個專案在做什麼

`readout`：一條「論文 → 標注檔＋簡報」的流程。repo 本身是一個 Claude Code plugin（`.claude-plugin/`），
底下有三個 skill（claude.ai 則是各自上傳 `dist/*.skill`）：
- `skills/2annotate`：產出 paper.json、tables/、generate.py、style.css，讓使用者在本機產生雙欄導讀網頁。這是使用者原本就寫好的 skill。
- `skills/2slide`：只做簡報；把內容寫成 deck.json，再由 `scripts/build_deck.py`（python-pptx）渲染成 .pptx。
- `skills/both`：分派用，兩者都要時先跑 2annotate、再從 paper.json 跑 2slide。

簡報是**初稿**，使用者會自己在 PowerPoint 裡再修改，所以**可編輯性比精緻度重要**：
文字放在文字方塊、架構圖用原生圖形繪製並組成群組、命名，不要把內容壓成圖片。

## 風格來源

使用者提供了兩份自己的簡報作為目標，兩種風格都會用到：
- **單排**：研究提案（Google 簡報）
- **雙排**：論文報告（PowerPoint，10 × 5.625 吋）

規範已整理在 `skills/2slide/references/style.md`，**以該檔為準**。原始簡報不在 repo 裡（避免外流論文截圖與研究內容）；
使用者如果放在 `private/`，可以拿來對照。

## 已定案的設計決策

- 三個 skill 各做各的：沒有特別說明就不替使用者多做另一邊；明確說兩者都要時才用 `both`。
- 導覽列（`single` / `double`）和密度（`text` / `balanced` / `visual`）是兩個獨立設定。
- 單排時，章節就是故事線；雙排時，章節用通用名稱（介紹／研究方法／實驗結果／結論），故事線放在子章節。
- **不加**章節分隔頁、大綱頁、Q&A 頁。
- 紅色、綠色只用於「問題／解法」；結論橫條是正向成果時用綠色（`solution`），不要用紅色。
  截圖上的紅框代表「重點」，不分正負。
- 概念色是**選用的**：只有界線明確、反覆出現、需要對照的概念才用，不要硬湊；指定後整份不換色。
- 同一份 deck.json 必須能用任意導覽列與密度的組合建置；每個內容頁都要寫 `full`、`short`、`one_line` 三種文字，
  並同時寫 `tag`（單排）與 `title`（雙排）。
- 數字、年份不能捏造：沒有來源的數值一律寫「— 待填 —」。
- 動畫：結論橫條點一下出現（半透明白底置中）；一頁最多 3 組標註（「出現 → 消失」算一組）。
  有標註的頁面最後一下要回到「所有標註＋統整橫幅」的統整畫面，橫幅放在截圖下方、不加白底。
- 漸進聚焦展開成多頁（不用動畫），一步亮一塊；聚焦頁不放紅框之類的標註（標註是單頁重點，聚焦是展開說明）。
- 架構圖是報告者自己整理的概念架構，一頁滿版；跨頁沿用時位置固定不動。
- 有產出簡報時，回報一定要附上色彩與概念對照表（建置時會印出，格式見 2slide 的 SKILL.md 第 6 節）。
- 渲染器留在 skill 資料夾內（claude.ai 的 skill 必須能獨立運作），`pyproject.toml` 以 `package-dir` 指向它。
- 2annotate 是使用者原本寫好的 skill，修改前先問過使用者。

## 功能狀態

v0.3.0 起，`style.md` 描述的呈現手法都已實作：雙排外觀、兩種封面、結論橫條（點擊出現、`--keyframes`）、
概念色與色彩對照表、截圖標註（出現／消失、統整畫面）、架構圖、時間軸、漸進聚焦。
每項功能在 `tests/test_build.py` 都有測試；新增或修改功能時，同步更新 `spec.md`、範例與測試。

## 開發方式

用 [uv](https://docs.astral.sh/uv/) 管理環境（`uv.lock` 要一起提交）：

```bash
uv sync                      # 建立 .venv，安裝套件（可編輯模式）與 pytest
uv run pytest -q
uv run readout-slide skills/2slide/assets/example-deck.json -o out/ex.pptx --all-densities
uv run scripts/package_skills.py   # 輸出 dist/*.skill
```

轉成圖片檢查：有 LibreOffice 時用
`soffice --headless --convert-to pdf --outdir out out/ex-balanced.pptx && pdftoppm -jpeg -r 60 out/ex-balanced.pdf out/p`；
Windows 上也可以用 PowerPoint COM（`Presentation.Slides(i).Export(path, "PNG", 1280, 720)`），
結論橫條、截圖標註的動畫可以用 `Slide.TimeLine.MainSequence` 檢查。

每次修改外觀，都要轉成圖片**實際看過**，並和 style.md 的規範對照後才算完成。改外觀時，同時檢查單排與雙排、三種密度，
不要只看其中一種。
