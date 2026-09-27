# readout

丟進一篇論文，產出兩樣東西：

| 產出 | 用途 | Skill |
|---|---|---|
| **標注檔** | 自己讀：雙欄英中對照、技術高亮、逐段批註的網頁 | `2annotate` |
| **簡報** | 給別人聽：實驗室風格的可編輯 .pptx | `2slide` |
| **兩者都要** | 先標注、再從同一份 paper.json 做簡報 | `both` |

每個 skill 只做自己的產出；明確說兩者都要時才由 `both` 一起產。有產出簡報時，會一併附上「顏色 → 概念」對照表，方便之後手動修改時沿用同一套配色。

## 簡報風格

兩種導覽列 × 三種內容密度，可以任意搭配：

- **單排**：章節本身就是故事線，頁面左上用「/ 彩色標籤」標示（紅＝問題、綠＝解法、藍＝中性），常用左文右圖。
- **雙排**：章節用通用名稱，故事線放在子章節（子章節嵌在橘線上），頁面用黑字標題，問題以紅色結論橫條呈現。
- **密度**：`text`（段落＋本頁重點）／`balanced`（條列＋圖）／`visual`（一句話＋滿版圖）。

完整規範見 [`skills/2slide/references/style.md`](skills/2slide/references/style.md)。

## 安裝

**claude.ai**：到 [Releases](../../releases) 下載 `2slide.skill`、`2annotate.skill`、`both.skill`，在「設定 → Capabilities → Skills」上傳。

**Claude Code**：整個 repo 是一個 plugin，安裝後三個 skill 分別是 `/readout:2slide`、`/readout:2annotate`、`/readout:both`。

```bash
claude plugin marketplace add k7term1a/readout
claude plugin install readout@readout
```

在 Claude Code 對話中也可以用 `/plugin marketplace add k7term1a/readout`、`/plugin install readout@readout`。

**只要渲染器（命令列）**：

```bash
uv tool install git+https://github.com/k7term1a/readout
readout-slide deck.json -o talk.pptx --nav double --density visual
readout-slide deck.json -o talk.pptx --all-densities   # 一次輸出三種密度
readout-slide deck.json -o talk.pptx --keyframes       # 不用動畫，每一下點擊拆成一頁（匯出 PDF 用）
```

## 使用方式

在 Claude 裡上傳論文 PDF，直接說想要什麼：

- 只丟論文，或「整理這篇我自己看」→ 只產標注檔（`2annotate`）
- 「幫我做這篇的 lab meeting 簡報，用單排、平衡型」→ 只產簡報（`2slide`）
- 「批註和簡報都要」→ 標注檔＋簡報（`both`）

建議的流程是先產標注檔，把論文讀懂，並把截圖放進 `figures/figure_N.png`，再產簡報。兩份產出共用同一套圖片命名。

## 開發

用 [uv](https://docs.astral.sh/uv/) 管理環境：

```bash
uv sync                            # 建立 .venv，安裝套件與 pytest
uv run pytest -q                   # 所有範例 × 兩種導覽列 × 三種密度都要能建置，且沒有版面警告
uv run scripts/package_skills.py   # 輸出 dist/*.skill
```

推送 `v*` tag 時，CI 會自動把兩個 skill 打包並發佈到 Releases。

```
.claude-plugin/          plugin.json、marketplace.json（Claude Code plugin：readout）
skills/
  2slide/                SKILL.md、references/（style、narratives、spec）、scripts/build_deck.py、assets/
  2annotate/             SKILL.md、scripts/generate.py、assets/style.css、references/
  both/                  SKILL.md（分派：先 2annotate、再 2slide）
examples/                示範用的 deck.json（圖片一律使用佔位框）
tests/                   建置測試
```

渲染器放在 skill 資料夾裡，讓 skill 在 claude.ai 上也能獨立運作；`pyproject.toml` 直接指向同一份程式，不需要另外維護一份。

## 不要提交的東西

參考用的簡報、論文 PDF、論文截圖都不要放進 repo。需要在本機參考時，放在 `private/`（已列入 `.gitignore`）。
