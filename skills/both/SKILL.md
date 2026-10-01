---
name: both
description: >
  論文報告一條龍：丟進一篇論文，同時產出「標注檔」（2annotate 的雙欄導讀網頁，給自己讀）
  與「Lab Meeting 簡報」（2slide 的可編輯 .pptx，給別人聽），兩份共用同一份 paper.json 與圖片命名。
  只有在使用者**明確要兩者**時才使用此 skill，例如「批註（導讀）和簡報都要」「整理這篇並做成 lab meeting 簡報」
  「兩個都做」。沒有特別說明時不要替使用者多做另一邊：只要導讀網頁（或只丟論文）用 2annotate，
  只要簡報用 2slide。
---

# both

一篇論文進來，依序產出兩種東西：

| 產出 | 給誰 | 做法 |
|---|---|---|
| **標注檔** | 自己讀 | 依 **2annotate** skill 的 SKILL.md 產出 paper.json、tables/、generate.py、style.css |
| **簡報** | 給別人聽 | 依 **2slide** skill 的 SKILL.md 寫 deck.json，渲染成 .pptx |

兩個 skill 以 plugin 安裝時就在同一層（`../2annotate/`、`../2slide/`）；在 claude.ai 上則是各自上傳的 skill。
找不到其中一個時，告訴使用者需要安裝哪一個，先完成另一個。

## 1. 決定要產出什麼

只有使用者明確說兩者都要時才兩者都產。沒有特別說明，就只做使用者提到的那一邊：

| 使用者的說法 | 產出 |
|---|---|
| 「批註（導讀）和簡報都要」「整理這篇並做成簡報」「兩個都做」 | 兩者都產 |
| 只丟論文，或「整理這篇我自己看」「只要批註／導讀」 | 只產標注檔（照 2annotate 做即可） |
| 「做簡報」「直接做投影片」 | 只產簡報（照 2slide 做即可） |

簡報的導覽列、密度、故事線照 2slide 的預設。

## 2. 先標注，再做簡報

1. 完整依 2annotate 的 SKILL.md 產出標注檔。
2. 簡報從 **paper.json** 取材，不要重讀 PDF：兩份產出的內容、術語、圖表編號才會一致
   （對照方式見 2slide 的 `references/narratives.md`）。
3. 兩份產出共用同一個 `figures/` 資料夾與命名方式（`figures/figure_1.png`），使用者截一次圖兩邊都能用。

## 3. 回報

交付順序：標注檔 → .pptx → deck.json。簡報部分的回報照 2slide 第 6 節（一定要附色彩與概念對照表）。
建議使用者先看標注檔把論文讀懂、把截圖放進 `figures/`，再重新建置一次簡報。
