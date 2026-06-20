# docx-template — 公司 Word 文件樣板產生器

一套可在各種 AI 工具（TRAE、Claude、ChatGPT、Copilot 等）中使用的 Word 文件自動化技能，支援從 Markdown / 純文字 / JSON 一鍵產出符合公司格式的 `.docx` 文件，並內建品質檢查與文件修正工具。

---

## 目錄

- [功能特色](#功能特色)
- [快速開始](#快速開始)
- [安裝](#安裝)
- [支援的 AI 工具](#支援的-ai-工具)
- [使用範例](#使用範例)
  - [範例 1：從對話描述產生文件](#範例-1從對話描述產生文件)
  - [範例 2：從 JSON 產生文件](#範例-2從-json-產生文件)
  - [範例 3：檢查文件品質](#範例-3檢查文件品質)
  - [範例 4：修正文件問題](#範例-4修正文件問題)
  - [範例 5：產出示範文件](#範例-5產出示範文件)
- [文件類型支援](#文件類型支援)
- [QA 檢查清單](#qa-檢查清單)
- [常見問題修正指南](#常見問題修正指南)
- [目錄結構](#目錄結構)
- [注意事項](#注意事項)

---

## 功能特色

- **多格式輸入**：支援 Markdown、純文字大綱、JSON 結構化內容、口語描述
- **完整文件結構**：自動產生封面、修訂紀錄、目錄，以及內容含表格／圖片時才產生的表目錄、圖目錄、正文、附錄
- **多層標題**：支援 Heading 1 到 Heading 9 自動編號
- **四種清單系統**：字母（A/B/C）、國字（一/二/三）、符號（●/○/■）、數字（1/2/3），各五層
- **程式碼區塊**：支援有編號與無編號兩種格式，含灰底與框線
- **表格與圖片**：自動加標號（表 X-X、圖 X-X），表格含灰底表頭
- **內建品質檢查**：`inspect_docx.py` 自動掃描樣式、結構、錯誤
- **跨 AI 工具**：可在 TRAE、Claude、ChatGPT、Copilot 等支援 Python 執行環境的工具中使用

---

## 快速開始

```bash
# 1. 安裝相依套件
pip install python-docx Pillow

# 2. 產出示範文件（零設定，立即體驗）
cd docx-template
python scripts/create_docx.py --demo --output demo.docx

# 3. 從 JSON 產生你自己的文件
python scripts/create_docx.py --input scripts/example_input.json --output my_doc.docx

# 4. 檢查文件品質
python scripts/inspect_docx.py my_doc.docx
```

---

## 安裝

### 環境需求

| 需求 | 版本 |
|------|------|
| Python | 3.8 以上 |
| python-docx | >= 1.0.0 |
| Pillow | >= 9.0.0（選用，圖片 placeholder） |

### 安裝步驟

```bash
# 複製或下載專案到本地
# 進入專案目錄
cd docx-template

# 安裝相依套件
pip install python-docx Pillow
```

---

## 支援的 AI 工具

本技能設計為**可攜式（portable）**，可在任何支援 Python 執行環境的 AI 工具中使用：

| AI 工具 | 使用方式 | 說明 |
|---------|---------|------|
| **TRAE** | 匯入 Skill | 將 `SKILL.md` 與 `scripts/` 放入 Skill 目錄中，TRAE 會自動載入 |
| **Claude Code** | 上傳專案 | 將整個 `docx-template/` 目錄上傳或放入工作區，直接對 Claude 下指令 |
| **ChatGPT (Code Interpreter)** | 上傳檔案 | 上傳 `scripts/` 與 `references/` 到 ChatGPT，用自然語言描述需求 |
| **GitHub Copilot** | 本地使用 | 在終端機中直接執行 Python 腳本 |
| **Cursor / Windsurf** | 工作區 | 將專案放入工作區，在對話中直接描述需求 |

### 在 AI 工具中的通用 Prompt 範本

當你將此專案載入 AI 工具後，可以用以下方式描述需求：

> 「幫我產生一份系統規格書，產品名稱是 XXX，版本 V1.0.0，包含系統架構和 API 設計兩章」

> 「檢查 output.docx 這個文件有沒有問題」

> 「output.docx 的標題顏色不對，幫我修正後重新產出」

---

## 使用範例

### 範例 1：從對話描述產生文件

**情境**：你需要一份系統規格書，但沒有任何結構化內容，只有腦中的想法。

**步驟**：直接在 AI 工具中描述你的需求，AI 會自動建構 JSON 並產出文件。

**輸入範例**（對 AI 說）：
```
幫我產生一份「訂單管理系統」的系統規格書，版本 V1.0.0，作者是張小明。
需要包含以下章節：
1. 文件目的
2. 系統範圍（包含使用者管理、訂單處理、報表功能）
3. 系統架構（前端 React、後端 Spring Boot）
4. API 設計（至少 5 個 endpoint）
5. 附錄：錯誤碼對照表
```

**輸出**：一份完整的 `.docx` 文件，包含封面、修訂紀錄、目錄、所有章節內容、附錄。

---

### 範例 2：從 JSON 產生文件

**情境**：你已經有結構化的 JSON 內容，想直接產出文件。

**步驟**：

1. 準備 JSON 檔案（參考 `scripts/example_input.json`）
2. 執行產生指令

```bash
python scripts/create_docx.py --input my_content.json --output 規格書.docx
```

**JSON 最小結構範例**：

```json
{
  "product_name": "我的產品",
  "document_name": "系統規格書",
  "version": "V1.0.0",
  "date": "2025/06/18",
  "author": "作者名",
  "revision_history": [
    {
      "version": "V1.0.0",
      "date": "2025/06/18",
      "author": "作者名",
      "summary": "初版建立"
    }
  ],
  "sections": [
    {
      "level": 1,
      "title": "簡介",
      "blocks": [
        {
          "type": "paragraph",
          "text": "這是文件內容。",
          "indent": true
        }
      ],
      "children": []
    }
  ],
  "appendices": []
}
```

---

### 範例 3：檢查文件品質

**情境**：你已經產出了一份 `.docx` 文件，需要確認是否符合公司格式規範。

**步驟**：

```bash
python scripts/inspect_docx.py 規格書.docx
```

**檢查結果範例**：

```
============================================================
📋 docx 品質檢查報告：規格書.docx
============================================================

✅ PASS  必要樣式檢查：Heading 1-9 存在
✅ PASS  內文樣式檢查：_內文-縮排 / _內文-無縮排 存在
✅ PASS  表格樣式檢查：表_標號、表格標題、表格內容 存在
✅ PASS  程式碼樣式檢查：程式碼-有編號 / 程式碼-無編號 存在
✅ PASS  錯誤參照檢查：無 Error: Reference source not found
✅ PASS  封面檢查：產品名稱、版本、日期 存在
✅ PASS  修訂紀錄檢查：修訂紀錄表存在
✅ PASS  標題檢查：至少有一個 Heading 1

============================================================
📊 結果摘要：8/8 PASS，0 WARN，0 FAIL
============================================================
```

**檢查項目說明**：

| 檢查項目 | 說明 | 結果類型 |
|---------|------|----------|
| 必要樣式 | Heading 1-9、內文、表格、程式碼樣式是否存在 | PASS / FAIL |
| 錯誤參照 | 是否有 `Error: Reference source not found` | PASS / FAIL |
| 封面 | 產品名稱、文件名稱、版本、日期是否存在 | PASS / WARN / FAIL |
| 修訂紀錄 | 修訂紀錄表是否存在 | PASS / FAIL |
| 標題 | 是否至少有一個 Heading 1 | PASS / FAIL |
| Normal 段落 | 是否有過多 Normal 段落（可能濫用） | PASS / WARN |

---

### 範例 4：修正文件問題

**情境**：檢查後發現文件有問題，需要修正。

#### 常見問題 1：標題顏色不對

**問題描述**：`inspect_docx.py` 回報標題或內文文字顏色不是 `#000000`。

**修正方式**：確認 JSON 中沒有手動指定顏色，讓 `create_docx.py` 自動套用樣式。重新產出：

```bash
# 修正 JSON 內容後重新產出
python scripts/create_docx.py --input fixed_content.json --output 規格書_fixed.docx

# 再次檢查
python scripts/inspect_docx.py 規格書_fixed.docx
```

#### 常見問題 2：表格爆版

**問題描述**：表格欄位太多，寬度超出頁面。

**修正方式**：
- 減少表格欄位數
- 合併部分欄位
- 調整 JSON 中的表格結構後重新產出

#### 常見問題 3：缺少修訂紀錄

**問題描述**：`inspect_docx.py` 回報 FAIL: 修訂紀錄表不存在。

**修正方式**：在 JSON 中確保 `revision_history` 陣列不為空（至少一筆紀錄）：

```json
"revision_history": [
  {
    "version": "V1.0.0",
    "date": "2025/06/18",
    "author": "作者名",
    "summary": "初版建立"
  }
]
```

#### 常見問題 4：目錄顯示「錯誤！找不到參考來源」

**問題描述**：在 Word 中開啟後，目錄顯示 `Error: Reference source not found`。

**修正方式**：這不是文件的問題，而是 Word 欄位尚未更新。在 Word 中全選（Ctrl+A），按 F9 更新所有欄位即可。

#### 一鍵修正流程

```bash
# 1. 檢查現有文件
python scripts/inspect_docx.py output.docx

# 2. 根據檢查結果修正 JSON 內容
#    （調整章節結構、修正樣式、補齊欄位等）

# 3. 重新產出
python scripts/create_docx.py --input fixed_input.json --output output_fixed.docx

# 4. 再次檢查確認
python scripts/inspect_docx.py output_fixed.docx
```

---

### 範例 5：產出示範文件

**情境**：你想先看看文件格式長什麼樣子，再決定是否使用。

```bash
python scripts/create_docx.py --demo --output demo.docx
```

這會產出一份包含所有樣式示範的完整文件，涵蓋：
- 封面（產品名稱、文件名稱、版本、日期）
- 修訂紀錄表
- 目錄（需在 Word 中按 F9 更新）
- 表目錄、圖目錄
- 各種標題層級（Heading 1-9）
- 四種清單系統（字母、國字、符號、數字）
- 程式碼區塊（有編號／無編號）
- 表格（含表標號）
- 圖片（含圖標號）
- 附錄（含程式碼與表格）

---

## 文件類型支援

本技能支援以下文件類型，每種類型都有建議的章節結構（詳見 `references/document-structure.md`）：

| 文件類型 | 典型章節 | 適用場景 |
|---------|---------|---------|
| 系統規格書 | 文件目的、系統範圍、架構、功能需求、API、資料庫、部署 | 軟體開發專案 |
| API 文件 | 認證方式、共通格式、Endpoint、錯誤碼、範例 | 前後端分離專案 |
| 操作手冊 | 操作前準備、操作流程、FAQ、異常排除 | 系統上線交付 |
| 驗收文件 | 驗收範圍、測試案例、驗收結果、問題追蹤 | 專案驗收階段 |
| 教育訓練文件 | 課程目的、先備知識、課程內容、實作練習、評量 | 內部教育訓練 |
| 專案文件 | 專案概述、組織、時程、需求、設計、測試、風險 | 專案管理 |

---

## QA 檢查清單

產出文件前後，建議依序檢查以下項目（詳見 `references/qa-checklist.md`）：

### 產出前檢查

- [ ] 產品名稱、文件名稱、版本、日期、作者是否齊全？
- [ ] 文件類型是否明確？
- [ ] 章節內容是否準備好？
- [ ] 修訂紀錄資料是否準備好？

### 產出中檢查

- [ ] 標題是否使用 Heading 1-9（非 Normal 模擬）？
- [ ] 內文是否使用 `_內文-縮排` 或 `_內文-無縮排`？
- [ ] 程式碼是否使用程式碼區塊（非一般段落）？
- [ ] 表格是否有表標號（`表_標號`）？
- [ ] 圖片是否有圖標號（`圖_標號`）？
- [ ] 附錄標題是否使用附錄格式？

### 產出後檢查

- [ ] 是否執行 `inspect_docx.py`？
- [ ] 是否沒有 `Error: Reference source not found`？
- [ ] 是否提醒使用者更新 Word 欄位（Ctrl+A, F9）？
- [ ] 表格沒有爆版？
- [ ] 標題字型為標楷體、內文字型為微軟正黑體？
- [ ] 標題與內文文字顏色為 #000000？

---

## 常見問題修正指南

| 問題 | 原因 | 修正方式 |
|------|------|---------|
| 目錄顯示「錯誤！找不到參考來源」 | Word 欄位未更新 | 在 Word 中全選（Ctrl+A）後按 F9 |
| 表格爆版 | 欄位過多或內容過長 | 減少欄位數，或縮短內容文字 |
| 字型不正確 | 系統未安裝微軟正黑體 | 安裝微軟正黑體，或在 Word 中內嵌字型 |
| 頁碼從封面開始 | section break 問題 | 在 Word 中手動設定分節符號 |
| 標題顏色不對 | 手動覆蓋了樣式顏色 | 重新產出，不要手動指定顏色 |
| 程式碼變成一般文字 | JSON 中未使用 code 類型 | 將 `type` 改為 `"code"` |
| 缺少修訂紀錄 | JSON 中 `revision_history` 為空 | 填入至少一筆修訂紀錄 |
| 清單層級不正確 | `level` 值設定錯誤 | 確認 level 從 1 開始，最多 5 層 |

---

## 目錄結構

```
docx-template/
├── SKILL.md                          # 技能定義檔（AI 工具讀取入口）
├── agents/
│   └── openai.yaml                   # OpenAI / ChatGPT 代理設定
├── assets/
│   └── README.md                     # 素材說明
├── references/
│   ├── style-guide.md                # 視覺風格指南（字型、顏色、頁面設定）
│   ├── document-structure.md         # 文件結構指南（各類型文件骨架）
│   ├── list-rules.md                 # 清單規則（四種清單系統 × 五層）
│   ├── content-json-schema.md        # JSON 內容結構定義
│   └── qa-checklist.md               # QA 檢查清單
├── scripts/
│   ├── create_docx.py                # 文件產生腳本
│   ├── inspect_docx.py               # 文件品質檢查腳本
│   ├── example_input.json            # JSON 輸入範例
│   └── README.md                     # 腳本說明
└── README.md                         # 本檔案
```

---

## 注意事項

1. **Word 欄位更新**：`python-docx` 產出的目錄（TOC）、表目錄、圖目錄使用 Word 欄位，需在 Word 中開啟後全選（Ctrl+A）按 F9 更新。

2. **羅馬頁碼**：因 `python-docx` 限制，前置頁（封面、修訂紀錄、目錄頁）目前不顯示頁碼，正文從阿拉伯數字 1 開始。如需羅馬頁碼，請在 Word 中手動設定。

3. **多層清單**：本工具使用手動前綴文字模擬 Word 清單結構，視覺效果一致但非 Word 原生清單功能。

4. **封面斜線裝飾**：原始設計包含斜線裝飾，因 `python-docx` 限制無法繪製。可在 Word 中手動加入。

5. **字型依賴**：文件使用「微軟正黑體」，請確保系統已安裝該字型。若在 macOS 或 Linux 上開啟，建議安裝對應字型或替換為系統預設字型。

6. **跨 AI 工具使用**：在不同 AI 工具中使用時，需確保 AI 工具具備 Python 執行能力（Code Interpreter 或終端機存取）。若 AI 工具不支援直接執行 Python，可將腳本下載到本地執行。