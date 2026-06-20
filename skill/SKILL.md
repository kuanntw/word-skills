---
name: docx-template
description: >
  當使用者要求產生公司格式 Word 文件、將 Markdown / 純文字 / JSON 轉成 .docx、
  產出技術文件、系統規格書、API 文件、操作手冊、驗收文件、教育訓練文件、專案文件，
  或要求套用「公司文件樣板」「文件樣版 V1.1.0」「正式文件格式」「規格書格式」「報告書格式」時，
  觸發此 Skill。
---

# docx-template Skill

## 何時使用此 Skill

本 Skill 用於以下情境：

1. 使用者要求建立公司格式 Word 文件（.docx）。
2. 使用者要求將 Markdown / 純文字 / JSON 結構化內容轉成 Word。
3. 使用者要求產出技術文件、系統規格書、API 文件、操作手冊、驗收文件、維運文件、教育訓練文件、專案文件。
4. 使用者要求套用「公司文件樣板」、「文件樣版 V1.1.0」、「正式文件格式」、「規格書格式」、「報告書格式」。
5. 使用者要求自動產出 docx，並包含封面、修訂紀錄、目錄，以及內容含表格／圖片時才產生的表目錄、圖目錄、章節、表格、程式碼、附錄。

## 輸入來源判斷

本 Skill 支援多種輸入格式，依情境自動判斷處理方式：

| 輸入類型 | 處理方式 |
|---------|---------|
| Markdown | 先解析章節結構，轉為 JSON，再呼叫 `create_docx.py --input` |
| 純文字大綱 | 先整理成章節結構，再轉為 JSON，呼叫 `create_docx.py --input` |
| JSON | 直接呼叫 `create_docx.py --input` |
| 使用者口語描述 | 先以 LLM 理解章節需求，建構 JSON，再呼叫 `create_docx.py --input` |
| 無內容 / 僅要求格式 | 呼叫 `create_docx.py --demo` 產生示範文件 |

## 產出流程

1. **確認基本資訊**：產品名稱、文件名稱、版本、日期、作者、文件類型。
2. **決定文件結構**：依據文件類型選擇骨架，參考 `references/document-structure.md`。
3. **建構 JSON 內容**：將內容結構化為符合 `references/content-json-schema.md` 的 JSON。
4. **產出 .docx**：執行 `python scripts/create_docx.py --input <json_path> --output <output_path>`。
5. **品質檢查**：執行 `python scripts/inspect_docx.py <output_path>`。
6. **修正問題**：若有 FAIL 或 WARN，修正後重新產出。
7. **通知使用者**：提供產出檔案路徑、檢查結果，並提醒更新 Word 欄位。

## 樣式套用規則

- 正式章節標題只能使用 **Heading 1 到 Heading 9**，不可用 Normal 直接模擬標題，不可手動輸入編號後再套 Normal。
- 一般段落使用 **_內文-縮排**（首行有縮排）或 **_內文-無縮排**（無縮排）。
- 註解使用 **註解** 樣式。
- 程式碼片段不得用一般段落，必須使用 **程式碼-有編號** 或 **程式碼-無編號** 區塊。
- 表格必須加上表標號（**表_標號** 樣式），置於表格上方。
- 圖片必須加上圖標號（**圖_標號** 樣式），置於圖片下方。
- 清單使用四種清單系統（字母、國字、符號、數字），每種支援五層。
- 附錄標題使用 **附錄** 格式。

詳細樣式規格請參考 `references/style-guide.md`。

## 參考文件使用時機

| 參考文件 | 使用時機 |
|---------|---------|
| `references/style-guide.md` | 需要確認樣式細節、顏色、字型、頁面設定時 |
| `references/list-rules.md` | 需要決定清單類型、層級與清單內文格式時 |
| `references/document-structure.md` | 需要決定文件骨架（規格書、API 文件、操作手冊等）時 |
| `references/content-json-schema.md` | 建構 JSON 輸入時參考欄位結構 |
| `references/qa-checklist.md` | 產出前、產出中、產出後進行 QA 檢查時 |

## 腳本使用說明

### create_docx.py

```bash
# 從 JSON 產出文件
python scripts/create_docx.py --input example_input.json --output output.docx

# 產出示範文件
python scripts/create_docx.py --demo --output demo.docx
```

### inspect_docx.py

```bash
# 檢查產出的 .docx 是否符合規範
python scripts/inspect_docx.py output.docx

# 檢查舊版 .doc 檔案的基本內容（需安裝 olefile）
python scripts/inspect_docx.py CM03A_V1.doc
```

`.docx` 支援完整檢查（樣式、段落、封面、標題、表格、頁首頁尾等）。
`.doc` 僅支援基本內容檢查（錯誤參照、修訂紀錄、內容結構），樣式檢查需先將 `.doc` 轉為 `.docx`。

## 產出後 QA 規則

產出 .docx 後，必須執行以下檢查（參考 `references/qa-checklist.md`）：

1. 執行 `inspect_docx.py` 檢查樣式完整性。
2. 確認文件中**沒有** `Error: Reference source not found`。
3. 確認標題字型為標楷體、內文字型為微軟正黑體。
4. 確認標題與內文文字顏色為 #000000。
5. 確認表格沒有爆版。
6. 確認頁首頁尾正常。

## Word 欄位更新提醒

**重要**：`python-docx` 產出的目錄（TOC）、表目錄、圖目錄與交互參照為 Word 欄位，需要在 Microsoft Word 或 LibreOffice 中開啟後手動更新：

> **開啟 Word 後請全選（Ctrl+A），按 F9 更新目錄，以及內容含表格／圖片時才產生的表目錄、圖目錄與交互參照。**

## 特定文件類型指南

### API 文件
- 使用表格呈現 endpoint、method、request、response、錯誤碼。
- 程式碼範例使用「程式碼-有編號」區塊。

### 系統設計書
- 應包含架構、模組、流程、資料表、API、權限、部署、維運。
- 架構圖可使用圖片 placeholder，後續手動替換。

### 操作手冊
- 使用數字清單表示操作步驟。
- 每個步驟下方可用「清單內文」補充說明。
- 使用程式碼區塊呈現命令列指令。

## 相依性

Python 套件需求（詳見 `scripts/README.md`）：
- `python-docx` >= 1.0.0
- `olefile` >= 0.46（選用，用於檢查 .doc 檔案）
- `Pillow`（選用，用於圖片 placeholder）

```bash
pip install python-docx Pillow olefile
```