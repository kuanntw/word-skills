# Scripts

本目錄包含產生與檢查公司格式 Word 文件的 Python 腳本。

## 環境需求

- Python 3.8 以上
- 以下 Python 套件：

```bash
pip install python-docx Pillow olefile
```

| 套件 | 版本 | 用途 |
|------|------|------|
| `python-docx` | >= 1.0.0 | 建立與編輯 .docx 文件 |
| `olefile` | >= 0.46 | 選用，用於檢查 .doc 檔案 |
| `Pillow` | >= 9.0.0 | 選用，用於建立圖片 placeholder |

## 腳本說明

### create_docx.py

公司文件樣板產生器，將 JSON 結構化內容轉為符合公司格式的 .docx。

```bash
# 從 JSON 產出文件
python create_docx.py --input example_input.json --output output.docx

# 產出示範文件（包含所有樣式與結構）
python create_docx.py --demo --output demo.docx
```

參數說明：
- `--input`：輸入 JSON 檔案路徑
- `--output`：輸出 .docx 檔案路徑
- `--demo`：產出完整示範文件（不使用 --input）
- 若未指定 `--output`，預設為 `output.docx`

### inspect_docx.py

檢查產出的 Word 文件是否符合樣式規範（支援 .docx 與 .doc）。

```bash
python inspect_docx.py output.docx
python inspect_docx.py CM03A_V1.doc
```

檢查項目：
- `.docx`：必要樣式是否存在（Heading 1-9、內文樣式、表格樣式等）
- `.docx`：是否有 `Error: Reference source not found` 錯誤
- `.docx`：是否有過多 Normal 段落
- `.docx`：封面必要文字是否存在
- `.docx`：是否有修訂紀錄表
- `.docx`：是否至少有一個 Heading 1
- `.doc`：錯誤參照檢查（是否有 `Error: Reference source not found`）
- `.doc`：基本內容結構檢查

輸出格式：終端機列出 PASS / WARN / FAIL，若有問題列出問題段落或摘要。

## Word 欄位更新提醒

**重要**：`python-docx` 產出的目錄（TOC）、表目錄、圖目錄使用 Word 欄位（field code），在 python-docx 產出後為未更新狀態。

> 開啟 Word 後請全選（Ctrl+A），按 F9 更新目錄，以及內容含表格／圖片時才產生的表目錄、圖目錄與交互參照。

## 已知限制

1. **羅馬頁碼**：`python-docx` 對 section break 與頁碼格式的支援有限。目前前置頁（封面、修訂紀錄、目錄頁）不顯示頁碼，正文從阿拉伯數字 1 開始。若需要前置頁使用羅馬頁碼，請在 Word 中手動設定。

2. **多層清單**：`python-docx` 無法完整建立 Word 原生多層清單，本程式採用手動前綴文字模擬清單結構，視覺效果一致但非 Word 原生清單。

3. **TOC 欄位**：目錄，以及內容含表格／圖片時才產生的表目錄、圖目錄使用 Word field code 建立，需在 Word 中手動更新（F9）。

4. **封面斜線裝飾**：原始設計包含左上角與右下角淡灰斜線裝飾，因 python-docx 限制無法繪製，使用簡潔版封面取代。可在 Word 中手動加入裝飾。