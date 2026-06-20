# 內容 JSON Schema

本文件定義 `create_docx.py` 接受的 JSON 輸入格式。

## 頂層結構

```json
{
  "product_name": "string (必填) - 產品名稱",
  "document_name": "string (必填) - 文件名稱",
  "version": "string (必填) - 版本號，如 V1.0.0",
  "date": "string (必填) - 日期，格式 YYYY/MM/DD",
  "author": "string (可選) - 作者",
  "company_name": "string (可選) - 頁首公司名稱，預設為範例公司並提示修改",
  "confidentiality": "string (可選) - 頁首機密等級，預設為內部使用並提示修改",
  "logo_path": "string (可選) - 公司 logo 圖檔路徑，未提供時使用 [公司 Logo] placeholder",
  "footer_text": "string (可選) - 頁尾補充文字，預設提示使用者修改",
  "revision_history": "array (可選) - 修訂紀錄",
  "sections": "array (必填) - 章節內容",
  "appendices": "array (可選) - 附錄"
}
```

## 頁首頁尾欄位

- `company_name`：頁首公司名稱；未提供時使用「範例股份有限公司（請修改）」。
- `confidentiality`：頁首機密等級；未提供時使用「內部使用（請修改）」。
- `logo_path`：公司 logo 圖檔路徑；未提供或檔案不存在時，頁首會顯示 `[公司 Logo]` 文字 placeholder。
- `footer_text`：頁尾補充資訊；未提供時使用「請修改頁尾資訊」。

## revision_history

修訂紀錄陣列，每筆包含：

```json
{
  "version": "string - 版本號",
  "date": "string - 日期",
  "author": "string - 作者",
  "summary": "string - 修改摘要"
}
```

## sections

章節陣列，每個章節包含：

```json
{
  "level": "integer (1-9) - 標題層級",
  "title": "string - 章節標題（不含編號，程式會自動產生）",
  "blocks": "array - 內容區塊陣列",
  "children": "array - 子章節陣列（結構同上）"
}
```

## blocks 區塊類型

### paragraph - 一般段落

```json
{
  "type": "paragraph",
  "text": "string - 段落文字",
  "indent": "boolean - 是否首行縮排（true = _內文-縮排，false = _內文-無縮排）"
}
```

### note - 註解

```json
{
  "type": "note",
  "text": "string - 註解文字，如「註 1：這是註解內容。」"
}
```

### list - 清單

```json
{
  "type": "list",
  "list_type": "string - 清單類型：letter / chinese / symbol / number",
  "items": [
    {
      "level": "integer (1-5) - 清單層級",
      "text": "string - 清單項目文字（不含前綴）",
      "note": "string (可選) - 清單內文，置於該項目下方"
    }
  ]
}
```

### code - 程式碼區塊

```json
{
  "type": "code",
  "numbered": "boolean - 是否顯示行號",
  "language": "string (可選) - 程式語言標記",
  "code": "string - 程式碼內容，使用 \\n 分隔各行"
}
```

### table - 表格

```json
{
  "type": "table",
  "title": "string - 表格名稱（不含編號前綴，程式自動產出「表 X-Y」）",
  "headers": ["string array - 表頭欄位"],
  "rows": [
    ["string array - 每列資料"]
  ]
}
```

### image - 圖片

```json
{
  "type": "image",
  "title": "string - 圖片名稱（不含編號前綴，程式自動產出「圖 X-Y」）",
  "path": "string (可選) - 圖片檔案路徑，若無則產生 placeholder",
  "width": "number (可選) - 圖片寬度（inches），預設 5.5"
}
```

## appendices

附錄陣列，每個附錄包含：

```json
{
  "label": "string - 附錄標籤，如 A、B、C",
  "title": "string - 附錄標題",
  "blocks": "array - 內容區塊（同 sections 的 blocks）"
}
```

## 完整範例

請參考 `scripts/example_input.json`。