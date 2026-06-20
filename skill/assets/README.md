# assets

此目錄保留給靜態資源，例如公司 logo、預設圖片 placeholder 或文件示意圖。

## 公司 Logo

建議將正式公司 logo 放在：

```text
assets/company-logo.png
```

並於 JSON 輸入中設定：

```json
{
  "logo_path": "assets/company-logo.png"
}
```

若未提供 `logo_path` 或檔案不存在，`create_docx.py` 會在頁首放入 `[公司 Logo]` placeholder，並搭配公司名稱、文件名稱、版本與機密等級範例文字，提醒使用者於正式文件中修改。
