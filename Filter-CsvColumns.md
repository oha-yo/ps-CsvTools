
---

## 📄 Filter-CsvColumns.md

```markdown
# 🧹 Filter-CsvColumns.ps1

指定した項目を除外したCSVファイルを生成するフィルタリングスクリプトです。個人情報の除去や特定列のマスキングなどに活用できます。

---

## 🧰 主な機能

- 除外対象列の指定（番号またはヘッダー名）
- 完全一致 / 部分一致モード
- 出力ファイルの自動命名
- ログ出力対応（任意）

---

## 🚀 使用例

```powershell
.\Filter-CsvColumns.ps1 `
  -InputCsv "data.csv" `
  -ExcludeColumns @(2,4) `
  -Mode "exact"
