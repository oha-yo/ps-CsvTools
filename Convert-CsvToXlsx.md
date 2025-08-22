# 🔍 Compare-Csv.ps1

2つのCSVファイルを比較し、差分をExcel形式で出力するスクリプトです。キー項目による行単位の比較を行い、対象列の一致・不一致を明示的に記録します。

---

## 🧰 主な機能

- 複数列によるキー指定
- 比較対象列のフィルタリング（include / exclude）
- 文字コード・区切り文字の指定
- Excel形式での差分出力（〇/×）
- 行数制限・開始行指定

---

## 🚀 使用例

```powershell
.\Compare-Csv.ps1 `
  -InCsv1 "data1.csv" `
  -InCsv2 "data2.csv" `
  -KeyItem @(1,2) `
  -TargetColumns @(3,4) `
  -Mode "exclude"
