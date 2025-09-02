# 🔍 Compare-Csv.ps1

2つのCSVファイルを比較し、差分をExcelファイル（.xlsx）として出力するPowerShellスクリプトです。行毎に指定列単位の比較を行い、対象列の一致・不一致を明示的に記録します。キー項目（KeyItem）で指定した列は比較結果ファイル(.xlsx)へCSVファイル１の値をそのまま出力します。差分検証の際ご利用ください。

---

## 🧩 特徴

- 任意の列をキー項目として指定可能（複数可）
- 比較対象列の柔軟なフィルタリング（`include` / `exclude` モード）
- UTF-8以外の文字コードにも対応
- 区切り文字のカスタマイズ
- 比較結果をExcel形式で出力（EPPlus.dll使用）
- 行数制限や開始行の指定も可能
- CSV2がCSV1より長い場合は警告を出して比較処理をスキップ

---

## 📦 必要環境

- PowerShell 7.x 以上
- `EPPlus.dll`（`Modules/ImportExcel/7.8.10/EPPlus.dll` に配置）
- 共通関数群（`Common/*.ps1`）が `$PSScriptRoot/Common` に存在すること

---
## ⚙️ パラメータ一覧

| パラメータ     | 必須 | 説明                                                                 |
|---------------|------|----------------------------------------------------------------------|
| `InCsv1`        | ✔    | 比較元CSVファイルのパス                                               |
| `InCsv2`        | ✔    | 比較対象CSVファイルのパス                                             |
| `ResultXlsx`    |      | 出力先Excelファイルのパス（省略時は自動生成）                         |
| `KeyItem`       |      | キー項目の列番号（1始まり、複数可）                                   |
| `StartRow`      |      | 比較開始行（1始まり）                                                 |
| `MaxRows`       |      | 最大比較行数（0なら全ての行を対象）                                    |
| `Separator`     |      | CSVファイル区切り文字（例: `","`, `"\t"`）                            |
| `EncodingName`  |      | 比較元ファイルの文字コード（例: `"utf-8"`, `"shift_jis"`）             |
| `TargetColumns` |      | 比較対象列の番号（1始まり）                                           |
| `Mode`          |      | `"include"` または `"exclude"`（include:指定列の比較  exclude:指定列を除いた比較）|

---

## 🚀 使い方

```powershell
.\Compare-Csv.ps1 `
  -InCsv1 ./testdata/fruit1.csv `
  -InCsv2 ./testdata/fruit2.csv `
  -KeyItem 1,2 `
  -EncodingName utf-8 `
  -Separator "," `
  -TargetColumns 3,2 `
  -StartRow 1 `
  -Mode exclude
