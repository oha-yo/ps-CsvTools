# 🔍 Convert-CsvToXlsx.ps1

このスクリプトは、指定されたCSVファイルを読み込み、選択されたカラムのみを抽出してExcel形式（`.xlsx`）に変換するツールです。大規模データにも対応し、Shift_JISなどの日本語エンコードも扱えます。

---

## 🧩 特徴

- 任意の開始行・最大行数の指定
- カラムの選択（include / exclude モード）
- 区切り文字の正規化（`,` や `\t` など）
- エンコードの正規化（Shift_JIS, UTF-8 等）
- 列番号のヘッダー追加オプション
- Excel出力時のオートフィット対応
- EPPlus.dll を使用した高速なExcel生成

---
## ⚙️ パラメータ一覧

| パラメータ        | 必須 | 型       | デフォルト値 | 説明 |
|-------------------|------|----------|--------------|------|
| `InputFile`       | ✅   | `string` | ―            | 入力CSVファイルのパス。Shift_JIS または UTF-8（BOMあり／なし）に対応。 |
| `StartRow`        |    | `int`    | `1`          | 処理開始行番号（1始まり）。ヘッダーをスキップする場合に指定。 |
| `MaxRows`         |    | `int`    | `0`          | 最大処理行数。`0` を指定すると全行を処理。 |
| `Separator`       |    | `string` | `","`        | 区切り文字。TSVの場合は `"\t"`（タブ）を指定。 |
| `EncodingName`    |    | `string` | `"Shift_JIS"` | ファイルの文字コード。`Shift_JIS` または `UTF-8` を指定。BOMの有無は自動判定。 |
| `AddColumnNumbers`|    | `bool` | `false`        | ヘッダーに列番号を追加するかどうか。 |
| `TargetColumns`   |    | `int[]`  | ―            | 対象カラム番号（1始まり）。複数指定可能（例: `1,3,5`）。省略したら全てのカラムが対象 |
| `Mode`            |    | `string` | `"include"`  | `"include"`：指定カラムのみ抽出、`"exclude"`：指定カラムを除外。 |

---

## 🚀 使用例

```powershell
.\Convert-CsvToXlsx.ps1 `
-InputFile  ".\testdata\test_sjis.csv" `
-AddColumnNumbers $true `
-Mode include `
-TargetColumns 2,3,4

```
