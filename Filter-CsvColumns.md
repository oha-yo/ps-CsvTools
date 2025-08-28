# 🔍 Filter-CsvColumns.md

PowerShellスクリプトによる高速・柔軟なCSVカラム抽出ツール。  
指定したカラムのみを抽出する「include」モードと、指定したカラムを除外する「exclude」モードをサポート。  
Shift_JISやUTF-8などの文字コード、タブ区切りなどのセパレータにも対応。

---

## 🧩 特徴

- 大規模CSVでも高速処理（ストリームベース）
- `include` / `exclude` モードで柔軟なカラム選択
- Shift_JIS, UTF-8, EUC-JP など多様なエンコーディングに対応
- 改行コード・BOMの自動判定
- セパレータの正規化（`,` や `\t` など）
- 処理履歴の保存（`Write-ExecutionHistory`）

---
## ⚙️ パラメータ一覧

| パラメータ        | 必須 | 型       | デフォルト値 | 説明 |
|-------------------|------|----------|--------------|------|
| `InputFile`       | ✅   | `string` | ―            | 入力CSVファイルのパス。Shift_JIS または UTF-8（BOMあり／なし）に対応。 |
| `StartRow`        |    | `int`    | `1`          | 処理開始行番号（1始まり）。ヘッダーをスキップする場合に指定。 |
| `MaxRows`         |    | `int`    | `0`          | 最大処理行数。`0` を指定すると全行を処理。 |
| `Separator`       |    | `string` | `","`        | 区切り文字。TSVの場合は `"\t"`（タブ）を指定。 |
| `EncodingName`    |    | `string` | `"Shift_JIS"` | ファイルの文字コード。`Shift_JIS` または `UTF-8` を指定。BOMの有無は自動判定。 |
| `TargetColumns`   |    | `int[]`  | ―            | 対象カラム番号（1始まり）。複数指定可能（例: `1,3,5`）。省略したら全てのカラムが対象 |
| `Mode`            |    | `string` | `"include"`  | `"include"`：指定カラムのみ抽出、`"exclude"`：指定カラムを除外。 |

---

### 補足

- **カラム番号は1始まり**です（Excelと同じ感覚）。
- **出力ファイル名は自動生成**され、元ファイル名に `_include.csv` または `_exclude.csv` が付加されます。
- **Shift_JIS の場合は文字化け対策として明示的に指定するのが推奨**です。
- **UTF-8 の場合、BOMの有無は自動判定されます**。

---

## 使い方

#### utf-8で書かれたcsvの3,4番目のカラムを除去して、以下ファイル名で出力する。
#### .\testdata\test_utf8_exclude.csv
```powershell
.\Filter-CsvColumns.ps1 `
-InputFile ".\testdata\test_utf8.csv" `
-StartRow 2 `
-Encoding utf8 `
-Separator "," `
-Mode exclude `
-TargetColumns 3,4  

````