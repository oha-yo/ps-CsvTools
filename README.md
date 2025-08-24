# ps-CsvTools

# 📊 ps-CsvTools

CSVファイルの比較・変換・フィルタリングを効率的に行うPowerShellスクリプト群です。データ検証やETL前処理など、日常的なCSV操作を安全かつ柔軟にサポートします。

---

## 🧩 主なコマンド

| コマンド名 | 説明 | README |
|------------|------|--------|
| `Compare-Csv.ps1` | 2つのCSVを比較し、差分をExcelで出力 | [Compare-Csv README](./Compare-Csv.md) |
| `Convert-CsvToXlsx.ps1` | CSVをExcel形式に変換 | [Convert-CsvToXlsx README](./Convert-CsvToXlsx.md) |
| `Filter-CsvColumns.ps1` | 指定項目を除外したCSVを生成 | [Filter-CsvColumns README](./Filter-CsvColumns.md) |

---
## 必須要件

- PowerShell 7.x
- [EPPlus.dll](https://www.nuget.org/packages/EPPlus)  
  ※本スクリプトでは ImportExcel モジュールに同梱された DLL を使用しますので以下手順で取得してください。
### ダウンロード方法
```bash
mkdir .\Modules
Save-Module -Name ImportExcel -Path .\Modules
```

## 📁 ディレクトリ構成

