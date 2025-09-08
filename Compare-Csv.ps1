param(
    [Parameter(Mandatory = $true)][string]$InCsv1,
    [Parameter(Mandatory = $true)][string]$InCsv2,
    [Parameter()][string]$ResultXlsx,
    [Parameter()][int[]]$KeyItem,
    [Parameter()][int]$StartRow = 1,
    [Parameter()][int]$MaxRows = 0,
    [Parameter()][string]$Separator = ",",
    [Parameter()][string]$EncodingName = "Shift_JIS",
    [Parameter()][int[]]$TargetColumns = @(),
    [Parameter()][ValidateSet("exclude", "include")]
    [string]$Mode = "include"
)
# function Compare-Columns {
#     param(
#         [string[]]$Row1,
#         [string[]]$Row2,
#         [int[]]$CompareIndexes
#     )
# 
#     $results = @()
#     foreach ($i in $CompareIndexes) {
#         $val1 = if ($i -le $Row1.Count) { $Row1[$i - 1] } else { "<null>" }
#         $val2 = if ($i -le $Row2.Count) { $Row2[$i - 1] } else { "<null>" }
#         $results += if ($val1 -eq $val2) { "〇" } else { "×" }
#     }
#     return $results
# }
# 
# function Get-DiffColumnText {
#     param (
#         [string[]]$Values1,
#         [string[]]$Values2,
#         [int[]]$CompareIndexes
#     )
# 
#     #$diffColumns = @()
#     $diffColumns = [System.Collections.Generic.List[int]]::new()
#     foreach ($i in $CompareIndexes) {
#         $val1 = if ($i -le $Values1.Count) { $Values1[$i - 1] } else { "<null>" }
#         $val2 = if ($i -le $Values2.Count) { $Values2[$i - 1] } else { "<null>" }
# 
#         if ($val1 -ne $val2) {
#         #if ('"{0}"' -f $val1 -ne '"{0}"' -f $val2) {
#             #$diffColumns += $i
#             $diffColumns.Add($i)
#         }
#     }
# 
#     if ($diffColumns.Count -gt 0) {
#         # List を配列に変換して文字列結合
#         return "No. " + ($diffColumns.ToArray() -join ",")
#         #return "No. " + ($diffColumns -join ",")
#     } else {
#         return "無し"
#     }
# }
# 
# function Write-CompareRowToSheet {
#     param(
#         [OfficeOpenXml.ExcelWorksheet]$Sheet,
#         [int]$RowIndex,
#         [string[]]$Row1,
#         [string[]]$Row2,
#         [int[]]$EffectiveColumns,
#         [int[]]$KeyItem,
#         [int]$LineNo
#     )
#     $colIndex = 1
#     # 行番号
#     $Sheet.Cells.Item($RowIndex, $colIndex++).Value = $LineNo
#     # キー項目
#     foreach ($idx in $KeyItem) {
#         $val = if ($idx -le $Row1.Count) { $Row1[$idx - 1] } else { "" }
#         $Sheet.Cells.Item($RowIndex, $colIndex++).Value = $val
#     }
#     # 列比較結果（〇/×）
#     $compareResults = Compare-Columns -Row1 $Row1 -Row2 $Row2 -CompareIndexes $EffectiveColumns
#     foreach ($res in $compareResults) {
#         $Sheet.Cells.Item($RowIndex, $colIndex++).Value = $res
#     }
#     # 差異列番号
#     $diffText = Get-DiffColumnText $Row1 $Row2 $EffectiveColumns
#     $Sheet.Cells.Item($RowIndex, $colIndex++).Value = $diffText
# }

# 共通関数ロード
Get-ChildItem -Path "$PSScriptRoot\Common" -Recurse -Filter *.ps1 | ForEach-Object {
    . $_.FullName
}
# 区切り文字を 内部処理用に正規化
$Separator = Format-Separator $Separator
#エンコード名の正規化(曖昧な入力エンコードをPowershellの正規なエンコード名に変換)
$EncodingName = ConvertTo-EncodingName $EncodingName
Write-Debug "EncodingName  :$EncodingName"

# 入力チェック
if (-not (Test-Path $InCsv1)) { Write-Error "ファイルが見つかりません: $InCsv1"; exit 1 }
if (-not (Test-Path $InCsv2)) { Write-Error "ファイルが見つかりません: $InCsv2"; exit 1 }

# 比較結果出力先ファイル名の取得
if (-not $ResultXlsx) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InCsv1)
    $dir  = [System.IO.Path]::GetDirectoryName((Resolve-Path $InCsv1))
    $ResultXlsx = Join-Path $dir ($base + "_result.xlsx")
}

# EPPlus.dll 読み込み
$epplusPath = ".\Modules\ImportExcel\7.8.10\EPPlus.dll"
if (-not (Import-EpplusAssembly -DllPath $epplusPath)) {
    Write-Error "EPPlus.dllが見つかりません: $epplusPath"
    exit 1
}
# 比較対象ファイルの行数カウント
$Encoding = ConvertTo-Encoding $EncodingName
$lineCount1 = Get-LineCount $InCsv1 $Encoding
$lineCount2 = Get-LineCount $InCsv2 $Encoding
if ($lineCount1 -ne $lineCount2) {
    Write-Error "CSVファイルのレコード数が一致しません。比較できません。"
    Write-Error "InCsv1: $lineCount1 行, InCsv2: $lineCount2 行"
    exit 1
}

# メモリを効率よく利用するためあらかじめCSV1とCSV2を結合しtemp_compare.csvを作成する。
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($InCsv1)
$directory = [System.IO.Path]::GetDirectoryName((Resolve-Path $InCsv1))
$OutCsvPath = Join-Path $directory "$baseName`_temp_compare.csv"
Write-Debug "比較用一時テーブル:$OutCsvPath"
Write-Debug "EncodingName    :$EncodingName"
Write-Debug "Encoding        :$Encoding"
Join-CsvFiles -Csv1Path $InCsv1 $InCsv2 $OutCsvPath $Encoding $Separator
Write-Debug "比較用一時テーブルを作成しました。"

# 比較対象先頭行から比較対象カラム数を求める。
$maxCols = Get-CsvColumnCount $InCsv1 $Encoding $Separator $StartRow
Write-Debug "対象行のカラム数: $maxCols"

# 比較対象カラムの決定
if ($TargetColumns.Count -eq 0) {
    Write-Debug "TargetColumns が未指定または空のため、全カラムを対象にします。"
    $TargetColumns = 1..$maxCols
}
# Modeによって比較対象カラムが決定する
$effectiveColumns = @()
if ($Mode -eq "include") {
    $effectiveColumns = $TargetColumns
}
elseif ($Mode -eq "exclude") {
    $effectiveColumns = (1..$maxCols) | Where-Object {
        $TargetColumns -notcontains $_
    }
}
if ($effectiveColumns.Count -eq 0) {
    Write-Warning "比較対象列が空です。TargetColumns の指定を確認してください。"
    exit 1
}

# Excel出力準備
$package = New-Object OfficeOpenXml.ExcelPackage
$sheet   = $package.Workbook.Worksheets.Add("Compare")

# 結果出力ファイルのヘッダー行書き込み
$colIndex = 1
$sheet.Cells.Item(1,$colIndex++).Value = "行番号"
if ($KeyItem.Count -gt 0) {
    $ki = 1
    foreach ($idx in $KeyItem) {
        $sheet.Cells.Item(1,$colIndex++).Value = "キー項目$ki"
        $ki++
    }
}
# 出力列名は TargetColumns の順番に「列1」「列2」…と振り直す
foreach ($colNum in $effectiveColumns) {
    $sheet.Cells.Item(1, $colIndex++).Value = "列$colNum"
}
$sheet.Cells.Item(1, $colIndex++).Value = "相違カラムNo."

# 比較データファイルのリーダーを取得
$reader = Get-StreamReader -FilePath $OutCsvPath -Encoding $Encoding
$splitter = [CsvSplitter]::new($Separator)
$rowIndex = 2
$lineNo = 1
try {
    while (-not $reader.EndOfStream) {
        $line = $reader.ReadLine()
        # MaxRows制限
        if ($MaxRows -gt 0 -and $lineNo -gt $MaxRows) { break }
        # StartRowスキップ
        if ($lineNo -lt $StartRow) {
            $lineNo++
            continue
        }
        $row = $splitter.SplitAndClean($line)
        $row1 = $row[0..($maxCols - 1)]
        $row2 = $row[$maxCols..($row.Count - 1)]
        Write-CompareRowToSheet $sheet $rowIndex $row1 $row2 $effectiveColumns $KeyItem $lineNo
        $rowIndex++
        $lineNo++
    }
}
finally {
    $reader.Close()
}

$status_code=1
try {
    # Excelへ保存
    $sheet.Cells.AutoFitColumns()
    $package.SaveAs([System.IO.FileInfo]::new($ResultXlsx))
    Write-Information "比較結果を出力しました: $ResultXlsx"
    # 実行パラメータを履歴ファイルへ保存
    Write-ExecutionHistory
    $status_code=0
    Write-Host "出力レコード数: $($rowIndex - 2)"
    Write-Host "出力ファイル: $(Resolve-Path $ResultXlsx)"

} catch {
    Write-Error "保存時にエラー: $($_.Exception.Message)"
    $status_code=1
}
exit $status_code