param(
    [Parameter(Mandatory = $true)][string]$InCsv1,
    [Parameter(Mandatory = $true)][string]$InCsv2,
    [Parameter()][string]$ResultXlsx,
    [Parameter()][int[]]$KeyItem,
    [Parameter()][int]$StartRow = 1,
    [Parameter()][int]$MaxRows = 0,
    [Parameter()][string]$Separator = ",",
    [Parameter()][string]$Encoding = "Shift_JIS",
    [Parameter()][int[]]$TargetColumns = @(),
    [Parameter()][ValidateSet("exclude", "include")]
    [string]$Mode = "include"
)
# 区切り文字の正規化
# Powershellでは「タブ」を `t で表記するため
switch ($Separator) {
    '\t' { $Separator = "`t" }
    '\\t' { $Separator = "`t" }
}
# 共通関数ロード
Get-ChildItem -Path "$PSScriptRoot\Common" -Recurse -Filter *.ps1 | ForEach-Object {
    . $_.FullName
}

# 入力チェック
if (-not (Test-Path $InCsv1)) { Write-Error "ファイルが見つかりません: $InCsv1"; exit 1 }
if (-not (Test-Path $InCsv2)) { Write-Error "ファイルが見つかりません: $InCsv2"; exit 1 }

# 出力先
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

function Get-LineCount {
    param (
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][System.Text.Encoding]$Encoding
    )

    $count = 0
    try {
        $reader = [System.IO.StreamReader]::new($FilePath, $Encoding)
        while (-not $reader.EndOfStream) {
            $reader.ReadLine() | Out-Null
            $count++
        }
        $reader.Close()
        return $count
    }
    catch {
        Write-Error "行数カウント中にエラーが発生しました: $($_.Exception.Message)"
        return -1
    }
}

function Join-CsvFiles {
    param (
        [string]$Csv1Path,
        [string]$Csv2Path,
        [string]$OutCsvPath,
        [System.Text.Encoding]$Encoding,
        [string]$Separator
    )

    $reader1 = Get-StreamReader -FilePath $Csv1Path -Encoding $Encoding
    $reader2 = Get-StreamReader -FilePath $Csv2Path -Encoding $Encoding
    $writer  = New-StreamWriter -FilePath $OutCsvPath -Encoding $Encoding

    try {
        while (-not $reader1.EndOfStream -and -not $reader2.EndOfStream) {
            $line1 = $reader1.ReadLine()
            $line2 = $reader2.ReadLine()

            # そのままの行テキストを結合（区切りはカンマ＋スペース）
            $joined = "$line1 ,$line2"
            $writer.WriteLine($joined)
        }
    }
    finally {
        $reader1.Close()
        $reader2.Close()
        $writer.Close()
    }
}

function New-StreamWriter {
    param (
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][System.Text.Encoding]$Encoding
    )

    try {
        $stream = [System.IO.StreamWriter]::new($FilePath, $false, $Encoding)
        return $stream
    }
    catch {
        Write-Error "StreamWriterの生成に失敗しました: $($_.Exception.Message)"
        return $null
    }
}

$readerencoding = Get-ReaderEncoding -Encoding $Encoding
$lineCount1 = Get-LineCount -FilePath $InCsv1 -Encoding $readerencoding
$lineCount2 = Get-LineCount -FilePath $InCsv2 -Encoding $readerencoding

if ($lineCount1 -ne $lineCount2) {
    Write-Error "CSVファイルのレコード数が一致しません。比較できません。"
    Write-Error "InCsv1: $lineCount1 行, InCsv2: $lineCount2 行"
    exit 1
}

$readerencoding = Get-ReaderEncoding -Encoding $Encoding
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($InCsv1)
$directory = [System.IO.Path]::GetDirectoryName((Resolve-Path $InCsv1))
$OutCsvPath = Join-Path $directory "$baseName`_compare.csv"
Write-Debug "比較用一時テーブル: $OutCsvPath"
Write-Debug "Encoding            :$Encoding"
Write-Debug "readerencoding      :$readerencoding"
Join-CsvFiles -Csv1Path $InCsv1  -Csv2Path $InCsv2 -OutCsvPath $OutCsvPath -Encoding $readerencoding -Separator $Separator
Write-Debug "比較用一時テーブルを作成しました。"


# Excel出力準備
$package = New-Object OfficeOpenXml.ExcelPackage
$sheet   = $package.Workbook.Worksheets.Add("Compare")
$maxCols = Get-CsvColumnCount -FilePath $InCsv1 `
    -Encoding $readerencoding `
    -Separator $Separator `
    -StartRow $StartRow
Write-Debug "対象行のカラム数: $maxCols"

# 比較対象カラムの決定
if ($TargetColumns.Count -eq 0) {
    Write-Debug "TargetColumns が未指定または空のため、全カラムを対象にします。"
    $TargetColumns = 1..$maxCols
}

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

# ヘッダー行
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

$reader = Get-StreamReader -FilePath $OutCsvPath -Encoding $readerencoding
$splitter = [CsvSplitter]::new($Separator)

$rowIndex = 2
$lineNo = 1

while (-not $reader.EndOfStream) {
    $line = $reader.ReadLine()

    # MaxRows制限
    if ($MaxRows -gt 0 -and $lineNo -gt $MaxRows) { break }

    # StartRowスキップ
    if ($lineNo -lt $StartRow) {
        $lineNo++
        continue
    }

    # 1行を分割（左右のCSVが連結されている前提）
    $row = $splitter.SplitAndClean($line)
    $row1 = $row[0..($maxCols - 1)]
    $row2 = $row[$maxCols..($row.Count - 1)]

    $colIndex = 1
    $sheet.Cells.Item($rowIndex, $colIndex++).Value = $lineNo

    # キー項目出力（左側CSVから）
    foreach ($idx in $KeyItem) {
        $val = if ($idx -le $row1.Count) { $row1[$idx - 1] } else { "" }
        $sheet.Cells.Item($rowIndex, $colIndex++).Value = $val
    }

    # 比較結果（〇/×）
    foreach ($i in $effectiveColumns) {
        $raw1 = if ($i -le $row1.Count) { $row1[$i - 1] } else { "<null>" }
        $raw2 = if ($i -le $row2.Count) { $row2[$i - 1] } else { "<null>" }

        $val1 = '"{0}"' -f $raw1
        $val2 = '"{0}"' -f $raw2

        Write-Debug "Compare: val1=$val1 val2=$val2"

        $result = if ($val1 -eq $val2) { "〇" } else { "×" }
        $sheet.Cells.Item($rowIndex, $colIndex++).Value = $result
    }

    $rowIndex++
    $lineNo++
}

$reader.Close()

## --------------------------------------------------------------------------
# 保存
$sheet.Cells.AutoFitColumns()
try {
    $package.SaveAs([System.IO.FileInfo]::new($ResultXlsx))
    Write-Host "比較結果を出力しました: $ResultXlsx"
    # 実行パラメータを履歴ファイルへ保存
    Write-ExecutionHistory
} catch {
    Write-Error "保存時にエラー: $($_.Exception.Message)"
    exit 1
}