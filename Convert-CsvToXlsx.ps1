param(
    [Parameter(Mandatory = $true)][string]$InCsv1,
    [Parameter(Mandatory = $true)][string]$InCsv2,
    [Parameter()][string]$ResultXlsx,
    [Parameter(Mandatory = $true)][int]$KeyItem,          # 1始まりのキー項目
    [Parameter()][int[]]$TargetColumns = @(),             # 空なら全カラム
    [Parameter()][string]$Encoding = "utf-8",
    [Parameter()][string]$Separator = ",",
    [Parameter()][int]$StartRow = 2,                      # スキップする先頭行
    [Parameter()][int]$MaxRows = 0                        # 0なら制限なし
)

# 共通関数ロード
Get-ChildItem -Path "$PSScriptRoot\Common" -Recurse -Filter *.ps1 | ForEach-Object { . $_.FullName }

# 入力ファイル存在チェック
foreach ($file in @($InCsv1,$InCsv2)) {
    if (-not (Test-Path $file)) {
        Write-Error "ファイルが見つかりません: $file"
        exit 1
    }
}

# 出力ファイル名決定
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

# ===== 行単位で読み込み可能なEnumeratorを作成 =====
function Get-CsvEnumerator {
    param(
        [string]$Path,
        [string]$Encoding,
        [string]$Separator,
        [int]$StartRow
    )

    $enc = Convert-EncodingName -enc $Encoding
    $readerEncoding = Get-ReaderEncoding -Encoding $enc
    $reader = Get-StreamReader -FilePath $Path -Encoding $readerEncoding
    if ($null -eq $reader) { throw "Stream Readerの作成に失敗: $Path" }

    $splitter = [CsvSplitter]::new($Separator)

    # StartRowまでスキップ
    for ($i = 1; $i -lt $StartRow; $i++) {
        if (-not $reader.EndOfStream) { $reader.ReadLine() | Out-Null }
    }

    return [PSCustomObject]@{
        Reader   = $reader
        Splitter = $splitter
    }
}

# Enumeratorを取得
$csv1Enum = Get-CsvEnumerator -Path $InCsv1 -Encoding $Encoding -Separator $Separator -StartRow $StartRow
$csv2Enum = Get-CsvEnumerator -Path $InCsv2 -Encoding $Encoding -Separator $Separator -StartRow $StartRow

# Excel準備
$package = New-Object OfficeOpenXml.ExcelPackage
$sheet = $package.Workbook.Worksheets.Add("Compare")

# 最大列数算出（最初の行のみで判定）
$firstRow1 = if (-not $csv1Enum.Reader.EndOfStream) { $csv1Enum.Splitter.SplitAndClean($csv1Enum.Reader.ReadLine()) } else { @() }
$firstRow2 = if (-not $csv2Enum.Reader.EndOfStream) { $csv2Enum.Splitter.SplitAndClean($csv2Enum.Reader.ReadLine()) } else { @() }

# 対象列決定
function Get-TargetIndexes { param($firstRow, $TargetColumns) 
    if ($TargetColumns.Count -gt 0) { return $TargetColumns | ForEach-Object { $_ - 1 } } 
    else { return 0..($firstRow.Count - 1) } 
}
$targetIndexes = Get-TargetIndexes -firstRow $firstRow1 -TargetColumns $TargetColumns

# Excelヘッダー
$sheet.Cells.Item(1,1).Value = "行番号"
$sheet.Cells.Item(1,2).Value = "キー項目"
$colIndex = 3
foreach ($i in $targetIndexes) { $sheet.Cells.Item(1,$colIndex).Value = "列$($i+1)"; $colIndex++ }

# ===== 行単位比較 =====
$rowIndex = 2
$linesProcessed = 0
while ((-not $csv1Enum.Reader.EndOfStream -or -not $csv2Enum.Reader.EndOfStream) -and ($MaxRows -eq 0 -or $linesProcessed -lt $MaxRows)) {
    $line1 = if (-not $csv1Enum.Reader.EndOfStream) { $csv1Enum.Reader.ReadLine() } else { $null }
    $line2 = if (-not $csv2Enum.Reader.EndOfStream) { $csv2Enum.Reader.ReadLine() } else { $null }

    $row1 = if ($line1) { $csv1Enum.Splitter.SplitAndClean($line1) } else { @() }
    $row2 = if ($line2) { $csv2Enum.Splitter.SplitAndClean($line2) } else { @() }

    $sheet.Cells.Item($rowIndex,1).Value = $rowIndex - 1
    $sheet.Cells.Item($rowIndex,2).Value = if ($KeyItem -le $row1.Count) { $row1[$KeyItem-1] }

    $colIndex = 3
    foreach ($i in $targetIndexes) {
        $val1 = if ($i -lt $row1.Count) { $row1[$i] } else { $null }
        $val2 = if ($i -lt $row2.Count) { $row2[$i] } else { $null }
        $sheet.Cells.Item($rowIndex,$colIndex).Value = if ($val1 -eq $val2) { "〇" } else { "×" }
        $colIndex++
    }

    $rowIndex++
    $linesProcessed++
}

# クローズ
$csv1Enum.Reader.Close()
$csv2Enum.Reader.Close()

# 保存
$sheet.Cells.AutoFitColumns()
try { $package.SaveAs([System.IO.FileInfo]::new($ResultXlsx)); Write-Host "比較結果を出力しました: $ResultXlsx" }
catch { Write-Error "保存時にエラー: $($_.Exception.Message)"; exit 1 }
