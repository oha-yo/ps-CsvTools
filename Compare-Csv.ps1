param(
    [Parameter(Mandatory = $true)][string]$InCsv1,
    [Parameter(Mandatory = $true)][string]$InCsv2,
    [Parameter()][string]$ResultXlsx,
    [Parameter(Mandatory = $true)][int]$KeyItem,          # 1始まりのキー項目
    [Parameter()][int[]]$TargetColumns = @(),             # 空なら全カラム
    [Parameter()][string]$Encoding = "utf-8",
    [Parameter()][string]$Separator = ","
)

# $PSScriptRoot\Common 内の共通関数をロード
Get-ChildItem -Path "$PSScriptRoot\Common" -Recurse -Filter *.ps1 | ForEach-Object {
    . $_.FullName
}

# ===== 入力ファイル存在チェック =====
if (-not (Test-Path $InCsv1)) {
    Write-Error "ファイルが見つかりません: $InCsv1"
    exit 1
}
if (-not (Test-Path $InCsv2)) {
    Write-Error "ファイルが見つかりません: $InCsv2"
    exit 1
}

# ===== 出力ファイル名の決定 =====
if (-not $ResultXlsx) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InCsv1)
    $dir  = [System.IO.Path]::GetDirectoryName((Resolve-Path $InCsv1))
    $ResultXlsx = Join-Path $dir ($base + "_result.xlsx")
}

# ===== EPPlus.dll 読み込み =====
$epplusPath = ".\Modules\ImportExcel\7.8.10\EPPlus.dll"
if (-not (Import-EpplusAssembly -DllPath $epplusPath)) {
    Write-Error "EPPlus.dllが見つかりません: $epplusPath"
    exit 1
}

# ===== CSV 読み込み =====
function Read-CsvLines {
    param(
        [string]$Path,
        [string]$Encoding,
        [string]$Separator
    )
    $enc = Convert-EncodingName -enc $Encoding
    $readerencoding = Get-ReaderEncoding -Encoding $enc
    $reader = Get-StreamReader -FilePath $Path -Encoding $readerencoding
    if ($null -eq $reader) {
        throw "Stream Readerの作成に失敗しました: $Path"
    }

    $splitter = [CsvSplitter]::new($Separator)
    $rows = @()
    while (-not $reader.EndOfStream) {
        $line = $reader.ReadLine()
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue  # 空行やnullはスキップ
        }
        $rows += ,($splitter.SplitAndClean($line))
    }

    $reader.Close()
    return ,$rows
}

$csv1 = Read-CsvLines -Path $InCsv1 -Encoding $Encoding -Separator $Separator
$csv2 = Read-CsvLines -Path $InCsv2 -Encoding $Encoding -Separator $Separator

# ===== Excel 出力準備 =====
$package = New-Object OfficeOpenXml.ExcelPackage
$sheet   = $package.Workbook.Worksheets.Add("Compare")

# 最大列数を算出
$maxCols = [Math]::Max(
    ($csv1 | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum,
    ($csv2 | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
)

# 対象列
$targetIndexes = if ($TargetColumns.Count -gt 0) {
    $TargetColumns | ForEach-Object { $_ - 1 }
} else {
    0..($maxCols-1)
}

# ===== ヘッダー行 =====
$sheet.Cells.Item(1,1).Value = "行番号"
$sheet.Cells.Item(1,2).Value = "キー項目"

$colIndex = 3
foreach ($i in $targetIndexes) {
    $sheet.Cells.Item(1,$colIndex).Value = "列$($i+1)"
    $colIndex++
}

# ===== 比較処理 =====
$maxRows = [Math]::Max($csv1.Count, $csv2.Count)
$rowIndex = 2

for ($r = 0; $r -lt $maxRows; $r++) {
    $row1 = if ($r -lt $csv1.Count) { $csv1[$r] } else { @() }
    $row2 = if ($r -lt $csv2.Count) { $csv2[$r] } else { @() }

    # 行番号
    $sheet.Cells.Item($rowIndex,1).Value = $r + 1
    # キー項目（行1始まり、列は$KeyItem）
    $sheet.Cells.Item($rowIndex,2).Value = if ($KeyItem -le $row1.Count) { $row1[$KeyItem-1] }

    # 対象列比較
    $colIndex = 3
    foreach ($i in $targetIndexes) {
        $val1 = if ($i -lt $row1.Count) { $row1[$i] } else { $null }
        $val2 = if ($i -lt $row2.Count) { $row2[$i] } else { $null }
        $sheet.Cells.Item($rowIndex,$colIndex).Value = if ($val1 -eq $val2) { "〇" } else { "×" }
        $colIndex++
    }
    $rowIndex++
}

# ===== 保存 =====
$sheet.Cells.AutoFitColumns()
try {
    $package.SaveAs([System.IO.FileInfo]::new($ResultXlsx))
    Write-Host "比較結果を出力しました: $ResultXlsx"
} catch {
    Write-Error "保存時にエラー: $($_.Exception.Message)"
    exit 1
}
