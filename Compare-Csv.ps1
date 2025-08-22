param(
    [Parameter(Mandatory = $true)][string]$InCsv1,
    [Parameter(Mandatory = $true)][string]$InCsv2,
    [Parameter()][string]$ResultXlsx,
    [Parameter()][int[]]$KeyItem,
    [Parameter()][int[]]$TargetColumns = @(),
    [Parameter()][int]$StartRow = 1,
    [Parameter()][int]$MaxRows = 0,
    [Parameter()][string]$Encoding = "Shift_JIS",
    [Parameter()][string]$Separator = ",",
    [Parameter()][ValidateSet("exclude", "include")]
    [string]$Mode = "include"
)

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

# CSV読み込み関数
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

    try {
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            ,($splitter.SplitAndClean($line))
        }
    }
    finally { $reader.Close() }
}

# CSV読み込み
$csv1Lines = @((Read-CsvLines -Path $InCsv1 -Encoding $Encoding -Separator $Separator))
$csv2Lines = @((Read-CsvLines -Path $InCsv2 -Encoding $Encoding -Separator $Separator))

if ($csv2Lines.Count -gt $csv1Lines.Count) {
    Write-Warning "CSV2の行数がCSV1より多いため、比較処理をスキップします。"
    Write-Warning "CSV1: $($csv1Lines.Count) 行, CSV2: $($csv2Lines.Count) 行"
    return
}

# イテレータ作成
$enum1Enumerator = $csv1Lines.GetEnumerator()
$enum2Enumerator = $csv2Lines.GetEnumerator()

# StartRow分スキップ
for ($i = 1; $i -lt $StartRow; $i++) {
    if (-not $enum1Enumerator.MoveNext()) {
        Write-Warning "CSV1の行数が StartRow ($StartRow) に満たないため、比較できません。"
        return
    }
    if (-not $enum2Enumerator.MoveNext()) {
        Write-Warning "CSV2の行数が StartRow ($StartRow) に満たないため、比較できません。"
        return
    }
}

# Excel出力準備
$package = New-Object OfficeOpenXml.ExcelPackage
$sheet   = $package.Workbook.Worksheets.Add("Compare")

# 最大列数の決定
$maxCols = ($csv1Lines + $csv2Lines | Select-Object -First 100 | ForEach-Object { $_.Count }) | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

# 比較対象カラムの決定
if ($TargetColumns.Count -eq 0) {
    $TargetColumns = 1..$maxCols
} elseif ($Mode -eq "exclude") {
    $TargetColumns = (1..$maxCols) | Where-Object { $TargetColumns -notcontains $_ }
}

# ヘッダー行
$colIndex = 1
$sheet.Cells.Item(1,$colIndex++).Value = "行番号"

if ($KeyItem) {
    $ki = 1
    foreach ($idx in $KeyItem) {
        $sheet.Cells.Item(1,$colIndex++).Value = "キー項目$ki"
        $ki++
    }
}

foreach ($i in $TargetColumns) {
    $sheet.Cells.Item(1,$colIndex++).Value = "列$($i)"
}

# 比較処理
$rowIndex = 2
$rowNumber = $StartRow - 1
$processedCount = 0

while ($enum1Enumerator.MoveNext()) {
    $rowNumber++
    $processedCount++
    if ($MaxRows -gt 0 -and $processedCount -gt $MaxRows) { break }

    $row1 = $enum1Enumerator.Current
    $row2 = if ($enum2Enumerator.MoveNext()) { $enum2Enumerator.Current } else { @() }

    $colIndex = 1
    $sheet.Cells.Item($rowIndex,$colIndex++).Value = $rowNumber

    foreach ($idx in $KeyItem) {
        $val = if ($idx -le $row1.Count) { $row1[$idx-1] } else { $null }
        $sheet.Cells.Item($rowIndex,$colIndex++).Value = $val
    }

    foreach ($i in $TargetColumns) {
        $val1 = if ($i -le $row1.Count) { $row1[$i-1] } else { $null }
        $val2 = if ($i -le $row2.Count) { $row2[$i-1] } else { $null }
        $sheet.Cells.Item($rowIndex,$colIndex++).Value = if ($val1 -eq $val2) { "〇" } else { "×" }
    }

    $rowIndex++
}

# 保存
$sheet.Cells.AutoFitColumns()
try {
    $package.SaveAs([System.IO.FileInfo]::new($ResultXlsx))
    Write-Host "比較結果を出力しました: $ResultXlsx"
} catch {
    Write-Error "保存時にエラー: $($_.Exception.Message)"
    exit 1
}
