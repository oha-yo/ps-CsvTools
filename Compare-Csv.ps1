param(
    [Parameter(Mandatory = $true)][string]$InCsv1,
    [Parameter(Mandatory = $true)][string]$InCsv2,
    [Parameter()][string]$ResultXlsx,
    [Parameter(Mandatory = $true)][int[]]$KeyItem,        # 1始まりのキー項目（複数可）
    [Parameter()][int[]]$TargetColumns = @(),             # 空なら全カラム
    [Parameter()][int]$StartRow = 1,                      # 1始まりの行番号
    [Parameter()][int]$MaxRows = 0,                       # 0は制限なし
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

# EPPlus.dll
$epplusPath = ".\Modules\ImportExcel\7.8.10\EPPlus.dll"
if (-not (Import-EpplusAssembly -DllPath $epplusPath)) {
    Write-Error "EPPlus.dllが見つかりません: $epplusPath"
    exit 1
}

# --- CSV 読み込み関数（イテレータ） ---
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

# --- イテレータ作成 ---
$enum1 = Read-CsvLines -Path $InCsv1 -Encoding $Encoding -Separator $Separator
$enum2 = Read-CsvLines -Path $InCsv2 -Encoding $Encoding -Separator $Separator

# --- 行数チェック ---
$csv1Lines = @($enum1)
$csv2Lines = @($enum2)

if ($csv2Lines.Count -gt $csv1Lines.Count) {
    Write-Warning "CSV2の行数がCSV1より多いため、比較処理をスキップします。"
    Write-Warning "CSV1: $($csv1Lines.Count) 行, CSV2: $($csv2Lines.Count) 行"
    return
}

# イテレータを再生成（前の @() 展開で消費されたため）
$enum1Enumerator = $csv1Lines.GetEnumerator()
$enum2Enumerator = $csv2Lines.GetEnumerator()


# --- Excel 出力 ---
$package = New-Object OfficeOpenXml.ExcelPackage
$sheet   = $package.Workbook.Worksheets.Add("Compare")

# 最大列数の計算（両CSVの先頭数行を見て決定）
$maxCols = 0
foreach ($row in ($enum1 + $enum2 | Select-Object -First 100)) {
    if ($row.Count -gt $maxCols) { $maxCols = $row.Count }
}

# 比較対象カラムの決定
if ($TargetColumns.Count -eq 0) {
    $TargetColumns = 1..$maxCols
} elseif ($Mode -eq "exclude") {
    $TargetColumns = (1..$maxCols) | Where-Object { $TargetColumns -notcontains $_ }
}

# --- ヘッダー行 ---
$colIndex = 1
$sheet.Cells.Item(1,$colIndex++).Value = "行番号"

$ki = 1
foreach ($idx in $KeyItem) {
    $sheet.Cells.Item(1,$colIndex++).Value = "キー項目$ki"
    $ki++
}

foreach ($i in $TargetColumns) {
    $sheet.Cells.Item(1,$colIndex++).Value = "列$($i)"
}

# --- 比較処理 ---
$rowIndex = 2
$rowNumber = 0

$enum1Enumerator = $enum1.GetEnumerator()
$enum2Enumerator = $enum2.GetEnumerator()

while ($enum1Enumerator.MoveNext()) {
    $rowNumber++
    if ($rowNumber -lt $StartRow) { continue }
    if ($MaxRows -gt 0 -and ($rowNumber - $StartRow + 1) -gt $MaxRows) { break }

    $row1 = $enum1Enumerator.Current
    $row2 = if ($enum2Enumerator.MoveNext()) { $enum2Enumerator.Current } else { @() }

    $colIndex = 1
    $sheet.Cells.Item($rowIndex,$colIndex++).Value = $rowNumber

    # --- キー項目出力（CSV1の値をそのまま） ---
    foreach ($idx in $KeyItem) {
        $val = if ($idx -le $row1.Count) { $row1[$idx-1] } else { $null }
        $sheet.Cells.Item($rowIndex,$colIndex++).Value = $val
    }

    # --- 比較処理 ---
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
