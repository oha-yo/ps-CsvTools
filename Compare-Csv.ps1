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

# 実行パラメータを履歴ファイルへ保存
Write-ExecutionHistory

# CSV読み込み関数（遅延評価）
function Read-CsvLines {
    param(
        [string]$Path,
        [string]$Encoding,
        [string]$Separator
    )
    $enc = Convert-EncodingName -enc $Encoding
    $readerencoding = Get-ReaderEncoding -Encoding $enc
    $reader = Get-StreamReader -FilePath $Path -Encoding $readerencoding
    $splitter = [CsvSplitter]::new($Separator)

    try {
        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }

            $row = $splitter.SplitAndClean($line)
            Write-Output (, $row)  # ← これが正しい構文
        }
    }
    finally {
        $reader.Close()
    }
}

# Excel出力準備
$package = New-Object OfficeOpenXml.ExcelPackage
$sheet   = $package.Workbook.Worksheets.Add("Compare")

# 最大列数の推定（先頭数行のみ）
$peek1 = @(Read-CsvLines -Path $InCsv1 -Encoding $Encoding -Separator $Separator | Select-Object -First 50)
$peek2 = @(Read-CsvLines -Path $InCsv2 -Encoding $Encoding -Separator $Separator | Select-Object -First 50)
$maxCols = ($peek1 + $peek2 | ForEach-Object { $_.Count }) | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

# 比較対象カラムの決定
if ($TargetColumns.Count -eq 0) {
    $TargetColumns = 1..$maxCols
}
if ($Mode -eq "exclude") {
    $TargetColumns = (1..$maxCols) | Where-Object {
        ($TargetColumns -notcontains $_) -and ($KeyItem -notcontains $_)
    }
}

# ヘッダー行
$colIndex = 1
$sheet.Cells.Item(1,$colIndex++).Value = "行番号"

if ($KeyItem -and $KeyItem.Count -gt 0) {
    $ki = 1
    foreach ($idx in $KeyItem) {
        $sheet.Cells.Item(1,$colIndex++).Value = "キー項目$ki"
        $ki++
    }
}

# 出力列名は TargetColumns の順番に「列1」「列2」…と振り直す
for ($i = 1; $i -le $TargetColumns.Count; $i++) {
    $sheet.Cells.Item(1,$colIndex++).Value = "列$i"
}

# イテレータ（遅延評価）
$enum1Enumerator = @(Read-CsvLines -Path $InCsv1 -Encoding $Encoding -Separator $Separator).GetEnumerator()
$enum2Enumerator = @(Read-CsvLines -Path $InCsv2 -Encoding $Encoding -Separator $Separator).GetEnumerator()


# 比較処理
$rowIndex = 2
$rowNumber = 0
$processedCount = 0

while ($enum1Enumerator.MoveNext()) {
    if (-not $enum2Enumerator.MoveNext()) { break }

    $row1 = $enum1Enumerator.Current
    $row2 = $enum2Enumerator.Current

    $rowNumber++
    if ($rowNumber -lt $StartRow) { continue }

    $processedCount++
    if ($MaxRows -gt 0 -and $processedCount -gt $MaxRows) { break }

    if ($row1 -isnot [System.Collections.IList] -or $row2 -isnot [System.Collections.IList]) {
        Write-Warning "[$rowNumber] 行データが配列ではありません。スキップします。"
        continue
    }

    $colIndex = 1
    $sheet.Cells.Item($rowIndex,$colIndex++).Value = $rowNumber

    if ($KeyItem -and $KeyItem.Count -gt 0) {
        foreach ($idx in $KeyItem) {
            $val = if ($idx -le $row1.Count) { $row1[$idx-1] } else { $null }
            $sheet.Cells.Item($rowIndex,$colIndex++).Value = $val
        }
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
