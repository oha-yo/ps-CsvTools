param(
    [Parameter(Mandatory = $true)][string]$InputFile,
    [Parameter()][int]$StartRow = 1,
    [Parameter()][int]$MaxRows = 0,
    [Parameter()][string]$Separator = ",",
    [Parameter()][string]$Encoding = "Shift_JIS",
    [Parameter()][bool]$AddColumnNumbers = $false,
    [Parameter()][int[]]$TargetColumns = @(),  # 空なら全カラム
    [Parameter()][ValidateSet("include","exclude")]
    [string]$Mode = "include"
)

# 区切り文字の正規化
# Powershellでは「タブ」を `t で表記するため
switch ($Separator) {
    '\t' { $Separator = "`t" }
    '\\t' { $Separator = "`t" }
}

# $PSScriptRoot\Common フォルダ以下のすべての.ps1ファイルを再帰的に取得し、
# それぞれのファイルをドットソーシング（現在のスコープで読み込み）することで、
# 関数やクラスをモジュール内に定義・利用可能にする
Get-ChildItem -Path "$PSScriptRoot\Common" -Recurse -Filter *.ps1 | ForEach-Object {
    . $_.FullName
}

# EPPlus.dll の読み込み（ImportExcelモジュールから直接）
$epplusPath = ".\Modules\ImportExcel\7.8.10\EPPlus.dll"
if (-not (Import-EpplusAssembly -DllPath $epplusPath)) {
    Write-Error "EPPlus.dllが見つかりません: $epplusPath"
    exit 1
}
# インプットファイル存在チェック
if (-not (Test-Path -Path $InputFile -PathType Leaf)) {
    Write-Error "ファイルが見つかりません: $InputFile"
    exit 1
}
Write-Debug "InputFile     : $InputFile"

# 出力ファイル(FULL PATH)作成
$OutputFile = [System.IO.Path]::ChangeExtension($InputFile, "xlsx")
Write-Debug "OutputFile    : $OutputFile"

#エンコード名の正規化(曖昧な入力エンコードをPowershellの正規なエンコード名に変換)
$Encoding = Convert-EncodingName -enc $Encoding
Write-Debug "Encoding      :$Encoding"

# Stream Reader用エンコード取得
$readerencoding = Get-ReaderEncoding -Encoding $Encoding
Write-Debug "readerencoding:$readerencoding"

# Stream Readerの取得
$reader = Get-StreamReader -FilePath $InputFile -Encoding $readerencoding
if ($null -eq $reader) {
    Write-Error "Stream Readerの作成に失敗しました。処理を中断します。"
    exit 1
}

# 指定行までスキップ
$currentLineNumber = 0
while (-not $reader.EndOfStream -and $currentLineNumber -lt ($StartRow - 1)) {
    $reader.ReadLine() | Out-Null
    $currentLineNumber++
}

# 対象の先頭行を読み取り、列数を判定
$firstDataLine = $reader.ReadLine()
$splitter = [CsvSplitter]::new($Separator)
$allColumns = $splitter.Split($firstDataLine)

$columnCount = $allColumns.Count
$allIndexes = 0..($columnCount - 1)

# 対象カラムインデックスの取得
$targetIndexes = if ($TargetColumns.Count -gt 0) {
    $adjusted = $TargetColumns | Where-Object { $_ -ge 1 } | ForEach-Object { $_ - 1 }
    if ($Mode -eq "exclude") {
        $allIndexes | Where-Object { $adjusted -notcontains $_ }
    } else {
        $adjusted
    }
} else {
    # TargetColumnsを指定しない場合はModeを問わず全カラムを対象にする。
    $allIndexes
}

$headers = if ($TargetColumns.Count -gt 0) {
    if ($Mode -eq "exclude") {
        $targetIndexes | ForEach-Object { "$($_ + 1)" }
    } else {
        $TargetColumns | ForEach-Object { "$_" }
    }
} else {
    1..$columnCount | ForEach-Object { "$_" }
}

# データ行を List[string] で読み込み（1行目含む）
$linesToProcess = [System.Collections.Generic.List[string]]::new()
$linesToProcess.Add($firstDataLine)

$maxToRead = if ($MaxRows -gt 0) { $MaxRows - 1 } else { [int]::MaxValue }
while (-not $reader.EndOfStream -and $linesToProcess.Count -lt $maxToRead + 1) {
    $linesToProcess.Add($reader.ReadLine())
    if ($linesToProcess.Count % 50000 -eq 0) {
        Write-Debug "読み込み中: $($linesToProcess.Count) 行..."
    }
}
$reader.Close()
Write-Debug "InputFile読み込み完了: $($linesToProcess.Count) 行"

# Excelファイル作成
$package = New-Object OfficeOpenXml.ExcelPackage
$sheet = $package.Workbook.Worksheets.Add("Sheet1")

# ヘッダー出力（1行目に列通番をつける）
if ($AddColumnNumbers) {
    for ($j = 0; $j -lt $headers.Length; $j++) {
        $sheet.Cells.Item(1, $j + 1).Value = $headers[$j]
    }
}

# データ出力（ヘッダーの有無を考慮）
$rowIndex = if ($AddColumnNumbers) { 2 } else { 1 }
# インプットファイルをメモリに読み上げた後１行づつエクセルへ書き込む
foreach ($line in $linesToProcess) {
    $columns = $splitter.SplitAndClean($line)
    for ($i = 0; $i -lt $targetIndexes.Count; $i++) {
        $srcIndex = $targetIndexes[$i]
        $value = if ($srcIndex -lt $columns.Count) { $columns[$srcIndex] } else { $null }
        $sheet.Cells.Item($rowIndex, $i + 1).Value = $value
    }
    $rowIndex++
    if ($rowIndex % 50000 -eq 0) {
        Write-Debug "書き出し中: $rowIndex 行目..."
    }
}

# オートフィット・保存
Write-Debug "ファイル保存中..."
$sheet.Cells.AutoFitColumns()
try {
    $package.SaveAs([System.IO.FileInfo]::new($OutputFile))
} catch {
    Write-Error "ファイル保存時にエラーが発生しました: $($_.Exception.Message)"
    Write-Error "出力ファイルが既に開かれている可能性があります。閉じて再実行してください。"
    exit 1
}
# 実行パラメータを履歴ファイルへ保存
Write-ExecutionHistory
Write-Host "出力レコード数:  $($rowIndex -1)"
Write-Host "エクセル出力完了(${Mode}): $OutputFile"