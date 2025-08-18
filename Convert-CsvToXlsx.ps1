param(
    [Parameter(Mandatory = $true)][string]$InputFile,
    [Parameter()][string]$Encoding = "Shift_JIS",
    [Parameter()][int]$StartRow = 2,
    [Parameter()][int]$MaxRows = 0,
    [Parameter()][string]$Separator = ",",
    [Parameter()][bool]$AddColumnNumbers = $false,
    [Parameter()][int[]]$TargetColumns = @()  # 空なら全カラム
)

# $PSScriptRoot\Common フォルダ以下のすべての.ps1ファイルを再帰的に取得し、
# それぞれのファイルをドットソーシング（現在のスコープで読み込み）することで、
# 関数やクラスをモジュール内に定義・利用可能にする
Get-ChildItem -Path "$PSScriptRoot\Common" -Recurse -Filter *.ps1 | ForEach-Object {
    . $_.FullName
}

# 実行パラメータを履歴ファイルへ保存
Write-ExecutionHistory

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

Write-Debug "Separator: $Separator"
$splitter = [CsvSplitter]::new($Separator)

# 最初のデータ行を読み取り、列数を判定
$firstDataLine = $reader.ReadLine()
$allColumns = $splitter.Split($firstDataLine)

# 対象カラム（0始まりに変換）
$targetIndexes = if ($TargetColumns.Count -gt 0) {
    $TargetColumns | ForEach-Object { $_ - 1 }
} else {
    # 全てのカラムを対象
    0..($allColumns.Count - 1)
}

# 処理対象カラムのカラム数をカウント
$columnCount = $targetIndexes.Count
# 通番でヘッダー作成
$headers = @(1..$columnCount | ForEach-Object { "$_" })

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

Write-Debug "書き出し完了: $($rowIndex -1)行"
# オートフィット・保存
Write-Debug "ファイル保存中..."
$sheet.Cells.AutoFitColumns()
try {
    $package.SaveAs([System.IO.FileInfo]::new($OutputFile))
    Write-Debug "Excelファイル出力完了: $OutputFile"
} catch {
    Write-Error "ファイル保存時にエラーが発生しました: $($_.Exception.Message)"
    Write-Error "出力ファイルが既に開かれている可能性があります。閉じて再実行してください。"
    exit 1
}