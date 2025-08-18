param(
    [Parameter(Mandatory = $true)][string]$InputFile,
    [Parameter()][int]$StartRow = 1,
    [Parameter()][int]$MaxRows = 0,
    [Parameter()][string]$Separator = ",",
    [Parameter()][string]$Encoding = "Shift_JIS",
    [Parameter()][int[]]$TargetColumns = @(),
    [Parameter()][ValidateSet("exclude", "include")]
    [string]$Mode = "exclude"
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

function Get-StreamReader {
    # 補足 UTF8の場合readerはBOMの有無に関係なく求める事ができる
    param(
        [string]$FilePath,
        [System.Text.Encoding]$Encoding
    )
    
    try {
        return [System.IO.StreamReader]::new($FilePath, $Encoding)
    }
    catch [System.IO.FileNotFoundException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)] ファイルが見つかりません: $FilePath"
        return $null
    }
    catch [System.UnauthorizedAccessException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)]ファイルにアクセスできません。権限を確認してください。"
        return $null
    }
    catch {
        Write-Error "[$($MyInvocation.MyCommand.Name)]その他のエラー: $_.Exception.Message"
        return $null
    }
}


# Stream Writerを取得
function Get-StreamWriter {
    param(
        [string]$OutputFileName,
        [System.Text.Encoding]$Encoding,
        [string]$NewLineChar = "`n"
    )

    try {
        $writer = [System.IO.StreamWriter]::new($OutputFileName, $false, $Encoding)
        $writer.NewLine = $NewLineChar
        return $writer
    }
    catch [System.UnauthorizedAccessException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)]ファイルにアクセスできません。権限を確認してください: $OutputFileName"
        return $null
    }
    catch [System.IO.DirectoryNotFoundException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)]ディレクトリが見つかりません: $OutputFileName"
        return $null
    }
    catch [System.IO.IOException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)]ファイルにアクセスできません（IOエラー）: $OutputFileName"
        return $null
    }
    catch {
        Write-Error "[$($MyInvocation.MyCommand.Name)]その他のエラーが発生しました: $_"
        return $null
    }
}

# インプットファイル存在チェック
if (-not (Test-Path -Path $InputFile -PathType Leaf)) {
    Write-Error "ファイルが見つかりません: $InputFile"
    exit 1
}

# 入力ファイルの改行コード取得とBom判定
$encodingInfo = Get-FileBOMAndNewLine -FilePath $InputFile
$hasBOM       = $encodingInfo.HasBOM
$newLineChar  = $encodingInfo.newLineChar
$displayName  = $encodingInfo.DisplayName
Write-Debug "hasBOM        :$hasBOM"
Write-Debug "newLineChar   :$displayName"

# 出力ファイル(FULL PATH)作成
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
$folderPath = [System.IO.Path]::GetDirectoryName($InputFile)
$inputExtension = [System.IO.Path]::GetExtension($InputFile)
$OutputFileName = [System.IO.Path]::Combine($folderPath, "${baseName}_${Mode}${inputExtension}")
Write-Debug "baseName      :$baseName"
Write-Debug "folderPath    :$folderPath"
Write-Debug "inputExtension:$inputExtension"
Write-Debug "OutputFileName:$OutputFileName"

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

# Stream Writer用エンコーディングの取得
# インプットファイルのエンコーディングに合わせる
$writerEncoding = Get-WriterEncoding -Encoding $Encoding -HasBOM $hasBOM
Write-Debug "writerEncoding:$writerEncoding"

# Stream Writerの取得
$writer = Get-StreamWriter -OutputFileName $OutputFileName -Encoding $writerEncoding -NewLineChar $newLineChar
if ($null -eq $writer) {
    Write-Error "StreamWriterの作成に失敗しました。処理を中断します。"
    exit 1
}

# 実行パラメータを履歴ファイルへ保存
Write-ExecutionHistory

# 対象カラム（0始まりに変換）
$targetIndexes = if ($TargetColumns.Count -gt 0) {
    $TargetColumns | ForEach-Object { $_ - 1 }
} else {
    # 全てのカラムを対象
    0..($allColumns.Count - 1)
}
Write-Debug "targetIndexes:$targetIndexes"

# 行処理開始
$currentLineNumber = 0
$linesWritten = 0
$maxToRead = if ($MaxRows -gt 0) { $MaxRows } else { [int]::MaxValue }
$ProgressInterval = 10000
$splitter = [CsvSplitter]::new($Separator)
while (-not $reader.EndOfStream) {
    $line = $reader.ReadLine()
    $currentLineNumber++
    if ($currentLineNumber -lt $StartRow) {
        continue
    }
    if ($linesWritten -ge $maxToRead) {
        break
    }
    $columns = $splitter.Split($line)
    # Write-Debug "columns count: $(@($columns).Count)"
    $filtered = @(
        for ($i = 0; $i -lt $columns.Count; $i++) {
            $isTarget = @($targetIndexes) -contains $i
            if ($Mode -ieq "exclude" -and -not $isTarget) {
                $columns[$i]
            }
            elseif ($Mode -ieq "include" -and $isTarget) {
                $columns[$i]
            }
        }
    )
    Write-Debug "columns:$columns"
    Write-Debug "filtered:$filtered"

    # 対象カラム配列にセパレータを合わせて文字列化
    #Write-Debug "filtered:$filtered"
    $csvLine = $filtered -join $Separator
    $writer.WriteLine($csvLine)
    $linesWritten++
    if ($linesWritten -ge $ProgressInterval -and $linesWritten % $ProgressInterval -eq 0) {
        Write-Host "$linesWritten 行処理済み..."
    }
}
$reader.Close()
$writer.Close()

Write-Host "出力行数: $linesWritten"
Write-Host "${Mode} 処理後CSV出力完了: $OutputFileName"