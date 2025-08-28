function Get-CsvColumnCount {
    param (
        [string]$FilePath,
        [System.Text.Encoding]$Encoding,
        [string]$Separator,
        [int]$StartRow = 1
    )

    if (-not ($Encoding -is [System.Text.Encoding])) {
        Write-Error "EncodingはSystem.Text.Encoding型である必要があります。"
        return $null
    }

    #Write-Debug "Encoding--->:$Encoding"
    $reader = Get-StreamReader $FilePath $Encoding
    if ($null -eq $reader) {
        Write-Error "Stream Readerの作成に失敗しました。"
        return $null
    }

    $currentLineNumber = 0
    while (-not $reader.EndOfStream -and $currentLineNumber -lt ($StartRow - 1)) {
        $reader.ReadLine() | Out-Null
        $currentLineNumber++
    }

    $targetLine = $reader.ReadLine()
    $reader.Close()

    if ($null -eq $targetLine) {
        Write-Error "指定行が存在しません。"
        return $null
    }

    $splitter = [CsvSplitter]::new($Separator)
    $columns = $splitter.Split($targetLine)
    return $columns.Count
}