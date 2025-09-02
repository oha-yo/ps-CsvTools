function Join-CsvFiles {
    param (
        [string]$Csv1Path,
        [string]$Csv2Path,
        [string]$OutCsvPath,
        [System.Text.Encoding]$Encoding,
        [string]$Separator
    )

    $reader1 = Get-StreamReader $Csv1Path $Encoding
    $reader2 = Get-StreamReader $Csv2Path $Encoding

        # 入力ファイルの改行コード取得とBom判定
    $encodingInfo = Get-FileBOMAndNewLine -FilePath $Csv1Path
    $hasBOM       = $encodingInfo.HasBOM
    $newLineChar  = $encodingInfo.newLineChar
    $displayName  = $encodingInfo.DisplayName
    Write-Debug "hasBOM        :$hasBOM"
    Write-Debug "newLineChar   :$displayName"

    $writer  = Get-StreamWriter $OutCsvPath $Encoding $newLineChar

    try {
        while (-not $reader1.EndOfStream -and -not $reader2.EndOfStream) {
            $line1 = $reader1.ReadLine()
            $line2 = $reader2.ReadLine()
            # そのままの行テキストを結合（区切りは$Separator）
            $joined = "$line1$Separator$line2"
            $writer.WriteLine($joined)
        }
    }
    finally {
        $reader1.Close()
        $reader2.Close()
        $writer.Close()
    }
}