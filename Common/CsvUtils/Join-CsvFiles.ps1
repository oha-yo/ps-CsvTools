function Join-CsvFiles {
    param (
        [string]$Csv1Path,
        [string]$Csv2Path,
        [string]$OutCsvPath,
        [System.Text.Encoding]$Encoding,
        [string]$Separator
    )

    $reader1 = Get-StreamReader -FilePath $Csv1Path -Encoding $Encoding
    $reader2 = Get-StreamReader -FilePath $Csv2Path -Encoding $Encoding
    $writer  = New-StreamWriter -FilePath $OutCsvPath -Encoding $Encoding

    try {
        while (-not $reader1.EndOfStream -and -not $reader2.EndOfStream) {
            $line1 = $reader1.ReadLine()
            $line2 = $reader2.ReadLine()

            # そのままの行テキストを結合（区切りはカンマ＋スペース）
            $joined = "$line1 ,$line2"
            $writer.WriteLine($joined)
        }
    }
    finally {
        $reader1.Close()
        $reader2.Close()
        $writer.Close()
    }
}