function Get-ReaderEncoding  {
    param([string]$Encoding)

    if ($Encoding -eq "utf-8") {
        return New-Object System.Text.UTF8Encoding($true)
    }
    elseif ($Encoding -eq "shift_jis") {
        return [System.Text.Encoding]::GetEncoding("shift_jis")
    }
    elseif ($Encoding -eq "euc-jp") {
        return [System.Text.Encoding]::GetEncoding("euc-jp")
    }
    else {
        return [System.Text.Encoding]::GetEncoding($Encoding)
    }
}