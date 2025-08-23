# function Get-ReaderEncoding  {
#     param([string]$Encoding)
# 
#     if ($Encoding -eq "utf-8") {
#         return New-Object System.Text.UTF8Encoding($true)
#     }
#     elseif ($Encoding -eq "shift_jis") {
#         return [System.Text.Encoding]::GetEncoding("shift_jis")
#     }
#     elseif ($Encoding -eq "euc-jp") {
#         return [System.Text.Encoding]::GetEncoding("euc-jp")
#     }
#     else {
#         return [System.Text.Encoding]::GetEncoding($Encoding)
#     }
# }

function Get-ReaderEncoding {
    param([string]$Encoding)
    if ($Encoding -is [System.Text.Encoding]) {
        return $Encoding
    }
    $enc = switch ($Encoding.ToLower()) {
        "utf-8"     { New-Object System.Text.UTF8Encoding($true) }
        "shift_jis" { [System.Text.Encoding]::GetEncoding("shift_jis") }
        "euc-jp"    { [System.Text.Encoding]::GetEncoding("euc-jp") }
        default     { [System.Text.Encoding]::GetEncoding($Encoding) }
    }

    if (-not ($enc -is [System.Text.Encoding])) {
        throw "指定されたエンコーディング '$Encoding' は有効な System.Text.Encoding 型ではありません。"
    }

    return $enc
}
