
function ConvertTo-Encoding {
    param([string]$EncodingName)
#   if ($EncodingName -is [System.Text.Encoding]) {
#       return $EncodingName
#   }
    $Encoding = switch ($EncodingName.ToLower()) {
        "utf-8"     { [System.Text.Encoding]::GetEncoding("utf-8") }
        "shift_jis" { [System.Text.Encoding]::GetEncoding("shift_jis") }
        "euc-jp"    { [System.Text.Encoding]::GetEncoding("euc-jp") }
        default     { [System.Text.Encoding]::GetEncoding($Encoding) }
    }

    if (-not ($Encoding -is [System.Text.Encoding])) {
        throw "指定されたエンコーディング '$EncodingName' は有効な System.Text.Encoding 型ではありません。"
    }
    return $Encoding
}