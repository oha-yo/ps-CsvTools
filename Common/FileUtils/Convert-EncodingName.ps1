function Convert-EncodingName {
    param([string]$enc)
    $e = $enc.ToLower().Replace("_","").Replace("-","")

    switch ($e) {
        { $_ -in @("utf8") } { return "utf-8" }
        { $_ -in @("sjis","shiftjis","shiftjis") } { return "shift_jis" }
        { $_ -in @("euc","eucjp","eucjp") } { return "euc-jp" }
        default { return $enc }
    }
}