# 文字コード名の表記揺れを正規の名称へ統一する
#エンコード名の正規化(曖昧な入力エンコードをPowershellの正規なエンコード名に変換)
function ConvertTo-EncodingName {
    param([string]$EncodingName)
    $e = $EncodingName.ToLower().Replace("_","").Replace("-","")

    switch ($e) {
        { $_ -in @("utf8") } { return "utf-8" }
        { $_ -in @("sjis","shiftjis","shiftjis") } { return "shift_jis" }
        { $_ -in @("euc","eucjp","eucjp") } { return "euc-jp" }
        default { return $EncodingName }
    }
}