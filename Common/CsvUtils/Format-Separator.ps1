# 区切り文字を 内部処理用に正規化
# Powershellでは「タブ」を `t で表記するため
function Format-Separator {
    param(
        [Parameter(Mandatory)]
        [string]$sep
    )
    switch ($sep) {
        '\t'   { return "`t" }
        '\\t'  { return "`t" }
        default { return $sep }
    }
}
