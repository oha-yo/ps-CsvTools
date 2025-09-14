# $Pathの存在確認と、相対パスの場合フルパスにして返却
#
function Get-ValidatedFullPath {
    param (
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter()][string]$Label = "ファイル"
    )
    if (-not (Test-Path $Path -PathType Leaf)) {
        throw "$Label が存在しません: $Path"
    }
    return [System.IO.Path]::GetFullPath($Path)
}