#
# 指定ファイルに新規書き込みする StreamWriter を生成し返却する。
# StreamWriter のクローズ責任は呼び出し側にあるので注意
#
function New-StreamWriter {
    param (
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][System.Text.Encoding]$Encoding,
        [bool]$Append = $false  # $false は上書きモード
    )

    try {
        $stream = [System.IO.StreamWriter]::new($FilePath, $false, $Encoding)
        return $stream
    }
    catch {
        Write-Error "[ $($MyInvocation.MyCommand.Name) ] StreamWriterの生成に失敗しました: $($_.Exception.Message)"
        return $null
    }
}