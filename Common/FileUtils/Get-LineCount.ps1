#
# 指定ファイルを逐次読み込みし行数をカウントして返却。
#
function Get-LineCount {
    param (
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][System.Text.Encoding]$Encoding
    )

    $count = 0
    try {
        $reader = [System.IO.StreamReader]::new($FilePath, $Encoding)
        try {
            while (-not $reader.EndOfStream) {
                $reader.ReadLine() | Out-Null
                $count++
            }
        }
        finally {
            $reader.Close()
        }
        return $count
    }
    catch {
        Write-Error "[ $($MyInvocation.MyCommand.Name) ] 行数カウント中にエラーが発生しました: $($_.Exception.Message)"
        return -1
    }
}
