
# Stream Writerを取得
function Get-StreamWriter {
    param(
        [string]$OutputFileName,
        [System.Text.Encoding]$Encoding,
        [string]$NewLineChar = "`n"
    )

    try {
        $writer = [System.IO.StreamWriter]::new($OutputFileName, $false, $Encoding)
        $writer.NewLine = $NewLineChar
        return $writer
    }
    catch [System.UnauthorizedAccessException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)]ファイルにアクセスできません。権限を確認してください: $OutputFileName"
        return $null
    }
    catch [System.IO.DirectoryNotFoundException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)]ディレクトリが見つかりません: $OutputFileName"
        return $null
    }
    catch [System.IO.IOException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)]ファイルにアクセスできません（IOエラー）: $OutputFileName"
        return $null
    }
    catch {
        Write-Error "[$($MyInvocation.MyCommand.Name)]その他のエラーが発生しました: $_"
        return $null
    }
}