function Get-StreamReader {
    # 補足 UTF8の場合readerはBOMの有無に関係なく求める事ができる
    param(
        [string]$FilePath,
        [System.Text.Encoding]$Encoding
    )
    
    try {
        return [System.IO.StreamReader]::new($FilePath, $Encoding)
    }
    catch [System.IO.FileNotFoundException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)] ファイルが見つかりません: $FilePath"
        return $null
    }
    catch [System.UnauthorizedAccessException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)]ファイルにアクセスできません。権限を確認してください。"
        return $null
    }
    catch {
        Write-Error "[$($MyInvocation.MyCommand.Name)]その他のエラー: $_.Exception.Message"
        return $null
    }
}