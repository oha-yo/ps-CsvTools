function Get-FileBOMAndNewLine {
    param(
        [string]$FilePath
    )

    try {
        # BOM判定と改行コード検出（先頭4KBのみ読み込み）
        $fs = [System.IO.FileStream]::new($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
        $buffer = New-Object byte[] 4096
        $bytesRead = $fs.Read($buffer, 0, $buffer.Length)
        $fs.Close()

        # BOMの判定
        $hasBOM = ($bytesRead -ge 3 -and
            $buffer[0] -eq 0xEF -and
            $buffer[1] -eq 0xBB -and
            $buffer[2] -eq 0xBF)
        
        # 改行コード取得
        $nlInfo = DetectFirstNewLine -bytes $buffer

        return [PSCustomObject]@{
            HasBOM      = $hasBOM
            NewLineChar = $nlInfo.NewLineChar
            DisplayName = $nlInfo.DisplayName
        }
    }
    catch [System.IO.FileNotFoundException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)] ファイルが見つかりません: $FilePath"
        return $null
    }
    catch [System.UnauthorizedAccessException] {
        Write-Error "[$($MyInvocation.MyCommand.Name)] ファイルにアクセスできません: $FilePath"
        return $null
    }
    catch {
        Write-Error "[$($MyInvocation.MyCommand.Name)] 予期しないエラーが発生しました: $($_.Exception.Message)"
        return $null
    }
}


# 改行コードの取得
function DetectFirstNewLine {
    param([byte[]]$bytes)

    $newLineChar = "`n" # デフォルトは LF
    for ($i = 0; $i -lt $bytes.Length - 1; $i++) {
        if ($bytes[$i] -eq 0x0D -and $bytes[$i + 1] -eq 0x0A) {
            $newLineChar = "`r`n"
            break
        }
        elseif ($bytes[$i] -eq 0x0A) {
            $newLineChar = "`n"
            break
        }
        elseif ($bytes[$i] -eq 0x0D) {
            $newLineChar = "`r"
            break
        }
    }

    switch ($newLineChar) {
        "`r`n" { $displayName = "CRLF" }
        "`n"   { $displayName = "LF" }
        "`r"   { $displayName = "CR" }
        default { $displayName = "Unknown" }
    }

    return [PSCustomObject]@{
        NewLineChar = $newLineChar
        DisplayName = $displayName
    }
}