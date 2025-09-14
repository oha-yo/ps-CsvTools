
function Write-ExecutionHistory {
    param(
        [string]$LogExtension = ".history"
    )

    try {
        # 呼び出し元スクリプトの情報を取得
        $caller = Get-PSCallStack | Select-Object -Skip 1 -First 1
        $callerInvocation = $caller.InvocationInfo

        $scriptPath = $callerInvocation.MyCommand.Path
        $scriptDir  = [System.IO.Path]::GetDirectoryName($scriptPath)
        $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($callerInvocation.MyCommand.Name)

        $logFile = Join-Path $scriptDir ($scriptName + $LogExtension)

        # パラメータを再構築
        $invocationLine = ".\" + [System.IO.Path]::GetFileName($scriptPath) + " " +
        ($callerInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
            $key = $_.Key
            $value = $_.Value

            # スイッチ型（Verboseなど）はキーだけ出力
            if ($key -eq 'Verbose' -and $value -eq $true) {
                "-$key"
            }
            elseif ($value -is [Array]) {
                "-$key $($value -join ',')"
            }
            elseif ($value -is [string] -and $value -match '[\s,"]') {
                "-$key `"$value`""
            }
            else {
                "-$key $value"
            }
        }) -join ' '

        # タイムスタンプ付きで書き込み
        $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
        Add-Content -Path $logFile -Value "$timestamp $invocationLine"
    }
    catch {
        Write-Warning "[$($MyInvocation.MyCommand.Name)] ログファイルへの書き込みに失敗しました: $_"
    }
}