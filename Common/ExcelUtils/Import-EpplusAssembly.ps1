# EPPlus.dllを PowerShell に読み込んで、型情報（.NETクラス）を使えるようにする。
function Import-EpplusAssembly {
    param(
        [Parameter(Mandatory)]
        [string]$DllPath
    )
    try {
        if (-not (Test-Path -Path $DllPath -PathType Leaf)) {
            throw "DLL が見つかりません: $DllPath"
        }
        if (-not ([AppDomain]::CurrentDomain.GetAssemblies().Location -contains (Resolve-Path $DllPath))) {
            Add-Type -Path $DllPath
            [Reflection.Assembly]::LoadFrom($DllPath) | Out-Null
        }
        return $true
    }
    catch {
        Write-Error "[$($MyInvocation.MyCommand.Name)] DLLの読み込みに失敗しました: $($_.Exception.Message)"
        return $false
    }
}