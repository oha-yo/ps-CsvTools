function Get-WriterEncoding {
    param(
        [string]$EncodingName,
        [bool]$HasBOM
    )
    switch ($Encoding) {
        "utf-8" {
            if ($HasBOM) {
                return [System.Text.Encoding]::UTF8
            } else {
                return New-Object System.Text.UTF8Encoding($false)
            }
        }
        default {
            return [System.Text.Encoding]::GetEncoding($Encoding)
        }
    }
}