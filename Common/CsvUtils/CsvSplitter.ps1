class CsvSplitter {
    [string]$Pattern

    CsvSplitter([string]$Separator) {
        $escaped = [Regex]::Escape($Separator)
        $this.Pattern = "$escaped(?=(?:[^""]*""[^""]*"")*[^""]*$)"
    }

    [string[]] Split([string]$line) {
        return [regex]::Split($line, $this.Pattern)
    }

    [string[]] SplitAndClean([string]$line) {
        $csvline = $this.Split($line)
        $csvline = $csvline | ForEach-Object {
            $_.Trim() -replace '^"(.*)"$', '$1' -replace '""', '"'
        }
        return $csvline
    }
}