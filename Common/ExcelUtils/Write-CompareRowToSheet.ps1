function Compare-Columns {
    param(
        [string[]]$Row1,
        [string[]]$Row2,
        [int[]]$CompareIndexes
    )

    $results = @()
    foreach ($i in $CompareIndexes) {
        $val1 = if ($i -le $Row1.Count) { $Row1[$i - 1] } else { "<null>" }
        $val2 = if ($i -le $Row2.Count) { $Row2[$i - 1] } else { "<null>" }
        $results += if ($val1 -eq $val2) { "〇" } else { "×" }
    }
    return $results
}

function Get-DiffColumnText {
    param (
        [string[]]$Values1,
        [string[]]$Values2,
        [int[]]$CompareIndexes
    )

    #$diffColumns = @()
    $diffColumns = [System.Collections.Generic.List[int]]::new()
    foreach ($i in $CompareIndexes) {
        $val1 = if ($i -le $Values1.Count) { $Values1[$i - 1] } else { "<null>" }
        $val2 = if ($i -le $Values2.Count) { $Values2[$i - 1] } else { "<null>" }

        if ($val1 -ne $val2) {
        #if ('"{0}"' -f $val1 -ne '"{0}"' -f $val2) {
            #$diffColumns += $i
            $diffColumns.Add($i)
        }
    }

    if ($diffColumns.Count -gt 0) {
        # List を配列に変換して文字列結合
        return "No. " + ($diffColumns.ToArray() -join ",")
        #return "No. " + ($diffColumns -join ",")
    } else {
        return "無し"
    }
}

function Write-CompareRowToSheet {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Sheet,
        [int]$RowIndex,
        [string[]]$Row1,
        [string[]]$Row2,
        [int[]]$EffectiveColumns,
        [int[]]$KeyItem,
        [int]$LineNo
    )
    $colIndex = 1
    # 行番号
    $Sheet.Cells.Item($RowIndex, $colIndex++).Value = $LineNo
    # キー項目
    foreach ($idx in $KeyItem) {
        $val = if ($idx -le $Row1.Count) { $Row1[$idx - 1] } else { "" }
        $Sheet.Cells.Item($RowIndex, $colIndex++).Value = $val
    }
    # 列比較結果（〇/×）
    $compareResults = Compare-Columns -Row1 $Row1 -Row2 $Row2 -CompareIndexes $EffectiveColumns
    foreach ($res in $compareResults) {
        $Sheet.Cells.Item($RowIndex, $colIndex++).Value = $res
    }
    # 差異列番号
    $diffText = Get-DiffColumnText $Row1 $Row2 $EffectiveColumns
    $Sheet.Cells.Item($RowIndex, $colIndex++).Value = $diffText
}