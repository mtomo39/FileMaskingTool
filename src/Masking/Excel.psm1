function Get-ExcelSampleRowIndexes {
    param(
        [int] $DataRowCount,
        [int] $SampleCount,
        [int] $Seed
    )

    if ($DataRowCount -lt 0 -or $SampleCount -le 0) {
        throw 'Excel sample row counts are invalid'
    }
    if ($DataRowCount -eq 0) { return }

    $count = [Math]::Min($DataRowCount, $SampleCount)
    $selected = New-Object 'System.Collections.Generic.List[int]'
    $random = [System.Random]::new($Seed)

    for ($i = 0; $i -lt $DataRowCount; $i++) {
        if ($selected.Count -lt $count) {
            $selected.Add($i)
        }
        else {
            $replacementIndex = $random.Next(0, $i + 1)
            if ($replacementIndex -lt $count) {
                $selected[$replacementIndex] = $i
            }
        }
    }

    $result = [int[]]$selected.ToArray()
    [Array]::Sort($result)
    return $result
}

function Test-ExcelComAvailable {
    return $null -ne [Type]::GetTypeFromProgID('Excel.Application', $false)
}

function Release-ExcelComObject {
    param([object] $Object)

    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Get-ExcelWorksheetBounds {
    param([object] $Worksheet)

    $cells = $null
    $after = $null
    $lastRowCell = $null
    $lastColumnCell = $null
    try {
        $cells = $Worksheet.Cells
        $after = $cells.Item(1, 1)
        $missing = [Type]::Missing
        $lastRowCell = $cells.Find('*', $after, -4123, 2, 1, 2, $false, $missing, $missing)
        $lastColumnCell = $cells.Find('*', $after, -4123, 2, 2, 2, $false, $missing, $missing)
        if ($null -eq $lastRowCell -or $null -eq $lastColumnCell) {
            return [pscustomobject]@{ LastRow = 0; LastColumn = 0 }
        }
        return [pscustomobject]@{
            LastRow = [int]$lastRowCell.Row
            LastColumn = [int]$lastColumnCell.Column
        }
    }
    finally {
        Release-ExcelComObject $lastColumnCell
        Release-ExcelComObject $lastRowCell
        Release-ExcelComObject $after
        Release-ExcelComObject $cells
    }
}

function Convert-ExcelUsedRangesToValues {
    param(
        [object] $Application,
        [object] $Worksheets
    )

    for ($sheetIndex = 1; $sheetIndex -le $Worksheets.Count; $sheetIndex++) {
        $worksheet = $null
        $usedRange = $null
        $pastedRange = $null
        try {
            $worksheet = $Worksheets.Item($sheetIndex)
            $usedRange = $worksheet.UsedRange
            $usedRange.Copy()
            $pastedRange = $usedRange.PasteSpecial(-4163)
            $Application.CutCopyMode = $false
        }
        finally {
            Release-ExcelComObject $pastedRange
            Release-ExcelComObject $usedRange
            Release-ExcelComObject $worksheet
        }
    }
}

function Get-ExcelFieldColumns {
    param(
        [object] $Worksheet,
        [object[]] $Fields,
        [int] $HeaderRow,
        [int] $LastColumn
    )

    if ($LastColumn -lt 1) { throw 'Excel worksheet contains no used columns' }
    $cells = $null
    $startCell = $null
    $endCell = $null
    $headerRange = $null
    try {
        $cells = $Worksheet.Cells
        $startCell = $cells.Item($HeaderRow, 1)
        $endCell = $cells.Item($HeaderRow, $LastColumn)
        $headerRange = $Worksheet.Range($startCell, $endCell)
        $headerValues = $headerRange.Value2
    }
    finally {
        Release-ExcelComObject $headerRange
        Release-ExcelComObject $endCell
        Release-ExcelComObject $startCell
        Release-ExcelComObject $cells
    }

    $columns = New-Object 'System.Collections.Generic.List[object]'
    $usedColumns = New-Object 'System.Collections.Generic.HashSet[int]'
    if ($LastColumn -gt 1) {
        $headerRowLowerBound = $headerValues.GetLowerBound(0)
        $headerColumnLowerBound = $headerValues.GetLowerBound(1)
    }
    foreach ($field in $Fields) {
        if ($field.PSObject.Properties['ColumnName'] -and -not [string]::IsNullOrWhiteSpace([string]$field.ColumnName)) {
            $matches = New-Object 'System.Collections.Generic.List[int]'
            for ($column = 1; $column -le $LastColumn; $column++) {
                if ($LastColumn -eq 1) {
                    $headerValue = $headerValues
                }
                else {
                    $headerColumnIndex = $headerColumnLowerBound + $column - 1
                    $headerValue = $headerValues[$headerRowLowerBound, $headerColumnIndex]
                }
                if ([string]::Equals(
                    [System.Convert]::ToString($headerValue, [System.Globalization.CultureInfo]::InvariantCulture),
                    [string]$field.ColumnName,
                    [System.StringComparison]::Ordinal
                )) {
                    $matches.Add($column)
                }
            }
            if ($matches.Count -ne 1) {
                throw "Excel header must contain exactly one column named $($field.ColumnName)"
            }
            $columnIndex = $matches[0]
        }
        else {
            $columnIndex = [int]$field.ColumnIndex
            if ($columnIndex -lt 1 -or $columnIndex -gt $LastColumn) {
                throw "Excel field $($field.Name) refers to a column outside the used range"
            }
        }
        if (-not $usedColumns.Add($columnIndex)) {
            throw "Excel fields target the same column index $columnIndex"
        }
        $columns.Add([pscustomobject]@{ Rule = $field; ColumnIndex = $columnIndex })
    }
    return ,$columns.ToArray()
}

function Set-ExcelMaskedColumn {
    param(
        [object] $Worksheet,
        [int] $FirstRow,
        [int] $LastRow,
        [object] $Target,
        [object] $MappingIndexes,
        [System.Collections.Generic.HashSet[int]] $SelectedOffsets
    )

    $cells = $null
    $startCell = $null
    $endCell = $null
    $range = $null
    try {
        $cells = $Worksheet.Cells
        $startCell = $cells.Item($FirstRow, [int]$Target.ColumnIndex)
        $endCell = $cells.Item($LastRow, [int]$Target.ColumnIndex)
        $range = $Worksheet.Range($startCell, $endCell)
        $values = $range.Value2
        $rowCount = $LastRow - $FirstRow + 1
        $mappingName = [string]$Target.Rule.Mapping
        if ($null -eq $MappingIndexes -or -not $MappingIndexes.ContainsKey($mappingName)) {
            throw "Excel field $($Target.Rule.Name) refers to an unavailable mapping index"
        }
        $mappingIndex = $MappingIndexes[$mappingName]

        if ($rowCount -eq 1) {
            if ($null -eq $SelectedOffsets -or $SelectedOffsets.Contains(0)) {
                $sourceValue = [System.Convert]::ToString($values, [System.Globalization.CultureInfo]::InvariantCulture)
                [void]($range.Value2 = Get-MappedValue $mappingIndex $sourceValue ([string]$Target.Rule.TrimSpaces))
            }
        }
        else {
            $rowLowerBound = $values.GetLowerBound(0)
            $columnLowerBound = $values.GetLowerBound(1)
            for ($offset = 0; $offset -lt $rowCount; $offset++) {
                if ($null -ne $SelectedOffsets -and -not $SelectedOffsets.Contains($offset)) { continue }
                $valueRowIndex = $rowLowerBound + $offset
                $rawValue = $values[$valueRowIndex, $columnLowerBound]
                $sourceValue = [System.Convert]::ToString(
                    $rawValue,
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
                $values[$valueRowIndex, $columnLowerBound] =
                    Get-MappedValue $mappingIndex $sourceValue ([string]$Target.Rule.TrimSpaces)
            }
            [void]($range.Value2 = $values)
        }
    }
    finally {
        Release-ExcelComObject $range
        Release-ExcelComObject $endCell
        Release-ExcelComObject $startCell
        Release-ExcelComObject $cells
    }
}

function Remove-ExcelRowRange {
    param(
        [object] $Worksheet,
        [int] $FirstRow,
        [int] $LastRow
    )

    if ($LastRow -lt $FirstRow) { return }
    $range = $null
    $entireRows = $null
    try {
        $range = $Worksheet.Range("$FirstRow`:$LastRow")
        $entireRows = $range.EntireRow
        [void]$entireRows.Delete()
    }
    finally {
        Release-ExcelComObject $entireRows
        Release-ExcelComObject $range
    }
}

function Remove-ExcelUnselectedRows {
    param(
        [object] $Worksheet,
        [int] $FirstDataRow,
        [int] $LastRow,
        [int[]] $SelectedIndexes
    )

    $cursor = $LastRow
    for ($index = $SelectedIndexes.Length - 1; $index -ge 0; $index--) {
        $selectedRow = $FirstDataRow + $SelectedIndexes[$index]
        if ($selectedRow -lt $cursor) {
            Remove-ExcelRowRange $Worksheet ($selectedRow + 1) $cursor
        }
        $cursor = $selectedRow - 1
    }
    if ($cursor -ge $FirstDataRow) {
        Remove-ExcelRowRange $Worksheet $FirstDataRow $cursor
    }
}

function Get-ExcelAuxiliaryWarnings {
    param(
        [object] $Worksheet,
        [int] $SheetIndex,
        [System.Collections.Generic.List[string]] $Warnings
    )

    $comments = $null
    $hyperlinks = $null
    $shapes = $null
    try {
        $comments = $Worksheet.Comments
        $hyperlinks = $Worksheet.Hyperlinks
        $shapes = $Worksheet.Shapes
        if ([int]$comments.Count -gt 0) {
            $Warnings.Add("Worksheet $SheetIndex retains $([int]$comments.Count) comment(s)")
        }
        if ([int]$hyperlinks.Count -gt 0) {
            $Warnings.Add("Worksheet $SheetIndex retains $([int]$hyperlinks.Count) hyperlink(s)")
        }
        if ([int]$shapes.Count -gt 0) {
            $Warnings.Add("Worksheet $SheetIndex retains $([int]$shapes.Count) shape(s)")
        }
    }
    finally {
        Release-ExcelComObject $shapes
        Release-ExcelComObject $hyperlinks
        Release-ExcelComObject $comments
    }
}

function Get-ExcelWorksheetConfiguration {
    param(
        [object[]] $ConfiguredSheets,
        [string] $Name
    )

    $matches = @($ConfiguredSheets | Where-Object {
        [string]::Equals([string]$_.Name, $Name, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($matches.Count -ne 1) {
        throw "Excel worksheet does not have exactly one configuration: $Name"
    }
    return $matches[0]
}

function Assert-ExcelSheetCoverage {
    param(
        [string[]] $WorkbookSheetNames,
        [object[]] $ConfiguredSheets
    )

    $configuredNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($sheet in $ConfiguredSheets) {
        if ([string]::IsNullOrWhiteSpace([string]$sheet.Name)) {
            throw 'Excel sheet configuration contains an empty name'
        }
        if (-not $configuredNames.Add([string]$sheet.Name)) {
            throw "Duplicate Excel sheet configuration: $($sheet.Name)"
        }
    }
    foreach ($name in $WorkbookSheetNames) {
        if (-not $configuredNames.Contains([string]$name)) {
            throw "Missing configuration for worksheet: $name"
        }
    }
    foreach ($name in $configuredNames) {
        if ($WorkbookSheetNames -notcontains $name) {
            throw "Configured worksheet does not exist: $name"
        }
    }
}

function Invoke-ExcelJob {
    param(
        [object] $Job,
        [object] $MappingIndexes,
        [string] $OutputPath
    )

    if (-not (Test-ExcelComAvailable)) {
        throw 'Excel COM automation is not registered'
    }
    if ($null -eq $Job -or [string]::IsNullOrWhiteSpace([string]$Job.InputPath) -or
        -not (Test-Path -LiteralPath ([string]$Job.InputPath) -PathType Leaf)) {
        throw 'Excel input file does not exist'
    }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'Excel output path is required' }
    if (@('Sample', 'All') -notcontains [string]$Job.Mode) { throw 'Excel Mode must be Sample or All' }
    if (@('.xls', '.xlsx') -notcontains [System.IO.Path]::GetExtension([string]$Job.InputPath).ToLowerInvariant()) {
        throw 'Excel input must use the .xls or .xlsx extension'
    }

    $inputPath = [System.IO.Path]::GetFullPath([string]$Job.InputPath)
    $outputPath = [System.IO.Path]::GetFullPath($OutputPath)
    if ([string]::Equals($inputPath, $outputPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Excel InputPath and OutputPath must differ'
    }
    $outputDirectory = [System.IO.Path]::GetDirectoryName($outputPath)
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        throw 'Excel output directory does not exist'
    }

    $extension = [System.IO.Path]::GetExtension($inputPath)
    $workingName = [System.IO.Path]::GetFileNameWithoutExtension($outputPath) +
        '.' + [guid]::NewGuid().ToString('N') + $extension
    $workingPath = Join-Path $outputDirectory $workingName
    $application = $null
    $workbooks = $null
    $workbook = $null
    $worksheets = $null
    $validationWorkbook = $null
    $validationWorksheets = $null
    $sheetStats = New-Object 'System.Collections.Generic.List[object]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    $recordsRead = 0
    $recordsWritten = 0
    $expectedFileFormat = if ($extension -ieq '.xls') { 56 } else { 51 }

    try {
        [System.IO.File]::Copy($inputPath, $workingPath, $false)
        $application = New-Object -ComObject Excel.Application
        $application.Visible = $false
        $application.DisplayAlerts = $false
        $application.EnableEvents = $false
        $application.AskToUpdateLinks = $false
        $application.AutomationSecurity = 3
        $application.Calculation = -4135
        $workbooks = $application.Workbooks
        $workbook = $workbooks.Open($workingPath, 0, $false)
        if ([int]$workbook.FileFormat -ne $expectedFileFormat) {
            throw 'Excel workbook format does not match its file extension'
        }

        $application.CalculateFullRebuild()
        $worksheets = $workbook.Worksheets
        $sheetNames = New-Object 'System.Collections.Generic.List[string]'
        for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
            $worksheet = $null
            try {
                $worksheet = $worksheets.Item($sheetIndex)
                $sheetNames.Add([string]$worksheet.Name)
            }
            finally {
                Release-ExcelComObject $worksheet
            }
        }
        Assert-ExcelSheetCoverage $sheetNames.ToArray() @($Job.Layout.Sheets)
        Convert-ExcelUsedRangesToValues $application $worksheets

        for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
            $worksheet = $null
            try {
                $worksheet = $worksheets.Item($sheetIndex)
                $sheetConfiguration = Get-ExcelWorksheetConfiguration @($Job.Layout.Sheets) ([string]$worksheet.Name)
                Get-ExcelAuxiliaryWarnings $worksheet $sheetIndex $warnings
                if ([string]$sheetConfiguration.Role -eq 'NonPersonal') {
                    $sheetStats.Add([pscustomobject]@{
                        Name = [string]$worksheet.Name
                        RecordsRead = 0
                        RecordsWritten = 0
                    })
                    continue
                }

                $headerRow = [int]$sheetConfiguration.HeaderRow
                $firstDataRow = [int]$sheetConfiguration.FirstDataRow
                $bounds = Get-ExcelWorksheetBounds $worksheet
                if ($bounds.LastRow -lt $headerRow) {
                    throw "Excel worksheet $sheetIndex is missing its configured header row"
                }
                $targets = Get-ExcelFieldColumns $worksheet @($sheetConfiguration.Fields) $headerRow $bounds.LastColumn
                $rowCount = [Math]::Max(0, $bounds.LastRow - $firstDataRow + 1)
                $selectedIndexes = $null
                $selectedOffsets = $null
                if ([string]$Job.Mode -eq 'Sample') {
                    $selectedIndexes = [int[]]@(Get-ExcelSampleRowIndexes $rowCount ([int]$Job.SampleCount) ([int]$Job.Seed))
                    $selectedOffsets = New-Object 'System.Collections.Generic.HashSet[int]'
                    foreach ($index in $selectedIndexes) { [void]$selectedOffsets.Add([int]$index) }
                }
                if ($rowCount -gt 0) {
                    foreach ($target in $targets) {
                        Set-ExcelMaskedColumn $worksheet $firstDataRow $bounds.LastRow $target $MappingIndexes $selectedOffsets
                    }
                    if ($null -ne $selectedIndexes) {
                        Remove-ExcelUnselectedRows $worksheet $firstDataRow $bounds.LastRow $selectedIndexes
                        $written = $selectedIndexes.Length
                    }
                    else {
                        $written = $rowCount
                    }
                }
                else {
                    $written = 0
                }
                $recordsRead += $rowCount
                $recordsWritten += $written
                $sheetStats.Add([pscustomobject]@{
                    Name = [string]$worksheet.Name
                    RecordsRead = $rowCount
                    RecordsWritten = $written
                })
            }
            finally {
                Release-ExcelComObject $worksheet
            }
        }

        $warnings.Add('Workbook document properties were retained and not audited')
        $workbook.Save()
        $workbook.Close($false)
        Release-ExcelComObject $worksheets
        $worksheets = $null
        Release-ExcelComObject $workbook
        $workbook = $null
        if (-not (Test-Path -LiteralPath $workingPath -PathType Leaf) -or (Get-Item -LiteralPath $workingPath).Length -le 0) {
            throw 'Excel staged output was not saved'
        }
        $validationWorkbook = $workbooks.Open($workingPath, 0, $true)
        if ([int]$validationWorkbook.FileFormat -ne $expectedFileFormat) {
            throw 'Excel staged output format validation failed'
        }
        $validationWorksheets = $validationWorkbook.Worksheets
        $validatedSheetNames = New-Object 'System.Collections.Generic.List[string]'
        for ($sheetIndex = 1; $sheetIndex -le $validationWorksheets.Count; $sheetIndex++) {
            $worksheet = $null
            try {
                $worksheet = $validationWorksheets.Item($sheetIndex)
                $validatedSheetNames.Add([string]$worksheet.Name)
            }
            finally {
                Release-ExcelComObject $worksheet
            }
        }
        Assert-ExcelSheetCoverage $validatedSheetNames.ToArray() @($Job.Layout.Sheets)
        $validationWorkbook.Close($false)
        Release-ExcelComObject $validationWorksheets
        $validationWorksheets = $null
        Release-ExcelComObject $validationWorkbook
        $validationWorkbook = $null
        Release-ExcelComObject $workbooks
        $workbooks = $null
        $application.Quit()
        Release-ExcelComObject $application
        $application = $null

        if (Test-Path -LiteralPath $outputPath -PathType Leaf) {
            [System.IO.File]::Replace($workingPath, $outputPath, $null)
        }
        else {
            [System.IO.File]::Move($workingPath, $outputPath)
        }

        return [pscustomobject]@{
            RecordsRead = $recordsRead
            RecordsWritten = $recordsWritten
            Sheets = $sheetStats.ToArray()
            Warnings = $warnings.ToArray()
        }
    }
    finally {
        if ($null -ne $validationWorkbook) { $validationWorkbook.Close($false) }
        if ($null -ne $workbook) { $workbook.Close($false) }
        if ($null -ne $application) { $application.Quit() }
        Release-ExcelComObject $validationWorksheets
        Release-ExcelComObject $validationWorkbook
        Release-ExcelComObject $worksheets
        Release-ExcelComObject $workbook
        Release-ExcelComObject $workbooks
        Release-ExcelComObject $application
        if (Test-Path -LiteralPath $workingPath -PathType Leaf) {
            [System.IO.File]::Delete($workingPath)
        }
    }
}

Export-ModuleMember -Function Get-ExcelSampleRowIndexes, Test-ExcelComAvailable, Invoke-ExcelJob
