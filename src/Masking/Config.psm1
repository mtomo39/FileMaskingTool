function Resolve-ConfiguredPath {
    param(
        [string] $Path,
        [string] $BaseDirectory
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $BaseDirectory $Path))
}

function Assert-ConfigEncoding {
    param(
        [string] $Name,
        [string] $Label
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        throw "$Label is required"
    }

    try {
        $codePage = 0
        if ([int]::TryParse($Name, [ref]$codePage)) {
            [void][System.Text.Encoding]::GetEncoding($codePage)
        }
        else {
            [void][System.Text.Encoding]::GetEncoding($Name)
        }
    }
    catch {
        throw "$Label is not a supported .NET encoding"
    }
}

function Get-ObjectPropertyNames {
    param([object] $Value)

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Collections.IDictionary]) {
        return @($Value.Keys | ForEach-Object { [string]$_ })
    }

    return @($Value.PSObject.Properties | ForEach-Object { $_.Name })
}

function Get-ObjectPropertyValue {
    param(
        [object] $Value,
        [string] $Name
    )

    if ($Value -is [System.Collections.IDictionary]) {
        return $Value[$Name]
    }

    $property = $Value.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Test-TrimMode {
    param(
        [string] $Value,
        [string] $Label
    )

    if (@('None', 'Half', 'Full', 'HalfAndFull') -notcontains $Value) {
        throw "$Label must be None, Half, Full, or HalfAndFull"
    }
}

function Assert-MappingReference {
    param(
        [string] $MappingName,
        [string[]] $MappingNames,
        [string] $Label
    )

    if ([string]::IsNullOrWhiteSpace($MappingName) -or $MappingNames -notcontains $MappingName) {
        throw "$Label refers to an unknown mapping"
    }
}

function Test-FieldRules {
    param(
        [object[]] $Fields,
        [string[]] $MappingNames,
        [string] $Label,
        [int] $RecordWidthBytes,
        [bool] $IsFixedWidth
    )

    if (@($Fields).Count -eq 0) {
        throw "$Label must define at least one field"
    }

    foreach ($field in $Fields) {
        if ([string]::IsNullOrWhiteSpace([string]$field.Name)) {
            throw "$Label contains a field without a name"
        }
        Assert-MappingReference ([string]$field.Mapping) $MappingNames "$Label field $($field.Name)"
        Test-TrimMode ([string]$field.TrimSpaces) "$Label field $($field.Name) TrimSpaces"

        if ($IsFixedWidth) {
            if ($field.PSObject.Properties['Segments']) {
                $segments = @($field.Segments)
                if ($segments.Count -lt 2) {
                    throw "$Label composite field $($field.Name) requires at least two segments"
                }
                if ([int]$field.ByteLength -ne (($segments | Measure-Object -Property ByteLength -Sum).Sum)) {
                    throw "$Label composite field $($field.Name) byte length does not match its segments"
                }
                foreach ($segment in $segments) {
                    if ([int]$segment.StartByte -lt 1 -or [int]$segment.ByteLength -lt 1) {
                        throw "$Label composite field $($field.Name) has an invalid segment range"
                    }
                    if (([int]$segment.StartByte - 1 + [int]$segment.ByteLength) -gt $RecordWidthBytes) {
                        throw "$Label composite field $($field.Name) exceeds record width"
                    }
                }
            }
            else {
                if ([int]$field.StartByte -lt 1 -or [int]$field.ByteLength -lt 1) {
                    throw "$Label field $($field.Name) has an invalid byte range"
                }
                if (([int]$field.StartByte - 1 + [int]$field.ByteLength) -gt $RecordWidthBytes) {
                    throw "$Label field $($field.Name) exceeds record width"
                }
                if (@('Half', 'Full') -notcontains [string]$field.PadSpace) {
                    throw "$Label field $($field.Name) PadSpace must be Half or Full"
                }
            }
        }
        else {
            $hasColumnName = -not [string]::IsNullOrWhiteSpace([string]$field.ColumnName)
            $hasColumnIndex = $null -ne $field.PSObject.Properties['ColumnIndex'] -and [int]$field.ColumnIndex -ge 1
            if (-not $hasColumnName -and -not $hasColumnIndex) {
                throw "$Label field $($field.Name) requires a 1-based ColumnIndex or ColumnName"
            }
            if ($hasColumnName -and $hasColumnIndex) {
                throw "$Label field $($field.Name) cannot specify both ColumnIndex and ColumnName"
            }
        }
    }

    if ($IsFixedWidth) {
        $ranges = New-Object 'System.Collections.Generic.List[object]'
        foreach ($field in $Fields) {
            if ($field.PSObject.Properties['Segments']) {
                $segmentNumber = 0
                foreach ($segment in @($field.Segments)) {
                    $segmentNumber++
                    $ranges.Add([pscustomobject]@{
                        Name  = "$($field.Name).$segmentNumber"
                        Start = [int]$segment.StartByte
                        End   = [int]$segment.StartByte + [int]$segment.ByteLength - 1
                    })
                }
            }
            else {
                $ranges.Add([pscustomobject]@{
                    Name  = [string]$field.Name
                    Start = [int]$field.StartByte
                    End   = [int]$field.StartByte + [int]$field.ByteLength - 1
                })
            }
        }

        for ($i = 0; $i -lt $ranges.Count; $i++) {
            for ($j = $i + 1; $j -lt $ranges.Count; $j++) {
                if ($ranges[$i].Start -le $ranges[$j].End -and $ranges[$j].Start -le $ranges[$i].End) {
                    throw "Fixed-width fields overlap: $($ranges[$i].Name) and $($ranges[$j].Name)"
                }
            }
        }
    }
}

function Test-ExcelSheetCoverage {
    param(
        [string[]] $WorkbookSheetNames,
        [object[]] $ConfiguredSheets
    )

    $configuredNames = @{}
    foreach ($sheet in $ConfiguredSheets) {
        $name = [string]$sheet.Name
        if ([string]::IsNullOrWhiteSpace($name)) {
            throw 'Excel sheet configuration contains an empty name'
        }
        if ($configuredNames.ContainsKey($name)) {
            throw "Duplicate Excel sheet configuration: $name"
        }
        $configuredNames[$name] = $true
    }

    $actualNames = @{}
    foreach ($name in $WorkbookSheetNames) {
        $actualNames[[string]$name] = $true
    }

    foreach ($name in $WorkbookSheetNames) {
        if (-not $configuredNames.ContainsKey([string]$name)) {
            throw "Missing configuration for worksheet: $name"
        }
    }
    foreach ($name in $configuredNames.Keys) {
        if (-not $actualNames.ContainsKey([string]$name)) {
            throw "Configured worksheet does not exist: $name"
        }
    }
}

function Test-MaskingConfig {
    param([object] $Config)

    if ($null -eq $Config) {
        throw 'Configuration is empty'
    }

    $mappingNames = Get-ObjectPropertyNames $Config.Mappings
    if ($mappingNames.Count -eq 0) {
        throw 'At least one mapping table is required'
    }

    foreach ($mappingName in $mappingNames) {
        $mapping = Get-ObjectPropertyValue $Config.Mappings $mappingName
        if ([string]::IsNullOrWhiteSpace([string]$mapping.Path) -or -not (Test-Path -LiteralPath $mapping.Path -PathType Leaf)) {
            throw "Mapping file does not exist: $mappingName"
        }
        Assert-ConfigEncoding ([string]$mapping.Encoding) "Mapping $mappingName Encoding"
        Test-TrimMode ([string]$mapping.KeyTrimMode) "Mapping $mappingName KeyTrimMode"
    }

    $jobs = @($Config.Jobs)
    if ($jobs.Count -eq 0) {
        throw 'At least one job is required'
    }

    $outputPaths = @{}
    foreach ($job in $jobs) {
        if (@('FixedWidth', 'Csv', 'Excel') -notcontains [string]$job.Format) {
            throw 'Unsupported format'
        }
        if (@('Sample', 'All') -notcontains [string]$job.Mode) {
            throw 'Unsupported mode'
        }
        if ($job.Mode -eq 'Sample') {
            if ($null -eq $job.SampleCount -or [int]$job.SampleCount -le 0) {
                throw 'SampleCount must be positive in Sample mode'
            }
            if ($null -eq $job.Seed) {
                throw 'Seed is required in Sample mode'
            }
        }

        $inputPath = [System.IO.Path]::GetFullPath([string]$job.InputPath)
        $outputPath = [System.IO.Path]::GetFullPath([string]$job.OutputPath)
        if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) {
            throw 'Input file does not exist'
        }
        if ([string]::Equals($inputPath, $outputPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw 'InputPath and OutputPath must differ'
        }
        $outputDirectory = [System.IO.Path]::GetDirectoryName($outputPath)
        if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
            throw 'Output directory does not exist'
        }
        if ($outputPaths.ContainsKey($outputPath)) {
            throw 'Multiple jobs cannot share an OutputPath'
        }
        $outputPaths[$outputPath] = $true
        Assert-ConfigEncoding ([string]$job.Encoding) 'Job Encoding'

        if ($job.Format -eq 'FixedWidth') {
            $width = [int]$job.Layout.RecordWidthBytes
            if ($width -le 0) {
                throw 'RecordWidthBytes must be positive'
            }
            if (@('None', 'LF', 'CRLF') -notcontains [string]$job.Layout.RecordTerminator) {
                throw 'Unsupported fixed-width record terminator'
            }
            if (
                $job.Layout.PSObject.Properties['OutputTerminator'] -and
                @('None', 'LF', 'CRLF') -notcontains [string]$job.Layout.OutputTerminator
            ) {
                throw 'Unsupported fixed-width output terminator'
            }
            Test-FieldRules @($job.Fields) $mappingNames 'FixedWidth' $width $true
        }
        elseif ($job.Format -eq 'Csv') {
            $layout = $job.Layout
            if ([string]$layout.Delimiter -notmatch '^.$' -or [string]$layout.Quote -notmatch '^.$') {
                throw 'CSV Delimiter and Quote must each be one character'
            }
            if ([string]$layout.Delimiter -ceq [string]$layout.Quote) {
                throw 'CSV Delimiter and Quote must differ'
            }
            if ($layout.Header -isnot [bool]) {
                throw 'CSV Header must be a Boolean'
            }
            if ([string]$layout.ContainerMode -eq 'FixedWidthNullPadded') {
                if ([int]$layout.RecordWidthBytes -le 0) {
                    throw 'CSV RecordWidthBytes must be positive'
                }
                if ([int]$layout.NullPadByte -ne 0) {
                    throw 'CSV NullPadByte must be zero'
                }
                if (@('None', 'LF', 'CRLF') -notcontains [string]$layout.RecordTerminator) {
                    throw 'Unsupported CSV record terminator'
                }
            }
            elseif ([string]$layout.ContainerMode -ne 'Records') {
                throw 'Unsupported CSV ContainerMode'
            }
            Test-FieldRules @($job.Fields) $mappingNames 'CSV' 0 $false
        }
        else {
            $sheets = @($job.Layout.Sheets)
            if ($sheets.Count -eq 0) {
                throw 'Excel job must configure every worksheet'
            }
            $configuredSheetNames = @{}
            foreach ($sheet in $sheets) {
                $sheetName = [string]$sheet.Name
                if ([string]::IsNullOrWhiteSpace($sheetName)) {
                    throw 'Excel sheet configuration contains an empty name'
                }
                if ($configuredSheetNames.ContainsKey($sheetName)) {
                    throw "Duplicate Excel sheet configuration: $sheetName"
                }
                $configuredSheetNames[$sheetName] = $true
                if (@('Data', 'NonPersonal') -notcontains [string]$sheet.Role) {
                    throw "Excel sheet $sheetName Role must be Data or NonPersonal"
                }
                if ($sheet.Role -eq 'Data') {
                    if ([int]$sheet.HeaderRow -lt 1 -or [int]$sheet.FirstDataRow -le [int]$sheet.HeaderRow) {
                        throw "Excel sheet $sheetName has an invalid data range"
                    }
                    Test-FieldRules @($sheet.Fields) $mappingNames "Excel sheet $sheetName" 0 $false
                }
                elseif ($sheet.NonPersonalConfirmed -isnot [bool] -or -not $sheet.NonPersonalConfirmed) {
                    throw "Excel sheet $sheetName requires NonPersonalConfirmed=true"
                }
            }
        }
    }
}

function Import-MaskingConfig {
    param([string] $Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw 'Configuration file does not exist'
    }

    $fullConfigPath = [System.IO.Path]::GetFullPath($Path)
    $baseDirectory = [System.IO.Path]::GetDirectoryName($fullConfigPath)
    $json = [System.IO.File]::ReadAllText($fullConfigPath)
    $config = ConvertFrom-Json -InputObject $json -ErrorAction Stop

    foreach ($mappingName in (Get-ObjectPropertyNames $config.Mappings)) {
        $mapping = Get-ObjectPropertyValue $config.Mappings $mappingName
        if (-not [string]::IsNullOrWhiteSpace([string]$mapping.Path)) {
            $mapping.Path = Resolve-ConfiguredPath ([string]$mapping.Path) $baseDirectory
        }
    }
    foreach ($job in @($config.Jobs)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$job.InputPath)) {
            $job.InputPath = Resolve-ConfiguredPath ([string]$job.InputPath) $baseDirectory
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$job.OutputPath)) {
            $job.OutputPath = Resolve-ConfiguredPath ([string]$job.OutputPath) $baseDirectory
        }
    }

    Test-MaskingConfig $config
    return $config
}

Export-ModuleMember -Function Import-MaskingConfig, Test-MaskingConfig, Test-ExcelSheetCoverage
