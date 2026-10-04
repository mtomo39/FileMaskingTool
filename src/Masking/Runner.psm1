function Get-RunnerEncoding {
    param([string] $Name)

    $codePage = 0
    if ([int]::TryParse($Name, [ref]$codePage)) {
        return [System.Text.Encoding]::GetEncoding($codePage)
    }
    return [System.Text.Encoding]::GetEncoding($Name)
}

function Get-RunnerMappingNames {
    param([object] $Mappings)

    if ($Mappings -is [System.Collections.IDictionary]) {
        return ,@($Mappings.Keys | ForEach-Object { [string]$_ })
    }
    return ,@($Mappings.PSObject.Properties | ForEach-Object { $_.Name })
}

function Get-RunnerMapping {
    param(
        [object] $Mappings,
        [string] $Name
    )

    if ($Mappings -is [System.Collections.IDictionary]) {
        return $Mappings[$Name]
    }
    return $Mappings.PSObject.Properties[$Name].Value
}

function Get-RunnerSensitivePaths {
    param([object] $Config)

    $paths = New-Object 'System.Collections.Generic.List[string]'
    foreach ($mappingName in (Get-RunnerMappingNames $Config.Mappings)) {
        $mapping = Get-RunnerMapping $Config.Mappings $mappingName
        if (-not [string]::IsNullOrWhiteSpace([string]$mapping.Path)) {
            $paths.Add([string]$mapping.Path)
        }
    }
    foreach ($job in @($Config.Jobs)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$job.InputPath)) { $paths.Add([string]$job.InputPath) }
        if (-not [string]::IsNullOrWhiteSpace([string]$job.OutputPath)) { $paths.Add([string]$job.OutputPath) }
    }
    return ,$paths.ToArray()
}

function Import-RunnerMappingIndexes {
    param([object] $Mappings)

    $indexes = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::Ordinal)
    foreach ($mappingName in (Get-RunnerMappingNames $Mappings)) {
        $mapping = Get-RunnerMapping $Mappings $mappingName
        $index = Read-MappingTable ([string]$mapping.Path) (Get-RunnerEncoding ([string]$mapping.Encoding)) ([string]$mapping.KeyTrimMode)
        $indexes.Add([string]$mappingName, $index)
    }
    return $indexes
}

function Get-RunnerFixedRecordCount {
    param([object] $Job)

    $terminatorLength = switch ([string]$Job.Layout.RecordTerminator) {
        'None' { 0 }
        'LF' { 1 }
        'CRLF' { 2 }
        default { throw 'Unsupported fixed-width record terminator' }
    }
    $stride = [long][int]$Job.Layout.RecordWidthBytes + $terminatorLength
    $length = ([System.IO.FileInfo]::new([string]$Job.InputPath)).Length
    if (($length % $stride) -ne 0) { throw 'Fixed-width input size is not a whole number of records' }
    $count = [long]($length / $stride)
    if ($count -gt [int]::MaxValue) { throw 'Fixed-width sample record count exceeds the supported limit' }
    return [int]$count
}

function Get-RunnerSampleIndexes {
    param(
        [int] $RecordCount,
        [int] $SampleCount,
        [int] $Seed
    )

    if ($RecordCount -lt 0 -or $SampleCount -le 0) { throw 'Sample parameters are invalid' }
    if ($RecordCount -eq 0) { return }
    $sampleSize = [Math]::Min($RecordCount, $SampleCount)
    $selected = New-Object 'System.Collections.Generic.List[int]'
    $random = [System.Random]::new($Seed)
    for ($recordIndex = 0; $recordIndex -lt $RecordCount; $recordIndex++) {
        if ($selected.Count -lt $sampleSize) {
            $selected.Add($recordIndex)
        }
        else {
            $replacementIndex = $random.Next(0, $recordIndex + 1)
            if ($replacementIndex -lt $sampleSize) {
                $selected[$replacementIndex] = $recordIndex
            }
        }
    }
    $indexes = [int[]]$selected.ToArray()
    [Array]::Sort($indexes)
    return $indexes
}

function Publish-RunnerOutput {
    param(
        [string] $TemporaryPath,
        [string] $OutputPath
    )

    if (-not (Test-Path -LiteralPath $TemporaryPath -PathType Leaf)) {
        throw 'Format handler did not create its staged output'
    }
    if (Test-Path -LiteralPath $OutputPath -PathType Leaf) {
        [System.IO.File]::Replace($TemporaryPath, $OutputPath, $null)
    }
    else {
        [System.IO.File]::Move($TemporaryPath, $OutputPath)
    }
}

function Invoke-MaskingRun {
    param([object] $Config)

    $jobResults = New-Object 'System.Collections.Generic.List[object]'
    $sensitivePaths = Get-RunnerSensitivePaths $Config
    try {
        $mappingIndexes = Import-RunnerMappingIndexes $Config.Mappings
    }
    catch {
        $jobResults.Add([pscustomobject]@{
            Job = 0
            Format = 'Preflight'
            Status = 'Failed'
            RecordsRead = 0
            RecordsWritten = 0
            ElapsedMilliseconds = 0
            Error = ConvertTo-ValueFreeErrorMessage $_.Exception $sensitivePaths
            Warnings = @()
        })
        Write-MaskingRunReport $jobResults.ToArray() 2 'PreflightFailed'
        return 2
    }

    for ($jobIndex = 0; $jobIndex -lt @($Config.Jobs).Count; $jobIndex++) {
        $job = @($Config.Jobs)[$jobIndex]
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $outputDirectory = [System.IO.Path]::GetDirectoryName([string]$job.OutputPath)
        $temporaryName = '.' + [System.IO.Path]::GetFileName([string]$job.OutputPath) +
            '.masking.' + [guid]::NewGuid().ToString('N') + '.tmp'
        $temporaryPath = Join-Path $outputDirectory $temporaryName
        try {
            $handlerResult = $null
            switch ([string]$job.Format) {
                'FixedWidth' {
                    $handlerResult = Invoke-FixedWidthJob $job $mappingIndexes $null $temporaryPath
                }
                'Csv' {
                    $handlerResult = Invoke-CsvJob $job $mappingIndexes $temporaryPath
                }
                'Excel' {
                    $handlerResult = Invoke-ExcelJob $job $mappingIndexes $temporaryPath
                }
                default {
                    throw 'Unsupported job format'
                }
            }
            Publish-RunnerOutput $temporaryPath ([string]$job.OutputPath)
            $stopwatch.Stop()
            $warnings = @()
            if ($null -ne $handlerResult.Warnings) { $warnings = @($handlerResult.Warnings) }
            $jobResults.Add([pscustomobject]@{
                Job = $jobIndex + 1
                Format = [string]$job.Format
                Status = 'Succeeded'
                RecordsRead = [long]$handlerResult.RecordsRead
                RecordsWritten = [long]$handlerResult.RecordsWritten
                ElapsedMilliseconds = $stopwatch.ElapsedMilliseconds
                Error = $null
                Warnings = $warnings
            })
        }
        catch {
            $stopwatch.Stop()
            $jobResults.Add([pscustomobject]@{
                Job = $jobIndex + 1
                Format = [string]$job.Format
                Status = 'Failed'
                RecordsRead = 0
                RecordsWritten = 0
                ElapsedMilliseconds = $stopwatch.ElapsedMilliseconds
                Error = ConvertTo-ValueFreeErrorMessage $_.Exception ($sensitivePaths + @($temporaryPath))
                Warnings = @()
            })
        }
        finally {
            if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
                [System.IO.File]::Delete($temporaryPath)
            }
        }
    }

    $failedJobs = @($jobResults | Where-Object { $_.Status -eq 'Failed' }).Count
    $exitCode = if ($failedJobs -gt 0) { 1 } else { 0 }
    $status = if ($failedJobs -gt 0) { 'CompletedWithFailures' } else { 'Succeeded' }
    Write-MaskingRunReport $jobResults.ToArray() $exitCode $status
    return $exitCode
}

Export-ModuleMember -Function Invoke-MaskingRun
