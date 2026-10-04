$ErrorActionPreference = 'Stop'

$script:PassedTests = 0
$script:FailedTests = New-Object 'System.Collections.Generic.List[string]'

function Assert-Equal {
    param(
        [object] $Expected,
        [object] $Actual,
        [string] $Message = 'Values differ'
    )

    if ($Expected -cne $Actual) {
        throw "$Message. Expected: [$Expected]; actual: [$Actual]"
    }
}

function Assert-Throws {
    param(
        [scriptblock] $Action,
        [string] $Message = 'Expected an exception',
        [string] $ExpectedMessage
    )

    $actualMessage = $null
    try {
        & $Action
    }
    catch {
        $actualMessage = $_.Exception.Message
    }

    if ($null -eq $actualMessage) {
        throw $Message
    }
    if ($ExpectedMessage -and $actualMessage -notlike "*$ExpectedMessage*") {
        throw "Expected error containing [$ExpectedMessage], got [$actualMessage]"
    }
}

function Invoke-TestCase {
    param(
        [string] $Name,
        [scriptblock] $Body
    )

    try {
        & $Body
        $script:PassedTests++
        Write-Output "PASS: $Name"
    }
    catch {
        $script:FailedTests.Add("$Name - $($_.Exception.Message)")
        Write-Output "FAIL: $Name - $($_.Exception.Message)"
    }
}

function Invoke-PowerShellFile {
    param([string[]] $Arguments)

    $powershellExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $output = & $powershellExe @Arguments 2>&1
    [pscustomobject]@{
        ExitCode = $LASTEXITCODE
        Output   = @($output | ForEach-Object { "$_" })
    }
}

function Import-TestMaskingCore {
    $modulePath = Join-Path $PSScriptRoot '..\src\Masking\CompiledCore.psm1'
    if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
        throw 'Compiled core module is missing'
    }
    Import-Module -Force -ErrorAction Stop -Name $modulePath
    Import-MaskingCore
    $type = $null
    foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
        $type = $assembly.GetType('FileMaskingTool.Core.MaskingCore', $false)
        if ($null -ne $type) { break }
    }
    if ($null -eq $type) { throw 'Compiled masking core type is not loaded' }
    return $type
}

function New-TestMaskingCoreObject {
    param(
        [Type] $CoreType,
        [string] $TypeName
    )

    $type = $CoreType.Assembly.GetType("FileMaskingTool.Core.$TypeName", $false)
    if ($null -eq $type) { throw "Compiled core DTO is missing: $TypeName" }
    return [Activator]::CreateInstance($type)
}

function Invoke-TestCompiledFixedWidth {
    param(
        [Type] $CoreType,
        [object] $Options,
        [System.Collections.IDictionary] $MappingIndexes
    )

    $method = $CoreType.GetMethod(
        'ProcessFixedWidth',
        [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static
    )
    if ($null -eq $method) { throw 'Compiled fixed-width processor is missing' }
    return $method.Invoke($null, [object[]]@($Options, $MappingIndexes))
}

function New-TestCompiledField {
    param(
        [Type] $CoreType,
        [string] $Name,
        [int] $StartByte,
        [int] $ByteLength,
        [string] $MappingName,
        [string] $TrimSpaces = 'None',
        [string] $PadSpace = 'Half',
        [object[]] $Segments = @()
    )

    $field = New-TestMaskingCoreObject $CoreType 'FieldRule'
    $field.Name = $Name
    $field.StartByte = $StartByte
    $field.ByteLength = $ByteLength
    $field.MappingName = $MappingName
    $field.TrimSpaces = $TrimSpaces
    $field.PadSpace = $PadSpace
    $segmentType = $CoreType.Assembly.GetType('FileMaskingTool.Core.SegmentRule', $false)
    $segmentArray = [Array]::CreateInstance($segmentType, $Segments.Count)
    for ($i = 0; $i -lt $Segments.Count; $i++) {
        $segment = [Activator]::CreateInstance($segmentType)
        $segment.StartByte = [int]$Segments[$i].StartByte
        $segment.ByteLength = [int]$Segments[$i].ByteLength
        $segmentArray.SetValue($segment, $i)
    }
    $field.Segments = $segmentArray
    return $field
}

function New-TestCompiledFixedWidthOptions {
    param(
        [Type] $CoreType,
        [string] $InputPath,
        [string] $OutputPath,
        [int] $RecordWidthBytes,
        [string] $EncodingName,
        [object[]] $Fields,
        [byte[]] $InputTerminator = [byte[]]@(),
        [byte[]] $OutputTerminator = [byte[]]@(),
        [string] $Mode = 'All',
        [int] $SampleCount = 1,
        [int] $Seed = 17
    )

    $options = New-TestMaskingCoreObject $CoreType 'FixedWidthOptions'
    $options.InputPath = $InputPath
    $options.OutputPath = $OutputPath
    $options.RecordWidthBytes = $RecordWidthBytes
    $options.InputTerminator = $InputTerminator
    $options.OutputTerminator = $OutputTerminator
    $codePage = 0
    if ([int]::TryParse($EncodingName, [ref]$codePage)) {
        $options.Encoding = [System.Text.Encoding]::GetEncoding($codePage)
    }
    else {
        $options.Encoding = [System.Text.Encoding]::GetEncoding($EncodingName)
    }
    $fieldType = $CoreType.Assembly.GetType('FileMaskingTool.Core.FieldRule', $false)
    $fieldArray = [Array]::CreateInstance($fieldType, $Fields.Count)
    for ($i = 0; $i -lt $Fields.Count; $i++) { $fieldArray.SetValue($Fields[$i], $i) }
    $options.Fields = $fieldArray
    $options.Mode = $Mode
    $options.SampleCount = $SampleCount
    $options.Seed = $Seed
    return $options
}

function Invoke-TestCompiledCsv {
    param(
        [Type] $CoreType,
        [object] $Options,
        [System.Collections.IDictionary] $MappingIndexes
    )

    $method = $CoreType.GetMethod(
        'ProcessCsv',
        [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::Static
    )
    if ($null -eq $method) { throw 'Compiled CSV processor is missing' }
    try {
        return $method.Invoke($null, [object[]]@($Options, $MappingIndexes))
    }
    catch [System.Reflection.TargetInvocationException] {
        if ($null -ne $_.Exception.InnerException) { throw $_.Exception.InnerException.Message }
        throw
    }
}

function New-TestCompiledCsvField {
    param(
        [Type] $CoreType,
        [string] $Name,
        [int] $ColumnIndex,
        [string] $MappingName,
        [string] $TrimSpaces = 'None',
        [string] $ColumnName
    )

    $field = New-TestMaskingCoreObject $CoreType 'CsvFieldRule'
    $field.Name = $Name
    $field.ColumnIndex = $ColumnIndex
    $field.ColumnName = $ColumnName
    $field.MappingName = $MappingName
    $field.TrimSpaces = $TrimSpaces
    return $field
}

function New-TestCompiledCsvOptions {
    param(
        [Type] $CoreType,
        [string] $InputPath,
        [string] $OutputPath,
        [string] $EncodingName,
        [object[]] $Fields,
        [bool] $Header = $false,
        [string] $ContainerMode = 'Records',
        [int] $RecordWidthBytes = 0,
        [byte[]] $RecordTerminator = [byte[]]@(),
        [string] $Mode = 'All',
        [int] $SampleCount = 1,
        [int] $Seed = 17
    )

    $options = New-TestMaskingCoreObject $CoreType 'CsvOptions'
    $options.InputPath = $InputPath
    $options.OutputPath = $OutputPath
    $options.Delimiter = [char]','
    $options.Quote = [char]'"'
    $options.Header = $Header
    $options.ContainerMode = $ContainerMode
    $options.RecordWidthBytes = $RecordWidthBytes
    $options.NullPadByte = [byte]0
    $options.RecordTerminator = $RecordTerminator
    $codePage = 0
    if ([int]::TryParse($EncodingName, [ref]$codePage)) {
        $baseEncoding = [System.Text.Encoding]::GetEncoding($codePage)
    }
    else {
        $baseEncoding = [System.Text.Encoding]::GetEncoding($EncodingName)
    }
    $options.Encoding = [System.Text.Encoding]::GetEncoding(
        $baseEncoding.CodePage,
        (New-Object System.Text.EncoderExceptionFallback),
        (New-Object System.Text.DecoderExceptionFallback)
    )
    $fieldType = $CoreType.Assembly.GetType('FileMaskingTool.Core.CsvFieldRule', $false)
    $fieldArray = [Array]::CreateInstance($fieldType, $Fields.Count)
    for ($i = 0; $i -lt $Fields.Count; $i++) { $fieldArray.SetValue($Fields[$i], $i) }
    $options.Fields = $fieldArray
    $options.Mode = $Mode
    $options.SampleCount = $SampleCount
    $options.Seed = $Seed
    return $options
}

$script:ConfigTestRoot = Join-Path $env:TEMP ('MaskingConfigTests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $script:ConfigTestRoot | Out-Null
$script:ConfigInputPath = Join-Path $script:ConfigTestRoot 'input.dat'
$script:ConfigMappingPath = Join-Path $script:ConfigTestRoot 'member-map.tsv'
[System.IO.File]::WriteAllBytes($script:ConfigInputPath, [byte[]](48,49,50,51,52,53,54,55,56,57))
[System.IO.File]::WriteAllText($script:ConfigMappingPath, "0123456789`t9876543210")

function New-ValidTestConfig {
    $mapping = [pscustomobject]@{
        Path         = $script:ConfigMappingPath
        Encoding     = '932'
        KeyTrimMode  = 'None'
    }
    $field = [pscustomobject]@{
        Name       = 'MemberNumber'
        StartByte  = 1
        ByteLength = 10
        Mapping    = 'MemberNumber'
        TrimSpaces = 'None'
        PadSpace   = 'Half'
    }
    $job = [pscustomobject]@{
        InputPath   = $script:ConfigInputPath
        OutputPath  = Join-Path $script:ConfigTestRoot 'output.dat'
        Format      = 'FixedWidth'
        Mode        = 'Sample'
        SampleCount = 1
        Seed        = 17
        Encoding    = '932'
        Layout      = [pscustomobject]@{
            RecordWidthBytes = 10
            RecordTerminator = 'None'
        }
        Fields      = @($field)
    }

    [pscustomobject]@{
        Mappings = [pscustomobject]@{ MemberNumber = $mapping }
        Jobs     = @($job)
    }
}

function Write-TestConfig {
    param(
        [pscustomobject] $Config,
        [string] $Name = 'config.json'
    )

    $path = Join-Path $script:ConfigTestRoot $Name
    $json = $Config | ConvertTo-Json -Depth 12
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
    $path
}

$configModulePath = Join-Path $PSScriptRoot '..\src\Masking\Config.psm1'
Import-Module -Force -ErrorAction Stop -Name $configModulePath
Import-Module -Force -ErrorAction Stop -Name (Join-Path $PSScriptRoot '..\src\Masking\Mapping.psm1')
Import-Module -Force -ErrorAction Stop -Name (Join-Path $PSScriptRoot '..\src\Masking\Transform.psm1')
Import-Module -Force -ErrorAction Stop -Name (Join-Path $PSScriptRoot '..\src\Masking\CompiledCore.psm1')
Import-Module -Force -ErrorAction Stop -Name (Join-Path $PSScriptRoot '..\src\Masking\FixedWidth.psm1')
Import-Module -Force -ErrorAction Stop -Name (Join-Path $PSScriptRoot '..\src\Masking\Csv.psm1')
Import-Module -Force -ErrorAction Stop -Name (Join-Path $PSScriptRoot '..\src\Masking\Excel.psm1')
Import-Module -Force -ErrorAction Stop -Name (Join-Path $PSScriptRoot '..\src\Masking\Report.psm1')
Import-Module -Force -ErrorAction Stop -Name (Join-Path $PSScriptRoot '..\src\Masking\Runner.psm1')

Invoke-TestCase 'CLI help exits successfully and prints usage' {
    $cliPath = Join-Path $PSScriptRoot '..\MaskingTool.ps1'
    if (-not (Test-Path -LiteralPath $cliPath -PathType Leaf)) {
        throw 'CLI entry point is missing'
    }

    $result = Invoke-PowerShellFile @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $cliPath,
        '-Help'
    )
    Assert-Equal 0 $result.ExitCode 'CLI help exit code'
    if (($result.Output -join "`n") -notmatch '(?i)usage') {
        throw 'CLI help did not print usage'
    }
}

Invoke-TestCase 'all modules import without executing jobs' {
    $moduleNames = @(
        'Config.psm1',
        'Mapping.psm1',
        'Transform.psm1',
        'CompiledCore.psm1',
        'FixedWidth.psm1',
        'Csv.psm1',
        'Excel.psm1',
        'Runner.psm1',
        'Report.psm1'
    )
    $moduleDirectory = Join-Path $PSScriptRoot '..\src\Masking'

    foreach ($moduleName in $moduleNames) {
        $modulePath = Join-Path $moduleDirectory $moduleName
        if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
            throw "Module is missing: $moduleName"
        }

        $escapedPath = $modulePath.Replace("'", "''")
        $command = "Import-Module -Force -ErrorAction Stop -Name '$escapedPath'; Write-Output '__MODULE_IMPORT_OK__'"
        $result = Invoke-PowerShellFile @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-Command', $command
        )
        Assert-Equal 0 $result.ExitCode "Import exit code for $moduleName"
        Assert-Equal '__MODULE_IMPORT_OK__' ($result.Output | Select-Object -Last 1) "Import output for $moduleName"
    }
}

Invoke-TestCase 'valid JSON manifest imports as a validated configuration' {
    $config = New-ValidTestConfig
    $configPath = Write-TestConfig $config 'valid-config.json'
    $loaded = Import-MaskingConfig $configPath
    Assert-Equal 'FixedWidth' $loaded.Jobs[0].Format 'Imported job format'
}

Invoke-TestCase 'validation-only CLI creates no output file' {
    $config = New-ValidTestConfig
    $config.Jobs[0].OutputPath = Join-Path $script:ConfigTestRoot 'validation-only-output.dat'
    $configPath = Write-TestConfig $config 'validation-only.json'

    $result = Invoke-PowerShellFile @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $PSScriptRoot '..\MaskingTool.ps1'),
        '-Config', $configPath,
        '-ValidateOnly'
    )
    Assert-Equal 0 $result.ExitCode 'Validation-only exit code'
    Assert-Equal $false (Test-Path -LiteralPath $config.Jobs[0].OutputPath) 'Validation-only output existence'
}

Invoke-TestCase 'CLI runs configured jobs and returns the aggregate runner status' {
    $config = New-ValidTestConfig
    $config.Jobs[0].Mode = 'All'
    $config.Jobs[0].OutputPath = Join-Path $script:ConfigTestRoot 'cli-run-output.dat'
    $configPath = Write-TestConfig $config 'cli-run.json'

    $result = Invoke-PowerShellFile @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $PSScriptRoot '..\MaskingTool.ps1'),
        '-Config', $configPath
    )
    Assert-Equal 0 $result.ExitCode 'CLI aggregate exit code'
    Assert-Equal '9876543210' ([System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($config.Jobs[0].OutputPath))) 'CLI published masked output'
    if (($result.Output -join "`n") -match '0123456789') { throw 'CLI report contains a source value' }
}

Invoke-TestCase 'CLI returns failure when a configured job cannot be masked' {
    $config = New-ValidTestConfig
    $inputPath = Join-Path $script:ConfigTestRoot 'cli-failure-input.dat'
    [System.IO.File]::WriteAllBytes($inputPath, [System.Text.Encoding]::ASCII.GetBytes('1111111111'))
    $config.Jobs[0].InputPath = $inputPath
    $config.Jobs[0].Mode = 'All'
    $config.Jobs[0].OutputPath = Join-Path $script:ConfigTestRoot 'cli-failure-output.dat'
    $configPath = Write-TestConfig $config 'cli-failure.json'

    $result = Invoke-PowerShellFile @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $PSScriptRoot '..\MaskingTool.ps1'),
        '-Config', $configPath
    )
    Assert-Equal 1 $result.ExitCode 'CLI failed-job exit code'
    Assert-Equal $false (Test-Path -LiteralPath $config.Jobs[0].OutputPath) 'Failed job output existence'
    if (($result.Output -join "`n") -notmatch 'CompletedWithFailures') {
        throw 'CLI report did not include the aggregate failure status'
    }
}

Invoke-TestCase 'CLI help and validation-only bypass core compilation' {
    $isolatedRoot = Join-Path $script:ConfigTestRoot 'invalid-core-validation'
    $maskingSource = Join-Path $PSScriptRoot '..\src\Masking'
    New-Item -ItemType Directory -Path (Join-Path $isolatedRoot 'src') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot '..\MaskingTool.ps1') -Destination $isolatedRoot
    Copy-Item -LiteralPath $maskingSource -Destination (Join-Path $isolatedRoot 'src') -Recurse
    Set-Content -LiteralPath (Join-Path $isolatedRoot 'src\Masking\MaskingCore.cs') -Value 'this is not valid C#' -Encoding ASCII

    $help = Invoke-PowerShellFile @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $isolatedRoot 'MaskingTool.ps1'),
        '-Help'
    )
    Assert-Equal 0 $help.ExitCode 'Invalid core CLI help exit code'

    $config = New-ValidTestConfig
    $config.Jobs[0].OutputPath = Join-Path $script:ConfigTestRoot 'validate-only-no-compile.dat'
    $configPath = Write-TestConfig $config 'validate-only-no-compile.json'
    $validation = Invoke-PowerShellFile @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $isolatedRoot 'MaskingTool.ps1'),
        '-Config', $configPath,
        '-ValidateOnly'
    )
    Assert-Equal 0 $validation.ExitCode 'Invalid core validation-only exit code'
    Assert-Equal $false (Test-Path -LiteralPath $config.Jobs[0].OutputPath) 'Validation-only output existence'
}

Invoke-TestCase 'CLI reports compiled-core load failure as preflight exit 2' {
    $isolatedRoot = Join-Path $script:ConfigTestRoot 'invalid-core-execution'
    $maskingSource = Join-Path $PSScriptRoot '..\src\Masking'
    New-Item -ItemType Directory -Path (Join-Path $isolatedRoot 'src') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot '..\MaskingTool.ps1') -Destination $isolatedRoot
    Copy-Item -LiteralPath $maskingSource -Destination (Join-Path $isolatedRoot 'src') -Recurse
    Set-Content -LiteralPath (Join-Path $isolatedRoot 'src\Masking\MaskingCore.cs') -Value 'this is not valid C#' -Encoding ASCII

    $config = New-ValidTestConfig
    $config.Jobs[0].Mode = 'All'
    $config.Jobs[0].OutputPath = Join-Path $script:ConfigTestRoot 'compiled-core-failure-output.dat'
    $configPath = Write-TestConfig $config 'compiled-core-failure.json'
    $result = Invoke-PowerShellFile @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $isolatedRoot 'MaskingTool.ps1'),
        '-Config', $configPath
    )
    Assert-Equal 2 $result.ExitCode 'Compiled-core preflight exit code'
    if (($result.Output -join "`n") -notmatch 'Compiled masking core could not be loaded') {
        throw 'CLI did not report compiled-core loading failure explicitly'
    }
    Assert-Equal $false (Test-Path -LiteralPath $config.Jobs[0].OutputPath) 'Failed compilation output existence'
}

Invoke-TestCase 'example manifest and operating guide cover every supported format' {
    $examplePath = Join-Path $PSScriptRoot '..\config.example.json'
    $exampleJson = [System.IO.File]::ReadAllText($examplePath)
    $exampleConfig = ConvertFrom-Json -InputObject $exampleJson -ErrorAction Stop
    $formats = @($exampleConfig.Jobs | ForEach-Object { [string]$_.Format } | Select-Object -Unique)
    foreach ($format in @('FixedWidth', 'Csv', 'Excel')) {
        if ($formats -notcontains $format) { throw "Example manifest is missing format $format" }
    }
    if (@($exampleConfig.Jobs | Where-Object { $_.Format -eq 'FixedWidth' -and $_.Layout.RecordWidthBytes -eq 150 -and $_.Layout.RecordTerminator -eq 'CRLF' }).Count -eq 0) {
        throw 'Example manifest is missing the 150-byte CRLF fixed-width format'
    }
    if (@($exampleConfig.Jobs | Where-Object { $_.Format -eq 'Csv' -and $_.Layout.ContainerMode -eq 'FixedWidthNullPadded' }).Count -eq 0) {
        throw 'Example manifest is missing fixed-width NUL-padded CSV'
    }
    if (@($exampleConfig.Jobs | Where-Object { $_.Format -eq 'Excel' -and $_.Layout.Sheets.Count -gt 1 }).Count -eq 0) {
        throw 'Example manifest is missing a multi-sheet Excel configuration'
    }
    if (@($exampleConfig.Jobs | Where-Object { $_.Format -eq 'Excel' -and [System.IO.Path]::GetExtension($_.InputPath) -eq '.xls' }).Count -eq 0 -or
        @($exampleConfig.Jobs | Where-Object { $_.Format -eq 'Excel' -and [System.IO.Path]::GetExtension($_.InputPath) -eq '.xlsx' }).Count -eq 0) {
        throw 'Example manifest must demonstrate both .xls and .xlsx workbooks'
    }
    if (@($exampleConfig.Jobs | Where-Object { $_.Mode -eq 'Sample' }).Count -eq 0 -or
        @($exampleConfig.Jobs | Where-Object { $_.Mode -eq 'All' }).Count -eq 0) {
        throw 'Example manifest must demonstrate Sample and All modes'
    }

    foreach ($mappingName in @($exampleConfig.Mappings.PSObject.Properties | ForEach-Object { $_.Name })) {
        $exampleConfig.Mappings.$mappingName.Path = $script:ConfigMappingPath
    }
    for ($jobIndex = 0; $jobIndex -lt $exampleConfig.Jobs.Count; $jobIndex++) {
        $job = $exampleConfig.Jobs[$jobIndex]
        $extension = [System.IO.Path]::GetExtension([string]$job.InputPath)
        $job.InputPath = Join-Path $script:ConfigTestRoot "example-input-$jobIndex$extension"
        $job.OutputPath = Join-Path $script:ConfigTestRoot "example-output-$jobIndex$extension"
        $bytes = New-Object byte[] ([int]$job.Layout.RecordWidthBytes)
        [System.IO.File]::WriteAllBytes($job.InputPath, $bytes)
    }
    Test-MaskingConfig $exampleConfig

    $readmePath = Join-Path $PSScriptRoot '..\README.md'
    if (-not (Test-Path -LiteralPath $readmePath -PathType Leaf)) { throw 'README.md is missing' }
    $readme = [System.IO.File]::ReadAllText($readmePath)
    foreach ($requiredText in @(
        'no-newline', '150-byte', 'quote state', 'FixedWidthNullPadded',
        'Segments', 'Sample', 'All', '.xls', '.xlsx', 'auxiliary'
    )) {
        if ($readme -notmatch [regex]::Escape($requiredText)) {
            throw "README.md does not document $requiredText"
        }
    }

    $help = Invoke-PowerShellFile @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $PSScriptRoot '..\MaskingTool.ps1'),
        '-Help'
    )
    Assert-Equal 0 $help.ExitCode 'CLI help exit code'
    if (($help.Output -join "`n") -notmatch 'README\.md') {
        throw 'CLI help does not direct users to README.md'
    }
    if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'Measure-Performance.ps1') -PathType Leaf)) {
        throw 'Synthetic benchmark command is missing'
    }
    for ($jobIndex = 0; $jobIndex -lt $exampleConfig.Jobs.Count; $jobIndex++) {
        $job = $exampleConfig.Jobs[$jobIndex]
        $extension = [System.IO.Path]::GetExtension([string]$job.InputPath)
        $job.InputPath = Join-Path $script:ConfigTestRoot "example-cli-input-$jobIndex$extension"
        $job.OutputPath = Join-Path $script:ConfigTestRoot "example-cli-output-$jobIndex$extension"
        $inputSize = if ($job.Format -eq 'FixedWidth') { [int]$job.Layout.RecordWidthBytes } else { 1 }
        [System.IO.File]::WriteAllBytes($job.InputPath, (New-Object byte[] $inputSize))
    }
    $exampleConfigPath = Write-TestConfig $exampleConfig 'example-cli-validate.json'
    $exampleValidation = Invoke-PowerShellFile @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $PSScriptRoot '..\MaskingTool.ps1'),
        '-Config', $exampleConfigPath,
        '-ValidateOnly'
    )
    Assert-Equal 0 $exampleValidation.ExitCode 'Example manifest CLI validation exit code'
    foreach ($job in $exampleConfig.Jobs) {
        Assert-Equal $false (Test-Path -LiteralPath $job.OutputPath) 'Example validation-only output existence'
    }

    $benchmark = Invoke-PowerShellFile @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', (Join-Path $PSScriptRoot 'Measure-Performance.ps1'),
        '-RecordCount', '2'
    )
    Assert-Equal 0 $benchmark.ExitCode 'Small synthetic benchmark exit code'
    if (($benchmark.Output -join "`n") -notmatch 'Synthetic setup: records=2' -or
        ($benchmark.Output -join "`n") -notmatch 'Core compilation:' -or
        ($benchmark.Output -join "`n") -notmatch 'Mapping load:' -or
        ($benchmark.Output -join "`n") -notmatch 'FixedWidth: records=2' -or
        ($benchmark.Output -join "`n") -notmatch 'CSV: records=2' -or
        ($benchmark.Output -join "`n") -notmatch 'Excel separate timing:') {
        throw 'Synthetic benchmark did not report compilation, mapping, processing, and Excel stages'
    }
}

Invoke-TestCase 'configuration rejects a missing input path' {
    $config = New-ValidTestConfig
    $config.Jobs[0].InputPath = Join-Path $script:ConfigTestRoot 'missing.dat'
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'Input file does not exist'
}

Invoke-TestCase 'configuration rejects input and output path collision' {
    $config = New-ValidTestConfig
    $config.Jobs[0].OutputPath = $config.Jobs[0].InputPath
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'InputPath and OutputPath must differ'
}

Invoke-TestCase 'configuration rejects unknown format and mode' {
    $formatConfig = New-ValidTestConfig
    $formatConfig.Jobs[0].Format = 'Unknown'
    Assert-Throws { Test-MaskingConfig $formatConfig } -ExpectedMessage 'Unsupported format'

    $modeConfig = New-ValidTestConfig
    $modeConfig.Jobs[0].Mode = 'Unknown'
    Assert-Throws { Test-MaskingConfig $modeConfig } -ExpectedMessage 'Unsupported mode'
}

Invoke-TestCase 'configuration rejects non-cp932 text and CSV encodings' {
    $fixedConfig = New-ValidTestConfig
    $fixedConfig.Jobs[0].Encoding = 'UTF-8'
    Assert-Throws { Test-MaskingConfig $fixedConfig } -ExpectedMessage 'must be cp932'

    $csvConfig = New-ValidTestConfig
    $csvConfig.Jobs[0].Format = 'Csv'
    $csvConfig.Jobs[0].Fields = @([pscustomobject]@{
        Name = 'MemberNumber'; ColumnIndex = 1; Mapping = 'MemberNumber'; TrimSpaces = 'None'
    })
    $csvConfig.Jobs[0].Layout = [pscustomobject]@{
        Delimiter = ','; Quote = '"'; Header = $false; ContainerMode = 'Records'
    }
    $csvConfig.Jobs[0].Encoding = 'UTF-8'
    Assert-Throws { Test-MaskingConfig $csvConfig } -ExpectedMessage 'must be cp932'
}

Invoke-TestCase 'configuration rejects invalid sample counts and seeds' {
    $config = New-ValidTestConfig
    $config.Jobs[0].SampleCount = 0
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'SampleCount must be positive'

    $config = New-ValidTestConfig
    $config.Jobs[0].Seed = 'not-an-integer'
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'Seed must be a 32-bit integer'
}

Invoke-TestCase 'configuration rejects NUL-padded CSV without explicit padding or compatible encoding' {
    $config = New-ValidTestConfig
    $config.Jobs[0].Format = 'Csv'
    $config.Jobs[0].Fields = @([pscustomobject]@{
        Name = 'MemberNumber'; ColumnIndex = 1; Mapping = 'MemberNumber'; TrimSpaces = 'None'
    })
    $config.Jobs[0].Layout = [pscustomobject]@{
        Delimiter = ','; Quote = '"'; Header = $false
        ContainerMode = 'FixedWidthNullPadded'; RecordWidthBytes = 10; RecordTerminator = 'None'
    }
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'CSV NullPadByte must be zero'

    $config.Jobs[0].Layout | Add-Member -NotePropertyName NullPadByte -NotePropertyValue 0
    $config.Jobs[0].Encoding = 'Unicode'
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'must be cp932'
}

Invoke-TestCase 'configuration rejects overlapping fixed-width fields' {
    $config = New-ValidTestConfig
    $config.Jobs[0].Fields += [pscustomobject]@{
        Name       = 'Overlapping'
        StartByte  = 9
        ByteLength = 1
        Mapping    = 'MemberNumber'
        TrimSpaces = 'None'
        PadSpace   = 'Half'
    }
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'Fixed-width fields overlap'
}

Invoke-TestCase 'configuration rejects fields beyond the record width' {
    $config = New-ValidTestConfig
    $config.Jobs[0].Fields[0].StartByte = 9
    $config.Jobs[0].Fields[0].ByteLength = 3
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'exceeds record width'
}

Invoke-TestCase 'configuration rejects an invalid fixed-width output terminator' {
    $config = New-ValidTestConfig
    $config.Jobs[0].Layout | Add-Member -NotePropertyName OutputTerminator -NotePropertyValue 'BAD'
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'Unsupported fixed-width output terminator'
}

Invoke-TestCase 'Excel sheet coverage rejects unconfigured workbook sheets' {
    $configuredSheets = @(
        [pscustomobject]@{ Name = 'Data'; Role = 'Data' }
    )
    Assert-Throws {
        Test-ExcelSheetCoverage @('Data', 'Hidden') $configuredSheets
    } -ExpectedMessage 'Missing configuration for worksheet'
}

Invoke-TestCase 'mapping table loads 300,000 entries into an ordinal index' {
    $path = Join-Path $script:ConfigTestRoot 'large-map.tsv'
    $writer = New-Object System.IO.StreamWriter($path, $false, [System.Text.Encoding]::UTF8)
    try {
        for ($i = 0; $i -lt 300000; $i++) {
            $writer.WriteLine("$i`tMASK$i")
        }
    }
    finally {
        $writer.Dispose()
    }

    $index = Read-MappingTable $path ([System.Text.Encoding]::UTF8) 'None'
    Assert-Equal 300000 $index.Count 'Mapping index count'
    Assert-Equal 'MASK299999' (Get-MappedValue $index '299999' 'None') 'Last mapping lookup'
    Assert-Equal $true $index.Comparer.Equals('ABC', 'ABC') 'Ordinal comparer identity'
    Assert-Equal $false $index.Comparer.Equals('ABC', 'abc') 'Ordinal comparer case sensitivity'
}

Invoke-TestCase 'mapping table accepts duplicate normalized keys only when replacements agree' {
    $samePath = Join-Path $script:ConfigTestRoot 'same-duplicate.tsv'
    Set-Content -LiteralPath $samePath -Value @(" KEY`tMASK", "KEY `tMASK") -Encoding UTF8
    $index = Read-MappingTable $samePath ([System.Text.Encoding]::UTF8) 'HalfAndFull'
    Assert-Equal 1 $index.Count 'Normalized duplicate count'

    $conflictPath = Join-Path $script:ConfigTestRoot 'conflicting-duplicate.tsv'
    Set-Content -LiteralPath $conflictPath -Value @(" KEY`tMASK1", "KEY `tMASK2") -Encoding UTF8
    Assert-Throws {
        Read-MappingTable $conflictPath ([System.Text.Encoding]::UTF8) 'HalfAndFull'
    } -ExpectedMessage 'Conflicting duplicate mapping key'
}

Invoke-TestCase 'mapping table rejects records without exactly two tab-separated columns' {
    $noTabPath = Join-Path $script:ConfigTestRoot 'no-tab.tsv'
    Set-Content -LiteralPath $noTabPath -Value 'key-only' -Encoding UTF8
    Assert-Throws {
        Read-MappingTable $noTabPath ([System.Text.Encoding]::UTF8) 'None'
    } -ExpectedMessage 'exactly two tab-separated columns'

    $manyTabPath = Join-Path $script:ConfigTestRoot 'many-tabs.tsv'
    Set-Content -LiteralPath $manyTabPath -Value "key`tvalue`tthird" -Encoding UTF8
    Assert-Throws {
        Read-MappingTable $manyTabPath ([System.Text.Encoding]::UTF8) 'None'
    } -ExpectedMessage 'exactly two tab-separated columns'
}

Invoke-TestCase 'mapping lookup trims half- and full-width spaces on keys and input values' {
    $path = Join-Path $script:ConfigTestRoot 'trim-map.tsv'
    $fullSpace = [string][char]0x3000
    Set-Content -LiteralPath $path -Value (" $fullSpace" + "KEY$fullSpace `tMASK") -Encoding UTF8
    $index = Read-MappingTable $path ([System.Text.Encoding]::UTF8) 'HalfAndFull'
    Assert-Equal 'MASK' (Get-MappedValue $index ($fullSpace + ' KEY ' + $fullSpace) 'HalfAndFull') 'Trimmed lookup'
}

Invoke-TestCase 'mapping lookup fails closed for an unknown key' {
    $path = Join-Path $script:ConfigTestRoot 'unknown-map.tsv'
    Set-Content -LiteralPath $path -Value "known`tMASK" -Encoding UTF8
    $index = Read-MappingTable $path ([System.Text.Encoding]::UTF8) 'None'
    Assert-Throws { Get-MappedValue $index 'unknown' 'None' } -ExpectedMessage 'Mapping key not found'
}

Invoke-TestCase 'field transformation pads half-width and full-width spaces by encoded byte count' {
    $path = Join-Path $script:ConfigTestRoot 'padding-map.tsv'
    $changed = [string][char]0x5909
    $swap = [string][char]0x63DB
    $fullSpace = [string][char]0x3000
    Set-Content -LiteralPath $path -Value @(
        "kana`tMASK",
        ("kanji`t" + $changed + $swap),
        "exact`t12345678",
        "short`tA",
        "long`t123456789"
    ) -Encoding UTF8
    $encoding = [System.Text.Encoding]::GetEncoding(932)
    $index = Read-MappingTable $path ([System.Text.Encoding]::UTF8) 'None'

    $halfRule = [pscustomobject]@{ Name = 'Kana'; ByteLength = 8; TrimSpaces = 'HalfAndFull'; PadSpace = 'Half' }
    $halfResult = Convert-FieldValue (" kana" + $fullSpace) $halfRule $index $encoding
    Assert-Equal 8 $halfResult.Length 'Half-space padded byte length'
    Assert-Equal 'MASK    ' $encoding.GetString($halfResult) 'Half-space padded value'

    $fullRule = [pscustomobject]@{ Name = 'Kanji'; ByteLength = 8; TrimSpaces = 'None'; PadSpace = 'Full' }
    $fullResult = Convert-FieldValue 'kanji' $fullRule $index $encoding
    Assert-Equal 8 $fullResult.Length 'Full-space padded byte length'
    Assert-Equal ($changed + $swap + $fullSpace + $fullSpace) $encoding.GetString($fullResult) 'Full-space padded value'

    $exactRule = [pscustomobject]@{ Name = 'Exact'; ByteLength = 8; TrimSpaces = 'None'; PadSpace = 'Half' }
    Assert-Equal 8 (Convert-FieldValue 'exact' $exactRule $index $encoding).Length 'Exact-fit byte length'
}

Invoke-TestCase 'field transformation rejects overflow and a non-divisible full-space remainder' {
    $path = Join-Path $script:ConfigTestRoot 'width-map.tsv'
    Set-Content -LiteralPath $path -Value @(
        "short`tA",
        "long`t123456789"
    ) -Encoding UTF8
    $encoding = [System.Text.Encoding]::GetEncoding(932)
    $index = Read-MappingTable $path ([System.Text.Encoding]::UTF8) 'None'
    $overflowRule = [pscustomobject]@{ Name = 'Overflow'; ByteLength = 8; TrimSpaces = 'None'; PadSpace = 'Half' }
    Assert-Throws {
        Convert-FieldValue 'long' $overflowRule $index $encoding
    } -ExpectedMessage 'exceeds field width'

    $remainderRule = [pscustomobject]@{ Name = 'Remainder'; ByteLength = 2; TrimSpaces = 'None'; PadSpace = 'Full' }
    Assert-Throws {
        Convert-FieldValue 'short' $remainderRule $index $encoding
    } -ExpectedMessage 'cannot be filled exactly'
}

Invoke-TestCase 'compiled core import is idempotent' {
    $modulePath = Join-Path $PSScriptRoot '..\src\Masking\CompiledCore.psm1'
    Import-Module -Force -ErrorAction Stop -Name $modulePath
    Import-MaskingCore
    $firstType = Import-TestMaskingCore
    Import-Module -Force -ErrorAction Stop -Name $modulePath
    Import-MaskingCore
    $secondType = Import-TestMaskingCore
    Assert-Equal $true ([object]::ReferenceEquals($firstType, $secondType)) 'Stable compiled core type identity'
}

Invoke-TestCase 'compiled fixed-width core preserves bytes and terminators' {
    $coreType = Import-TestMaskingCore
    $mappingPath = Join-Path $script:ConfigTestRoot 'compiled-fixed-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "0123456789`t9876543210" -Encoding UTF8
    $mappingIndex = Read-MappingTable $mappingPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ MemberNumber = $mappingIndex }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-fixed-input.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-fixed-output.dat'
    $inputBytes = New-Object byte[] 3000
    for ($i = 0; $i -lt $inputBytes.Length; $i++) { $inputBytes[$i] = 90 }
    [System.Buffer]::BlockCopy([System.Text.Encoding]::ASCII.GetBytes('0123456789'), 0, $inputBytes, 4, 10)
    [System.IO.File]::WriteAllBytes($inputPath, $inputBytes)
    $field = New-TestCompiledField $coreType 'MemberNumber' 5 10 'MemberNumber'
    $options = New-TestCompiledFixedWidthOptions $coreType $inputPath $outputPath 3000 '932' @($field)

    $result = Invoke-TestCompiledFixedWidth $coreType $options $mappingIndexes
    $expected = [byte[]]$inputBytes.Clone()
    [System.Buffer]::BlockCopy([System.Text.Encoding]::ASCII.GetBytes('9876543210'), 0, $expected, 4, 10)
    $actual = [System.IO.File]::ReadAllBytes($outputPath)
    Assert-Equal ([Convert]::ToBase64String($expected)) ([Convert]::ToBase64String($actual)) 'Compiled fixed-width output bytes'
    Assert-Equal 1 $result.RecordsWritten 'Compiled fixed-width record count'
    Assert-Equal 1 $result.FieldsTransformed 'Compiled fixed-width transformed fields'
}

Invoke-TestCase 'compiled fixed-width core maps split member number once' {
    $coreType = Import-TestMaskingCore
    $mappingPath = Join-Path $script:ConfigTestRoot 'compiled-split-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "0123456789`t9876543210" -Encoding UTF8
    $mappingIndex = Read-MappingTable $mappingPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ MemberNumber = $mappingIndex }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-split-input.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-split-output.dat'
    $inputBytes = [System.Text.Encoding]::ASCII.GetBytes(('A' * 4) + '012' + ('B' * 9) + '3456789' + ('C' * 7))
    [System.IO.File]::WriteAllBytes($inputPath, $inputBytes)
    $segments = @(
        [pscustomobject]@{ StartByte = 5; ByteLength = 3 },
        [pscustomobject]@{ StartByte = 17; ByteLength = 7 }
    )
    $field = New-TestCompiledField $coreType 'MemberNumber' 0 10 'MemberNumber' -Segments $segments
    $options = New-TestCompiledFixedWidthOptions $coreType $inputPath $outputPath 30 '932' @($field)

    [void](Invoke-TestCompiledFixedWidth $coreType $options $mappingIndexes)
    $expectedText = ('A' * 4) + '987' + ('B' * 9) + '6543210' + ('C' * 7)
    Assert-Equal $expectedText ([System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($outputPath))) 'Compiled split-field output'
}

Invoke-TestCase 'compiled fixed-width core rejects non-cp932 encodings' {
    $coreType = Import-TestMaskingCore
    $mappingPath = Join-Path $script:ConfigTestRoot 'compiled-split-utf16-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "0123456789`t9876543210" -Encoding UTF8
    $mappingIndex = Read-MappingTable $mappingPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ MemberNumber = $mappingIndex }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-split-utf16-input.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-split-utf16-output.dat'
    $inputBytes = [System.Text.Encoding]::Unicode.GetBytes(('A' * 2) + '012' + ('B' * 2) + '3456789' + ('C' * 2))
    [System.IO.File]::WriteAllBytes($inputPath, $inputBytes)
    $segments = @(
        [pscustomobject]@{ StartByte = 5; ByteLength = 6 },
        [pscustomobject]@{ StartByte = 15; ByteLength = 14 }
    )
    $field = New-TestCompiledField $coreType 'MemberNumber' 0 20 'MemberNumber' -Segments $segments
    $options = New-TestCompiledFixedWidthOptions $coreType $inputPath $outputPath 32 'Unicode' @($field)

    Assert-Throws { Invoke-TestCompiledFixedWidth $coreType $options $mappingIndexes } -ExpectedMessage 'cp932'
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'Non-cp932 fixed-width output is not created'
}

Invoke-TestCase 'compiled fixed-width sampling is deterministic' {
    $coreType = Import-TestMaskingCore
    $mappingPath = Join-Path $script:ConfigTestRoot 'compiled-sample-map.tsv'
    $mappingRows = New-Object 'System.Collections.Generic.List[string]'
    $inputRows = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt 10; $i++) {
        $source = '{0:D10}' -f $i
        $replacement = '{0:D10}' -f (100 + $i)
        $mappingRows.Add($source + "`t" + $replacement)
        $inputRows.Add($source)
    }
    Set-Content -LiteralPath $mappingPath -Value $mappingRows.ToArray() -Encoding UTF8
    $mappingIndex = Read-MappingTable $mappingPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ MemberNumber = $mappingIndex }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-sample-input.dat'
    [System.IO.File]::WriteAllBytes($inputPath, [System.Text.Encoding]::ASCII.GetBytes([string]::Concat($inputRows.ToArray())))
    $field = New-TestCompiledField $coreType 'MemberNumber' 1 10 'MemberNumber'
    $firstPath = Join-Path $script:ConfigTestRoot 'compiled-sample-first.dat'
    $secondPath = Join-Path $script:ConfigTestRoot 'compiled-sample-second.dat'
    $firstOptions = New-TestCompiledFixedWidthOptions $coreType $inputPath $firstPath 10 '932' @($field) -Mode 'Sample' -SampleCount 4 -Seed 17
    $secondOptions = New-TestCompiledFixedWidthOptions $coreType $inputPath $secondPath 10 '932' @($field) -Mode 'Sample' -SampleCount 4 -Seed 17

    $firstResult = Invoke-TestCompiledFixedWidth $coreType $firstOptions $mappingIndexes
    $secondResult = Invoke-TestCompiledFixedWidth $coreType $secondOptions $mappingIndexes
    Assert-Equal 4 $firstResult.RecordsWritten 'Compiled sample row count'
    Assert-Equal ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($firstPath))) ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($secondPath))) 'Compiled sample determinism'
}

Invoke-TestCase 'compiled fixed-width mapping failure does not publish output' {
    [void](Import-TestMaskingCore)
    $mappingPath = Join-Path $script:ConfigTestRoot 'compiled-unknown-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "0123456789`t9876543210" -Encoding UTF8
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-unknown-input.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-unknown-output.dat'
    [System.IO.File]::WriteAllBytes($inputPath, [System.Text.Encoding]::ASCII.GetBytes('1111111111'))
    $config = New-ValidTestConfig
    $config.Jobs[0].InputPath = $inputPath
    $config.Jobs[0].OutputPath = $outputPath
    $config.Jobs[0].Mode = 'All'
    $config.Mappings.MemberNumber.Path = $mappingPath
    [void](Invoke-MaskingRun $config 6>&1)
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'Unknown mapping output was not published'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $script:ConfigTestRoot -Filter 'compiled-unknown-output.dat.masking.*.tmp' -File).Count 'Unknown mapping temporary output cleanup'
}

Invoke-TestCase 'compiled fixed-width output collision preserves existing file' {
    $coreType = Import-TestMaskingCore
    $mappingPath = Join-Path $script:ConfigTestRoot 'compiled-collision-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "0123456789`t9876543210" -Encoding UTF8
    $mappingIndex = Read-MappingTable $mappingPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ MemberNumber = $mappingIndex }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-collision-input.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-collision-output.dat'
    [System.IO.File]::WriteAllBytes($inputPath, [System.Text.Encoding]::ASCII.GetBytes('0123456789'))
    [System.IO.File]::WriteAllBytes($outputPath, [System.Text.Encoding]::ASCII.GetBytes('KEEP-EXISTING'))
    $field = New-TestCompiledField $coreType 'MemberNumber' 1 10 'MemberNumber'
    $options = New-TestCompiledFixedWidthOptions $coreType $inputPath $outputPath 10 '932' @($field)

    Assert-Throws { Invoke-TestCompiledFixedWidth $coreType $options $mappingIndexes } -ExpectedMessage 'already exists'
    Assert-Equal 'KEEP-EXISTING' ([System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($outputPath))) 'Existing output contents'
}

Invoke-TestCase 'compiled CSV preserves quote state and embedded records' {
    $coreType = Import-TestMaskingCore
    $mapPath = Join-Path $script:ConfigTestRoot 'compiled-csv-quote-map.tsv'
    Set-Content -LiteralPath $mapPath -Value @("001`t101", "002`t202") -Encoding UTF8
    $mappingIndexes = @{ Values = (Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-quote-input.csv'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-quote-output.csv'
    $source = '"001","line1' + "`r`n" + 'line ""2"""' + "`r`n" + '002,plain' + "`n"
    [System.IO.File]::WriteAllBytes($inputPath, [System.Text.Encoding]::UTF8.GetBytes($source))
    $field = New-TestCompiledCsvField $coreType 'Identifier' 1 'Values'
    $options = New-TestCompiledCsvOptions $coreType $inputPath $outputPath '932' @($field)

    $result = Invoke-TestCompiledCsv $coreType $options $mappingIndexes
    $actual = [System.Text.Encoding]::GetEncoding(932).GetString([System.IO.File]::ReadAllBytes($outputPath))
    $expected = '"101","line1' + "`r`n" + 'line ""2"""' + "`r`n" + '202,plain' + "`n"
    Assert-Equal $expected $actual 'Compiled CSV exact bytes and quote state'
    Assert-Equal 2 $result.DataRecordsWritten 'Compiled CSV data rows written'
}

Invoke-TestCase 'compiled CSV rejects UTF-8 BOM input under cp932' {
    $coreType = Import-TestMaskingCore
    $mapPath = Join-Path $script:ConfigTestRoot 'compiled-csv-bom-map.tsv'
    Set-Content -LiteralPath $mapPath -Value "001`t101" -Encoding UTF8
    $mappingIndexes = @{ Values = (Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-bom-input.csv'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-bom-output.csv'
    [System.IO.File]::WriteAllBytes($inputPath, [System.Text.UTF8Encoding]::new($true).GetPreamble() + [System.Text.Encoding]::UTF8.GetBytes('001'))
    $field = New-TestCompiledCsvField $coreType 'Identifier' 1 'Values'
    $options = New-TestCompiledCsvOptions $coreType $inputPath $outputPath '932' @($field)
    Assert-Throws { Invoke-TestCompiledCsv $coreType $options $mappingIndexes } -ExpectedMessage 'byte-order mark does not match'
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'UTF-8 BOM output is not created'
}

Invoke-TestCase 'compiled CSV sample is deterministic and keeps header' {
    $coreType = Import-TestMaskingCore
    $mapPath = Join-Path $script:ConfigTestRoot 'compiled-csv-sample-map.tsv'
    $mapRows = New-Object 'System.Collections.Generic.List[string]'
    $mapRows.Add("Identifier`tMasked")
    $sourceRows = New-Object 'System.Collections.Generic.List[string]'
    $sourceRows.Add('Identifier')
    for ($i = 1; $i -le 8; $i++) {
        $value = '{0:D3}' -f $i
        $mapRows.Add($value + "`t" + ('9' + $value.Substring(1)))
        $sourceRows.Add($value)
    }
    Set-Content -LiteralPath $mapPath -Value $mapRows.ToArray() -Encoding UTF8
    $mappingIndexes = @{ Values = (Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-sample-input.csv'
    [System.IO.File]::WriteAllText($inputPath, [string]::Join("`n", $sourceRows.ToArray()) + "`n", [System.Text.Encoding]::GetEncoding(932))
    $field = New-TestCompiledCsvField $coreType 'Identifier' 1 'Values' 'None' 'Identifier'
    $firstPath = Join-Path $script:ConfigTestRoot 'compiled-csv-sample-first.csv'
    $secondPath = Join-Path $script:ConfigTestRoot 'compiled-csv-sample-second.csv'
    $firstOptions = New-TestCompiledCsvOptions $coreType $inputPath $firstPath '932' @($field) -Header $true -Mode 'Sample' -SampleCount 3 -Seed 17
    $secondOptions = New-TestCompiledCsvOptions $coreType $inputPath $secondPath '932' @($field) -Header $true -Mode 'Sample' -SampleCount 3 -Seed 17
    $firstResult = Invoke-TestCompiledCsv $coreType $firstOptions $mappingIndexes
    [void](Invoke-TestCompiledCsv $coreType $secondOptions $mappingIndexes)
    $firstText = [System.Text.Encoding]::GetEncoding(932).GetString([System.IO.File]::ReadAllBytes($firstPath))
    Assert-Equal ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($firstPath))) ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($secondPath))) 'Compiled CSV sample is deterministic'
    if (-not $firstText.StartsWith("Identifier`n", [System.StringComparison]::Ordinal)) { throw 'Compiled CSV sample did not retain the header' }
    Assert-Equal 3 $firstResult.DataRecordsWritten 'Compiled CSV sample data row count'
}

Invoke-TestCase 'compiled CSV sampling reports invalid encoding without input bytes' {
    $coreType = Import-TestMaskingCore
    $mapPath = Join-Path $script:ConfigTestRoot 'compiled-csv-sample-invalid-map.tsv'
    Set-Content -LiteralPath $mapPath -Value "value`tmasked" -Encoding UTF8
    $mappingIndexes = @{ Values = (Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-sample-invalid-input.csv'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-sample-invalid-output.csv'
    [System.IO.File]::WriteAllBytes($inputPath, [byte[]]@(0x82, 0x20))
    $field = New-TestCompiledCsvField $coreType 'Value' 1 'Values'
    $options = New-TestCompiledCsvOptions $coreType $inputPath $outputPath '932' @($field) -Mode 'Sample' -SampleCount 1

    $errorMessage = $null
    try { [void](Invoke-TestCompiledCsv $coreType $options $mappingIndexes) }
    catch { $errorMessage = $_.Exception.Message }
    if ($null -eq $errorMessage) { throw 'Invalid cp932 input did not fail' }
    if ($errorMessage -notmatch 'CSV input contains bytes invalid for the configured encoding') {
        throw "Expected a sanitized encoding error, got [$errorMessage]"
    }
    if ($errorMessage -match '\[82\]') { throw 'Encoding error exposed source byte values' }
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'Invalid sample encoding output existence'
}

Invoke-TestCase 'compiled fixed-width CSV repads the same byte width' {
    $coreType = Import-TestMaskingCore
    $mapPath = Join-Path $script:ConfigTestRoot 'compiled-csv-fixed-map.tsv'
    Set-Content -LiteralPath $mapPath -Value "001`t999" -Encoding UTF8
    $mappingIndexes = @{ Values = (Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-fixed-input.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-fixed-output.dat'
    $width = 12
    $payload = [System.Text.Encoding]::UTF8.GetBytes('001,abc')
    $record = New-Object byte[] $width
    [Array]::Copy($payload, $record, $payload.Length)
    [System.IO.File]::WriteAllBytes($inputPath, $record + [byte[]]@(13, 10))
    $field = New-TestCompiledCsvField $coreType 'Identifier' 1 'Values'
    $options = New-TestCompiledCsvOptions $coreType $inputPath $outputPath '932' @($field) -ContainerMode 'FixedWidthNullPadded' -RecordWidthBytes $width -RecordTerminator ([byte[]]@(13, 10))

    $result = Invoke-TestCompiledCsv $coreType $options $mappingIndexes
    $actual = [System.IO.File]::ReadAllBytes($outputPath)
    Assert-Equal 14 $actual.Length 'Compiled fixed CSV output size'
    Assert-Equal '999,abc' ([System.Text.Encoding]::GetEncoding(932).GetString($actual, 0, 7)) 'Compiled fixed CSV transformed payload'
    Assert-Equal 0 $actual[7] 'Compiled fixed CSV first NUL pad'
    Assert-Equal 0 $actual[11] 'Compiled fixed CSV final NUL pad'
    Assert-Equal 13 $actual[12] 'Compiled fixed CSV CR'
    Assert-Equal 10 $actual[13] 'Compiled fixed CSV LF'
    Assert-Equal 1 $result.DataRecordsWritten 'Compiled fixed CSV row count'
}

Invoke-TestCase 'compiled fixed-width CSV rejects malformed payloads' {
    $coreType = Import-TestMaskingCore
    $mapPath = Join-Path $script:ConfigTestRoot 'compiled-csv-invalid-map.tsv'
    Set-Content -LiteralPath $mapPath -Value "001`t999" -Encoding UTF8
    $mappingIndexes = @{ Values = (Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None') }
    $field = New-TestCompiledCsvField $coreType 'Identifier' 1 'Values'
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-invalid-input.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-invalid-output.dat'
    $block = New-Object byte[] 12
    [Array]::Copy([System.Text.Encoding]::UTF8.GetBytes("001,$([char]0)abc"), $block, 8)
    [System.IO.File]::WriteAllBytes($inputPath, $block)
    $options = New-TestCompiledCsvOptions $coreType $inputPath $outputPath '932' @($field) -ContainerMode 'FixedWidthNullPadded' -RecordWidthBytes 12
    Assert-Throws { Invoke-TestCompiledCsv $coreType $options $mappingIndexes } -ExpectedMessage 'NUL byte occurs inside CSV payload'
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'Interior NUL output is not created'

    [System.IO.File]::WriteAllBytes($inputPath, [System.Text.Encoding]::UTF8.GetBytes('001,abcdefghijkl'))
    Assert-Throws { Invoke-TestCompiledCsv $coreType $options $mappingIndexes } -ExpectedMessage 'fixed-width file length is not a whole number'

    $options.Encoding = [System.Text.Encoding]::Unicode
    [System.IO.File]::WriteAllBytes($inputPath, [byte[]](49,0,0,0))
    Assert-Throws { Invoke-TestCompiledCsv $coreType $options $mappingIndexes } -ExpectedMessage 'cp932'
}

Invoke-TestCase 'compiled CSV fails closed on invalid row or mapping' {
    $coreType = Import-TestMaskingCore
    $mapPath = Join-Path $script:ConfigTestRoot 'compiled-csv-failure-map.tsv'
    Set-Content -LiteralPath $mapPath -Value @("001`tmasked", "002`tmasked") -Encoding UTF8
    $mappingIndexes = @{ Values = (Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None') }
    $field = New-TestCompiledCsvField $coreType 'Identifier' 1 'Values'
    $inputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-failure-input.csv'
    $outputPath = Join-Path $script:ConfigTestRoot 'compiled-csv-failure-output.csv'

    [System.IO.File]::WriteAllText($inputPath, "001`n002,extra`n", [System.Text.Encoding]::GetEncoding(932))
    $options = New-TestCompiledCsvOptions $coreType $inputPath $outputPath '932' @($field)
    Assert-Throws { Invoke-TestCompiledCsv $coreType $options $mappingIndexes } -ExpectedMessage 'inconsistent column count'
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'Invalid row output is not created'

    [System.IO.File]::WriteAllText($inputPath, "003`n", [System.Text.Encoding]::GetEncoding(932))
    Assert-Throws { Invoke-TestCompiledCsv $coreType $options $mappingIndexes } -ExpectedMessage 'Mapping key not found'
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'Unknown mapping output is not created'

    [System.IO.File]::WriteAllText($inputPath, "001`n", [System.Text.Encoding]::GetEncoding(932))
    Set-Content -LiteralPath $mapPath -Value "001`tvalue,needs-quote" -Encoding UTF8
    $mappingIndexes.Values = Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None'
    Assert-Throws { Invoke-TestCompiledCsv $coreType $options $mappingIndexes } -ExpectedMessage 'requires quoting'
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'Quote-state failure output is not created'
}

Invoke-TestCase 'fixed-width handler delegates sample selection to the compiled core' {
    $mappingPath = Join-Path $script:ConfigTestRoot 'fixed-compiled-sample-map.tsv'
    $mappingRows = New-Object 'System.Collections.Generic.List[string]'
    $inputRows = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt 10; $i++) {
        $source = '{0:D10}' -f $i
        $mappingRows.Add($source + "`t" + ('9' + $source.Substring(1)))
        $inputRows.Add($source)
    }
    Set-Content -LiteralPath $mappingPath -Value $mappingRows.ToArray() -Encoding UTF8
    $indexes = @{ MemberNumber = (Read-MappingTable $mappingPath ([System.Text.Encoding]::ASCII) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'fixed-compiled-sample-input.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'fixed-compiled-sample-output.dat'
    [System.IO.File]::WriteAllBytes($inputPath, [System.Text.Encoding]::ASCII.GetBytes([string]::Concat($inputRows.ToArray())))
    $job = [pscustomobject]@{
        InputPath = $inputPath
        Encoding = '932'
        Mode = 'Sample'
        SampleCount = 3
        Seed = 41
        Layout = [pscustomobject]@{ RecordWidthBytes = 10; RecordTerminator = 'None'; OutputTerminator = 'None' }
        Fields = @([pscustomobject]@{
            Name = 'MemberNumber'; StartByte = 1; ByteLength = 10
            Mapping = 'MemberNumber'; TrimSpaces = 'None'; PadSpace = 'Half'
        })
    }

    $result = Invoke-FixedWidthJob $job $indexes $null $outputPath
    Assert-Equal 10 $result.RecordsInSource 'Compiled sample source record count'
    Assert-Equal 3 $result.RecordsRead 'Compiled sample selected record count'
    Assert-Equal 30 ([System.IO.FileInfo]::new($outputPath).Length) 'Compiled sample output size'
}

Invoke-TestCase 'fixed-width handler masks selected 3,000-byte records and preserves FILLER bytes' {
    $mappingPath = Join-Path $script:ConfigTestRoot 'fixed-id-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "1234567890`t0987654321" -Encoding UTF8
    $index = Read-MappingTable $mappingPath ([System.Text.Encoding]::GetEncoding(932)) 'None'
    $indexes = @{ MemberNumber = $index }

    $inputPath = Join-Path $script:ConfigTestRoot 'fixed-3000.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'fixed-3000.masked.dat'
    $inputBytes = New-Object byte[] 6000
    for ($i = 0; $i -lt $inputBytes.Length; $i++) { $inputBytes[$i] = 0x58 }
    $sourceId = [System.Text.Encoding]::ASCII.GetBytes('1234567890')
    [System.Buffer]::BlockCopy($sourceId, 0, $inputBytes, 4, 10)
    [System.Buffer]::BlockCopy($sourceId, 0, $inputBytes, 3004, 10)
    [System.IO.File]::WriteAllBytes($inputPath, $inputBytes)

    $job = [pscustomobject]@{
        InputPath = $inputPath
        OutputPath = $outputPath
        Encoding = '932'
        Layout = [pscustomobject]@{ RecordWidthBytes = 3000; RecordTerminator = 'None' }
        Fields = @([pscustomobject]@{
            Name = 'MemberNumber'; StartByte = 5; ByteLength = 10
            Mapping = 'MemberNumber'; TrimSpaces = 'None'; PadSpace = 'Half'
        })
    }
    $selected = [int[]]@(1)
    $result = Invoke-FixedWidthJob $job $indexes $selected $outputPath
    $actual = [System.IO.File]::ReadAllBytes($outputPath)
    Assert-Equal 3000 $actual.Length 'Selected output width'
    Assert-Equal 1 $result.RecordsWritten 'Selected output record count'
    Assert-Equal 0x58 $actual[0] 'FILLER byte'
    Assert-Equal '0987654321' ([System.Text.Encoding]::ASCII.GetString($actual, 4, 10)) 'Mapped member number'

    $repeatPath = Join-Path $script:ConfigTestRoot 'fixed-3000.repeat.dat'
    [void](Invoke-FixedWidthJob $job $indexes $selected $repeatPath)
    Assert-Equal ([Convert]::ToBase64String($actual)) ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($repeatPath))) 'Repeatable selected output'
}

Invoke-TestCase 'fixed-width handler can add terminators only to masked output' {
    $mappingPath = Join-Path $script:ConfigTestRoot 'output-terminator-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "1234567890`t0987654321" -Encoding UTF8
    $indexes = @{ MemberNumber = (Read-MappingTable $mappingPath ([System.Text.Encoding]::GetEncoding(932)) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'no-terminator-input.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'with-terminator-output.dat'
    $inputBytes = New-Object byte[] 20
    [System.Buffer]::BlockCopy([System.Text.Encoding]::ASCII.GetBytes('1234567890'), 0, $inputBytes, 0, 10)
    [System.Buffer]::BlockCopy([System.Text.Encoding]::ASCII.GetBytes('1234567890'), 0, $inputBytes, 10, 10)
    [System.IO.File]::WriteAllBytes($inputPath, $inputBytes)
    $job = [pscustomobject]@{
        InputPath = $inputPath
        Encoding = '932'
        Layout = [pscustomobject]@{
            RecordWidthBytes = 10
            RecordTerminator = 'None'
            OutputTerminator = 'CRLF'
        }
        Fields = @([pscustomobject]@{
            Name = 'MemberNumber'; StartByte = 1; ByteLength = 10
            Mapping = 'MemberNumber'; TrimSpaces = 'None'; PadSpace = 'Half'
        })
    }

    [void](Invoke-FixedWidthJob $job $indexes ([int[]]@(0, 1)) $outputPath)
    $actual = [System.IO.File]::ReadAllBytes($outputPath)
    Assert-Equal 24 $actual.Length 'Output with inserted CRLF byte length'
    Assert-Equal 13 $actual[10] 'Output CR byte'
    Assert-Equal 10 $actual[11] 'Output LF byte'
    Assert-Equal 13 $actual[22] 'Final output CR byte'
    Assert-Equal 10 $actual[23] 'Final output LF byte'
    Assert-Equal ([Convert]::ToBase64String($inputBytes)) ([Convert]::ToBase64String([System.IO.File]::ReadAllBytes($inputPath))) 'Source remains unchanged'
}

Invoke-TestCase 'fixed-width handler preserves LF and CRLF record terminators' {
    $mappingPath = Join-Path $script:ConfigTestRoot 'terminated-id-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "1234567890`t0987654321" -Encoding UTF8
    $index = Read-MappingTable $mappingPath ([System.Text.Encoding]::GetEncoding(932)) 'None'
    $indexes = @{ MemberNumber = $index }
    $sourceId = [System.Text.Encoding]::ASCII.GetBytes('1234567890')
    $replacement = [System.Text.Encoding]::ASCII.GetBytes('0987654321')

    foreach ($terminatorName in @('LF', 'CRLF')) {
        if ($terminatorName -eq 'LF') { $terminator = [byte[]]@(10) } else { $terminator = [byte[]]@(13, 10) }
        $width = 150
        $stride = $width + $terminator.Length
        $inputBytes = New-Object byte[] ($stride * 2)
        for ($i = 0; $i -lt $inputBytes.Length; $i++) { $inputBytes[$i] = 0x58 }
        [System.Buffer]::BlockCopy($sourceId, 0, $inputBytes, 0, 10)
        [System.Buffer]::BlockCopy($terminator, 0, $inputBytes, $width, $terminator.Length)
        [System.Buffer]::BlockCopy($sourceId, 0, $inputBytes, $stride, 10)
        [System.Buffer]::BlockCopy($terminator, 0, $inputBytes, $stride + $width, $terminator.Length)

        $inputPath = Join-Path $script:ConfigTestRoot "$terminatorName-input.dat"
        $outputPath = Join-Path $script:ConfigTestRoot "$terminatorName-output.dat"
        [System.IO.File]::WriteAllBytes($inputPath, $inputBytes)
        $job = [pscustomobject]@{
            InputPath = $inputPath
            Encoding = '932'
            Layout = [pscustomobject]@{ RecordWidthBytes = $width; RecordTerminator = $terminatorName }
            Fields = @([pscustomobject]@{
                Name = 'MemberNumber'; StartByte = 1; ByteLength = 10
                Mapping = 'MemberNumber'; TrimSpaces = 'None'; PadSpace = 'Half'
            })
        }

        [void](Invoke-FixedWidthJob $job $indexes ([int[]]@(0, 1)) $outputPath)
        $actual = [System.IO.File]::ReadAllBytes($outputPath)
        Assert-Equal $inputBytes.Length $actual.Length "$terminatorName output byte length"
        Assert-Equal '0987654321' ([System.Text.Encoding]::ASCII.GetString($actual, 0, 10)) "$terminatorName first record value"
        Assert-Equal $terminator[0] $actual[$width] "$terminatorName terminator byte"
        Assert-Equal $terminator[$terminator.Length - 1] $actual[$stride - 1] "$terminatorName record stride"
        Assert-Equal ([Convert]::ToBase64String($terminator)) ([Convert]::ToBase64String([byte[]]$actual[$width..($stride - 1)])) "$terminatorName terminator bytes"
        Assert-Equal ([Convert]::ToBase64String($replacement)) ([Convert]::ToBase64String([byte[]]$actual[0..9])) "$terminatorName exact mapped bytes"
    }
}

Invoke-TestCase 'fixed-width handler rejects malformed record sizes and terminators' {
    $mappingPath = Join-Path $script:ConfigTestRoot 'malformed-id-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "1234567890`t0987654321" -Encoding UTF8
    $indexes = @{ MemberNumber = (Read-MappingTable $mappingPath ([System.Text.Encoding]::GetEncoding(932)) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'malformed-fixed.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'malformed-fixed-out.dat'

    [System.IO.File]::WriteAllBytes($inputPath, [byte[]]@(1,2,3,4,5,6,7,8,9,10,11))
    $job = [pscustomobject]@{
        InputPath = $inputPath
        Encoding = '932'
        Layout = [pscustomobject]@{ RecordWidthBytes = 10; RecordTerminator = 'None' }
        Fields = @([pscustomobject]@{
            Name = 'MemberNumber'; StartByte = 1; ByteLength = 10
            Mapping = 'MemberNumber'; TrimSpaces = 'None'; PadSpace = 'Half'
        })
    }
    Assert-Throws {
        Invoke-FixedWidthJob $job $indexes ([int[]]@(0)) $outputPath
    } -ExpectedMessage 'file size is not divisible by record stride'

    [System.IO.File]::WriteAllBytes($inputPath, [byte[]]@(49,50,51,52,53,54,55,56,57,48,0))
    $job.Layout.RecordWidthBytes = 10
    $job.Layout.RecordTerminator = 'LF'
    Assert-Throws {
        Invoke-FixedWidthJob $job $indexes ([int[]]@(0)) $outputPath
    } -ExpectedMessage 'record terminator mismatch'
}

Invoke-TestCase 'split member number maps ten digits once and scatters 3 plus 7 digits' {
    $mappingPath = Join-Path $script:ConfigTestRoot 'split-id-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "1234567890`t0987654321" -Encoding UTF8
    $indexes = @{ MemberNumber = (Read-MappingTable $mappingPath ([System.Text.Encoding]::GetEncoding(932)) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'split-id.dat'
    $outputPath = Join-Path $script:ConfigTestRoot 'split-id.masked.dat'
    $record = New-Object byte[] 40
    for ($i = 0; $i -lt $record.Length; $i++) { $record[$i] = 0x58 }
    [System.Buffer]::BlockCopy([System.Text.Encoding]::ASCII.GetBytes('123'), 0, $record, 0, 3)
    [System.Buffer]::BlockCopy([System.Text.Encoding]::ASCII.GetBytes('4567890'), 0, $record, 20, 7)
    [System.IO.File]::WriteAllBytes($inputPath, $record)

    $job = [pscustomobject]@{
        InputPath = $inputPath
        Encoding = '932'
        Layout = [pscustomobject]@{ RecordWidthBytes = 40; RecordTerminator = 'None' }
        Fields = @([pscustomobject]@{
            Name = 'MemberNumber'; ByteLength = 10; Mapping = 'MemberNumber'
            TrimSpaces = 'None'; Segments = @(
                [pscustomobject]@{ StartByte = 1; ByteLength = 3 },
                [pscustomobject]@{ StartByte = 21; ByteLength = 7 }
            )
        })
    }
    [void](Invoke-FixedWidthJob $job $indexes ([int[]]@(0)) $outputPath)
    $actual = [System.IO.File]::ReadAllBytes($outputPath)
    Assert-Equal '098' ([System.Text.Encoding]::ASCII.GetString($actual, 0, 3)) 'Scattered prefix'
    Assert-Equal '7654321' ([System.Text.Encoding]::ASCII.GetString($actual, 20, 7)) 'Scattered suffix'
    Assert-Equal '0' ([System.Text.Encoding]::ASCII.GetString($actual, 0, 1)) 'Leading zero retention'
    Assert-Equal 0x58 $actual[19] 'Intervening FILLER preservation'
}

Invoke-TestCase 'fixed-width handler rejects overlapping field ranges' {
    $mappingPath = Join-Path $script:ConfigTestRoot 'overlap-id-map.tsv'
    Set-Content -LiteralPath $mappingPath -Value "1234567890`t0987654321" -Encoding UTF8
    $indexes = @{ MemberNumber = (Read-MappingTable $mappingPath ([System.Text.Encoding]::GetEncoding(932)) 'None') }
    $inputPath = Join-Path $script:ConfigTestRoot 'overlap.dat'
    [System.IO.File]::WriteAllBytes($inputPath, (New-Object byte[] 20))
    $job = [pscustomobject]@{
        InputPath = $inputPath
        Encoding = '932'
        Layout = [pscustomobject]@{ RecordWidthBytes = 20; RecordTerminator = 'None' }
        Fields = @(
            [pscustomobject]@{ Name = 'First'; StartByte = 1; ByteLength = 10; Mapping = 'MemberNumber'; TrimSpaces = 'None'; PadSpace = 'Half' },
            [pscustomobject]@{ Name = 'Second'; StartByte = 10; ByteLength = 10; Mapping = 'MemberNumber'; TrimSpaces = 'None'; PadSpace = 'Half' }
        )
    }
    Assert-Throws {
        Invoke-FixedWidthJob $job $indexes ([int[]]@(0)) (Join-Path $script:ConfigTestRoot 'overlap-out.dat')
    } -ExpectedMessage 'Fixed-width fields overlap'
}

Invoke-TestCase 'CSV parser preserves quote states and quoted delimiters, newlines, and escaped quotes' {
    $csvText = '"a,b",plain,"a""b",,"line1' + "`r`n" + 'line2"' + "`r`n" + 'unquoted,"end"' + "`n"
    $reader = [System.IO.StringReader]::new($csvText)
    $records = @(Read-CsvRecords $reader ([char]',') ([char]'"'))
    Assert-Equal 2 $records.Count 'Logical CSV record count'
    Assert-Equal 5 $records[0].Fields.Count 'First record field count'
    Assert-Equal 'a,b' $records[0].Fields[0].Value 'Quoted delimiter value'
    Assert-Equal $true $records[0].Fields[0].WasQuoted 'Quoted delimiter flag'
    Assert-Equal 'plain' $records[0].Fields[1].Value 'Unquoted value'
    Assert-Equal $false $records[0].Fields[1].WasQuoted 'Unquoted flag'
    Assert-Equal 'a"b' $records[0].Fields[2].Value 'Escaped quote value'
    Assert-Equal ('line1' + "`r`n" + 'line2') $records[0].Fields[4].Value 'Quoted newline value'
    Assert-Equal "`r`n" $records[0].Terminator 'CRLF record terminator'
    Assert-Equal "`n" $records[1].Terminator 'LF record terminator'
    $reader.Dispose()
}

Invoke-TestCase 'CSV parser rejects malformed quote placement' {
    foreach ($text in @('"unterminated', 'left"right', '"closed"tail')) {
        $reader = [System.IO.StringReader]::new($text)
        try {
            Assert-Throws { Read-CsvRecords $reader ([char]',') ([char]'"') } -ExpectedMessage 'CSV'
        }
        finally {
            $reader.Dispose()
        }
    }
}

Invoke-TestCase 'CSV writer preserves quote state and rejects newly required quoting' {
    $record = [pscustomobject]@{
        Fields = @(
            [pscustomobject]@{ Value = 'a,b'; WasQuoted = $true },
            [pscustomobject]@{ Value = 'plain'; WasQuoted = $false },
            [pscustomobject]@{ Value = 'a"b'; WasQuoted = $true }
        )
    }
    Assert-Equal '"a,b",plain,"a""b"' (Write-CsvRecord $record ([char]',') ([char]'"')) 'Serialized quote state'

    $unquotedDelimiter = [pscustomobject]@{ Fields = @([pscustomobject]@{ Value = 'needs,quote'; WasQuoted = $false }) }
    Assert-Throws { Write-CsvRecord $unquotedDelimiter ([char]',') ([char]'"') } -ExpectedMessage 'Unquoted CSV field requires quoting'
    $unquotedQuote = [pscustomobject]@{ Fields = @([pscustomobject]@{ Value = 'needs"quote'; WasQuoted = $false }) }
    Assert-Throws { Write-CsvRecord $unquotedQuote ([char]',') ([char]'"') } -ExpectedMessage 'Unquoted CSV field requires quoting'
}

Invoke-TestCase 'CSV job masks configured columns without changing quote state or row terminators' {
    $mapPath = Join-Path $script:ConfigTestRoot 'csv-map.tsv'
    Set-Content -LiteralPath $mapPath -Value @(
        "001`t999",
        "Alice`tMASK, Person",
        "002`t888",
        "Bob`tBOB"
    ) -Encoding UTF8
    $index = Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ Values = $index }
    $inputPath = Join-Path $script:ConfigTestRoot 'ordinary.csv'
    $outputPath = Join-Path $script:ConfigTestRoot 'ordinary.masked.csv'
    $inputText = '"ID","Name"' + "`r`n" + '"001","Alice"' + "`r`n" + '002,Bob'
    [System.IO.File]::WriteAllText($inputPath, $inputText, [System.Text.Encoding]::GetEncoding(932))
    $job = [pscustomobject]@{
        InputPath = $inputPath
        Mode = 'All'
        Encoding = '932'
        Layout = [pscustomobject]@{ Delimiter = ','; Quote = '"'; Header = $true; ContainerMode = 'Records' }
        Fields = @(
            [pscustomobject]@{ Name = 'MemberNumber'; ColumnIndex = 1; Mapping = 'Values'; TrimSpaces = 'None' },
            [pscustomobject]@{ Name = 'Name'; ColumnName = 'Name'; Mapping = 'Values'; TrimSpaces = 'HalfAndFull' }
        )
    }

    $result = Invoke-CsvJob $job $mappingIndexes $outputPath
    $actualText = [System.IO.File]::ReadAllText($outputPath, [System.Text.Encoding]::GetEncoding(932))
    $expectedText = '"ID","Name"' + "`r`n" + '"999","MASK, Person"' + "`r`n" + '888,BOB'
    Assert-Equal $expectedText $actualText 'CSV logical output'
    Assert-Equal 2 $result.RecordsWritten 'CSV data records written'
}

Invoke-TestCase 'CSV job rejects data rows whose column count differs from the header' {
    $mapPath = Join-Path $script:ConfigTestRoot 'header-width-map.tsv'
    Set-Content -LiteralPath $mapPath -Value "x`tX" -Encoding UTF8
    $index = Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ Values = $index }
    $inputPath = Join-Path $script:ConfigTestRoot 'header-width.csv'
    $outputPath = Join-Path $script:ConfigTestRoot 'header-width.masked.csv'
    [System.IO.File]::WriteAllText($inputPath, "First,Second`n" + 'x', [System.Text.Encoding]::GetEncoding(932))
    $job = [pscustomobject]@{
        InputPath = $inputPath; Mode = 'All'; Encoding = '932'
        Layout = [pscustomobject]@{ Delimiter = ','; Quote = '"'; Header = $true; ContainerMode = 'Records' }
        Fields = @([pscustomobject]@{ Name = 'Value'; ColumnIndex = 1; Mapping = 'Values'; TrimSpaces = 'None' })
    }

    Assert-Throws { Invoke-CsvJob $job $mappingIndexes $outputPath } -ExpectedMessage 'inconsistent column count'
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'Invalid-column-count output is not published'
}

Invoke-TestCase 'CSV fixed-width NUL-padded mode re-pads each record to the same byte width' {
    $mapPath = Join-Path $script:ConfigTestRoot 'padded-csv-map.tsv'
    Set-Content -LiteralPath $mapPath -Value @(
        "001`t12345",
        "Alice`tBob",
        "002`t999",
        "Bob`tLonger"
    ) -Encoding UTF8
    $index = Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ Values = $index }
    $inputPath = Join-Path $script:ConfigTestRoot 'padded.csv'
    $outputPath = Join-Path $script:ConfigTestRoot 'padded.masked.csv'
    $width = 32
    $term = [byte[]]@(13, 10)
    $payloads = @('"001",Alice', '002,"Bob"')
    $inputBytes = New-Object byte[] (($width + $term.Length) * $payloads.Count)
    for ($i = 0; $i -lt $payloads.Count; $i++) {
        $payload = [System.Text.Encoding]::GetEncoding(932).GetBytes($payloads[$i])
        [System.Buffer]::BlockCopy($payload, 0, $inputBytes, $i * ($width + $term.Length), $payload.Length)
        [System.Buffer]::BlockCopy($term, 0, $inputBytes, $i * ($width + $term.Length) + $width, $term.Length)
    }
    [System.IO.File]::WriteAllBytes($inputPath, $inputBytes)
    $job = [pscustomobject]@{
        InputPath = $inputPath
        Mode = 'All'
        Encoding = '932'
        Layout = [pscustomobject]@{
            Delimiter = ','
            Quote = '"'
            Header = $false
            ContainerMode = 'FixedWidthNullPadded'
            RecordWidthBytes = $width
            NullPadByte = 0
            RecordTerminator = 'CRLF'
        }
        Fields = @(
            [pscustomobject]@{ Name = 'MemberNumber'; ColumnIndex = 1; Mapping = 'Values'; TrimSpaces = 'None' },
            [pscustomobject]@{ Name = 'Name'; ColumnIndex = 2; Mapping = 'Values'; TrimSpaces = 'None' }
        )
    }

    [void](Invoke-CsvJob $job $mappingIndexes $outputPath)
    $actual = [System.IO.File]::ReadAllBytes($outputPath)
    Assert-Equal $inputBytes.Length $actual.Length 'Fixed-width CSV output length'
    $firstPayloadEnd = [Array]::IndexOf($actual, [byte]0, 0)
    $firstPayload = [System.Text.Encoding]::GetEncoding(932).GetString($actual, 0, $firstPayloadEnd)
    Assert-Equal '"12345",Bob' $firstPayload 'First transformed payload'
    Assert-Equal 0 $actual[$firstPayloadEnd] 'First record NUL padding'
    Assert-Equal 13 $actual[$width] 'First CRLF CR'
    Assert-Equal 10 $actual[$width + 1] 'First CRLF LF'

    $secondStart = $width + $term.Length
    $secondPayloadEnd = [Array]::IndexOf($actual, [byte]0, $secondStart)
    $secondPayload = [System.Text.Encoding]::GetEncoding(932).GetString($actual, $secondStart, $secondPayloadEnd - $secondStart)
    Assert-Equal '999,"Longer"' $secondPayload 'Second transformed payload'
    Assert-Equal 13 $actual[$secondStart + $width] 'Second CRLF CR'
    Assert-Equal 10 $actual[$secondStart + $width + 1] 'Second CRLF LF'
}

Invoke-TestCase 'CSV fixed-width mode rejects output overflow and interior NUL bytes' {
    $mapPath = Join-Path $script:ConfigTestRoot 'overflow-csv-map.tsv'
    Set-Content -LiteralPath $mapPath -Value "x`t12345678" -Encoding UTF8
    $index = Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ Values = $index }
    $inputPath = Join-Path $script:ConfigTestRoot 'overflow.csv'
    $outputPath = Join-Path $script:ConfigTestRoot 'overflow.masked.csv'
    $width = 8
    $block = New-Object byte[] $width
    $payload = [System.Text.Encoding]::GetEncoding(932).GetBytes('x,y')
    [System.Buffer]::BlockCopy($payload, 0, $block, 0, $payload.Length)
    [System.IO.File]::WriteAllBytes($inputPath, $block)
    $job = [pscustomobject]@{
        InputPath = $inputPath
        Mode = 'All'
        Encoding = '932'
        Layout = [pscustomobject]@{
            Delimiter = ','; Quote = '"'; Header = $false
            ContainerMode = 'FixedWidthNullPadded'; RecordWidthBytes = $width
            NullPadByte = 0; RecordTerminator = 'None'
        }
        Fields = @([pscustomobject]@{ Name = 'Value'; ColumnIndex = 1; Mapping = 'Values'; TrimSpaces = 'None' })
    }
    Assert-Throws { Invoke-CsvJob $job $mappingIndexes $outputPath } -ExpectedMessage 'exceeds fixed-width CSV record payload'

    $interior = [byte[]]@(120, 0, 44, 121, 0, 0, 0, 0)
    [System.IO.File]::WriteAllBytes($inputPath, $interior)
    Assert-Throws { Invoke-CsvJob $job $mappingIndexes $outputPath } -ExpectedMessage 'NUL byte occurs inside CSV payload'

    $job.Encoding = 'Unicode'
    Assert-Throws { Invoke-CsvJob $job $mappingIndexes $outputPath } -ExpectedMessage 'must be cp932'
}

Invoke-TestCase 'CSV sample mode is deterministic and retains the header' {
    $mapPath = Join-Path $script:ConfigTestRoot 'sample-csv-map.tsv'
    Set-Content -LiteralPath $mapPath -Value @("a`tA", "b`tB", "c`tC", "d`tD") -Encoding UTF8
    $index = Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ Values = $index }
    $inputPath = Join-Path $script:ConfigTestRoot 'sample.csv'
    $inputText = 'Value' + "`n" + 'a' + "`n" + 'b' + "`n" + 'c' + "`n" + 'd' + "`n"
    [System.IO.File]::WriteAllText($inputPath, $inputText, [System.Text.Encoding]::GetEncoding(932))
    $job = [pscustomobject]@{
        InputPath = $inputPath; Mode = 'Sample'; SampleCount = 2; Seed = 123
        Encoding = '932'
        Layout = [pscustomobject]@{ Delimiter = ','; Quote = '"'; Header = $true; ContainerMode = 'Records' }
        Fields = @([pscustomobject]@{ Name = 'Value'; ColumnIndex = 1; Mapping = 'Values'; TrimSpaces = 'None' })
    }
    $firstPath = Join-Path $script:ConfigTestRoot 'sample-one.csv'
    $secondPath = Join-Path $script:ConfigTestRoot 'sample-two.csv'
    [void](Invoke-CsvJob $job $mappingIndexes $firstPath)
    [void](Invoke-CsvJob $job $mappingIndexes $secondPath)
    $first = [System.IO.File]::ReadAllText($firstPath, [System.Text.Encoding]::GetEncoding(932))
    $second = [System.IO.File]::ReadAllText($secondPath, [System.Text.Encoding]::GetEncoding(932))
    Assert-Equal $first $second 'Deterministic sample content'
    Assert-Equal 3 @($first.Split([char]"`n") | Where-Object { $_.Length -gt 0 }).Count 'Header and sampled row count'
}

Invoke-TestCase 'CSV rejects a byte-order mark that conflicts with the configured encoding' {
    $mapPath = Join-Path $script:ConfigTestRoot 'bom-map.tsv'
    Set-Content -LiteralPath $mapPath -Value "a`tA" -Encoding UTF8
    $index = Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None'
    $mappingIndexes = @{ Values = $index }
    $inputPath = Join-Path $script:ConfigTestRoot 'bom-input.csv'
    $outputPath = Join-Path $script:ConfigTestRoot 'bom-output.csv'
    $content = [System.Text.Encoding]::UTF8.GetBytes('a')
    $inputBytes = New-Object byte[] (3 + $content.Length)
    $inputBytes[0] = 0xEF; $inputBytes[1] = 0xBB; $inputBytes[2] = 0xBF
    [System.Buffer]::BlockCopy($content, 0, $inputBytes, 3, $content.Length)
    [System.IO.File]::WriteAllBytes($inputPath, $inputBytes)
    $job = [pscustomobject]@{
        InputPath = $inputPath; Mode = 'All'; Encoding = '932'
        Layout = [pscustomobject]@{ Delimiter = ','; Quote = '"'; Header = $false; ContainerMode = 'Records' }
        Fields = @([pscustomobject]@{ Name = 'Value'; ColumnIndex = 1; Mapping = 'Values'; TrimSpaces = 'None' })
    }

    Assert-Throws { Invoke-CsvJob $job $mappingIndexes $outputPath } -ExpectedMessage 'byte-order mark does not match'
    Assert-Equal $false (Test-Path -LiteralPath $outputPath) 'Mismatched-encoding output is not created'
}

Invoke-TestCase 'Excel row sampling is deterministic, sorted, and bounded' {
    $first = @(Get-ExcelSampleRowIndexes 20 5 1234)
    $second = @(Get-ExcelSampleRowIndexes 20 5 1234)
    Assert-Equal 5 $first.Count 'Excel selected row count'
    Assert-Equal ($first -join ',') ($second -join ',') 'Excel deterministic selection'
    for ($i = 0; $i -lt $first.Count; $i++) {
        if ($first[$i] -lt 0 -or $first[$i] -ge 20) { throw 'Excel sample index is out of bounds' }
        if ($i -gt 0 -and $first[$i - 1] -ge $first[$i]) { throw 'Excel sample indexes are not unique and sorted' }
    }
    Assert-Equal '0,1,2' (@(Get-ExcelSampleRowIndexes 3 5 1234) -join ',') 'Oversized sample retains all data rows'
}

if (Test-ExcelComAvailable) {
    Invoke-TestCase 'Excel job masks sampled rows, valueizes formulas, preserves sheets, and reports auxiliary content' {
        $mapPath = Join-Path $script:ConfigTestRoot 'excel-map.tsv'
        Set-Content -LiteralPath $mapPath -Value @(
            "001`tMASK-001",
            "002`tMASK-002",
            "003`tMASK-003"
        ) -Encoding UTF8
        $index = Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None'
        $mappingIndexes = @{ Values = $index }
        $inputPath = Join-Path $script:ConfigTestRoot 'excel-input.xlsx'
        $outputPath = Join-Path $script:ConfigTestRoot 'excel-output.xlsx'
        $excel = $null
        $workbook = $null
        $dataSheet = $null
        $notesSheet = $null
        $hyperlink = $null
        $checkExcel = $null
        $checkWorkbook = $null
        $checkData = $null
        $checkNotes = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.EnableEvents = $false
            $workbook = $excel.Workbooks.Add(-4167)
            $dataSheet = $workbook.Worksheets.Item(1)
            $dataSheet.Name = 'Data'
            $notesSheet = $workbook.Worksheets.Add()
            $notesSheet.Name = 'Notes'
            $dataSheet.Cells.Item(1, 1).Value2 = 'MemberNumber'
            $dataSheet.Cells.Item(1, 2).Value2 = 'Calculation'
            $dataSheet.Cells.Item(2, 1).NumberFormat = '@'
            $dataSheet.Cells.Item(2, 1).Value2 = '001'
            $dataSheet.Cells.Item(2, 2).Formula = '=1+1'
            $dataSheet.Cells.Item(3, 1).NumberFormat = '@'
            $dataSheet.Cells.Item(3, 1).Value2 = '002'
            $dataSheet.Cells.Item(3, 2).Formula = '=2+2'
            $dataSheet.Cells.Item(4, 1).NumberFormat = '@'
            $dataSheet.Cells.Item(4, 1).Value2 = '003'
            $dataSheet.Cells.Item(4, 2).Formula = '=3+3'
            $dataSheet.Cells.Item(1, 1).AddComment('synthetic comment')
            $hyperlink = $dataSheet.Hyperlinks.Add($dataSheet.Cells.Item(1, 2), 'https://example.invalid', [Type]::Missing, 'synthetic link', 'synthetic link')
            $notesSheet.Cells.Item(1, 1).Value2 = 'Synthetic non-personal note'
            $notesSheet.Visible = 0
            $workbook.SaveAs($inputPath, 51)
            $workbook.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($notesSheet)
            $notesSheet = $null
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($dataSheet)
            $dataSheet = $null
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
            $workbook = $null

            $sourceHash = [Convert]::ToBase64String([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.IO.File]::ReadAllBytes($inputPath)))
            $job = [pscustomobject]@{
                InputPath = $inputPath; Mode = 'Sample'; SampleCount = 2; Seed = 1234
                Layout = [pscustomobject]@{
                    Sheets = @(
                        [pscustomobject]@{
                            Name = 'Data'; Role = 'Data'; HeaderRow = 1; FirstDataRow = 2
                            Fields = @([pscustomobject]@{
                                Name = 'MemberNumber'; ColumnName = 'MemberNumber'; Mapping = 'Values'; TrimSpaces = 'None'
                            })
                        },
                        [pscustomobject]@{ Name = 'Notes'; Role = 'NonPersonal'; NonPersonalConfirmed = $true }
                    )
                }

            }
            $result = Invoke-ExcelJob $job $mappingIndexes $outputPath
            $dataStats = @($result.Sheets | Where-Object { $_.Name -eq 'Data' })
            Assert-Equal 2 $dataStats[0].RecordsWritten 'Sampled Excel rows written'
            if (-not ($result.Warnings -join ' ' -match '(?i)comment')) { throw 'Expected a value-free comment warning' }
            if (-not ($result.Warnings -join ' ' -match '(?i)hyperlink')) { throw 'Expected a value-free hyperlink warning' }
            $outputHash = [Convert]::ToBase64String([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.IO.File]::ReadAllBytes($inputPath)))
            Assert-Equal $sourceHash $outputHash 'Excel source remains unchanged'

            $checkExcel = New-Object -ComObject Excel.Application
            $checkExcel.Visible = $false
            $checkExcel.DisplayAlerts = $false
            $checkWorkbook = $checkExcel.Workbooks.Open($outputPath, 0, $true)
            $checkData = $checkWorkbook.Worksheets.Item('Data')
            $checkNotes = $checkWorkbook.Worksheets.Item('Notes')
            if ([string]$checkData.Cells.Item(2, 1).Value2 -notlike 'MASK-*') { throw 'Sampled Excel identifier was not masked' }
            Assert-Equal $false ([bool]$checkData.Cells.Item(2, 2).HasFormula) 'Sampled formula converted to a value'
            Assert-Equal 0 ([int]$checkNotes.Visible) 'Hidden non-personal worksheet retained'
            $checkWorkbook.Close($false)
            $checkExcel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($checkNotes)
            $checkNotes = $null
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($checkData)
            $checkData = $null
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($checkWorkbook)
            $checkWorkbook = $null
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($checkExcel)
            $checkExcel = $null
        }
        finally {
            if ($null -ne $checkWorkbook) { $checkWorkbook.Close($false) }
            if ($null -ne $checkExcel) { $checkExcel.Quit() }
            if ($null -ne $workbook) { $workbook.Close($false) }
            if ($null -ne $excel) { $excel.Quit() }
            foreach ($comObject in @($checkNotes, $checkData, $checkWorkbook, $checkExcel, $notesSheet, $dataSheet, $workbook, $excel)) {
                if ($null -ne $comObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
                }
            }
            if ($null -ne $hyperlink -and [System.Runtime.InteropServices.Marshal]::IsComObject($hyperlink)) {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($hyperlink)
            }
        }
    }

    Invoke-TestCase 'Excel job preserves legacy XLS format' {
        $mapPath = Join-Path $script:ConfigTestRoot 'excel-xls-map.tsv'
        Set-Content -LiteralPath $mapPath -Value "001`t999" -Encoding UTF8
        $index = Read-MappingTable $mapPath ([System.Text.Encoding]::UTF8) 'None'
        $mappingIndexes = @{ Values = $index }
        $inputPath = Join-Path $script:ConfigTestRoot 'excel-legacy-input.xls'
        $outputPath = Join-Path $script:ConfigTestRoot 'excel-legacy-output.xls'
        $excel = $null
        $workbook = $null
        $worksheet = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.EnableEvents = $false
            $workbook = $excel.Workbooks.Add(-4167)
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Name = 'Data'
            $worksheet.Cells.Item(1, 1).Value2 = 'MemberNumber'
            $worksheet.Cells.Item(2, 1).NumberFormat = '@'
            $worksheet.Cells.Item(2, 1).Value2 = '001'
            $workbook.SaveAs($inputPath, 56)
            $workbook.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet)
            $worksheet = $null
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
            $workbook = $null

            $job = [pscustomobject]@{
                InputPath = $inputPath; Mode = 'All'; Layout = [pscustomobject]@{
                    Sheets = @([pscustomobject]@{
                        Name = 'Data'; Role = 'Data'; HeaderRow = 1; FirstDataRow = 2
                        Fields = @([pscustomobject]@{
                            Name = 'MemberNumber'; ColumnIndex = 1; Mapping = 'Values'; TrimSpaces = 'None'
                        })
                    })
                }
            }
            $result = Invoke-ExcelJob $job $mappingIndexes $outputPath
            Assert-Equal 1 $result.Sheets[0].RecordsWritten 'Legacy workbook row count'
            Assert-Equal $true (Test-Path -LiteralPath $outputPath -PathType Leaf) 'Legacy workbook output exists'
        }
        finally {
            if ($null -ne $workbook) { $workbook.Close($false) }
            if ($null -ne $excel) { $excel.Quit() }
            foreach ($comObject in @($worksheet, $workbook, $excel)) {
                if ($null -ne $comObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
                }
            }
        }
    }
}
else {
    Write-Output 'BLOCKED: Excel workbook integration test requires registered Office Excel COM; this is not a passing integration test.'
    Invoke-TestCase 'Excel handler reports missing COM prerequisite explicitly' {
        Assert-Throws { Invoke-ExcelJob $null $null (Join-Path $script:ConfigTestRoot 'blocked.xlsx') } -ExpectedMessage 'Excel COM automation is not registered'
    }
}

Invoke-TestCase 'runner isolates jobs, continues after failures, and handles All and Sample modes' {
    $memberMapPath = Join-Path $script:ConfigTestRoot 'runner-member.tsv'
    $valueMapPath = Join-Path $script:ConfigTestRoot 'runner-values.tsv'
    Set-Content -LiteralPath $memberMapPath -Value "1234567890`t0987654321" -Encoding UTF8
    Set-Content -LiteralPath $valueMapPath -Value @("SAFE`tMASKED", "SAFE2`tMASKED2") -Encoding UTF8
    $badInput = Join-Path $script:ConfigTestRoot 'runner-private.csv'
    $fixedInput = Join-Path $script:ConfigTestRoot 'runner-fixed.dat'
    $sampleInput = Join-Path $script:ConfigTestRoot 'runner-sample.csv'
    $badOutput = Join-Path $script:ConfigTestRoot 'runner-private.masked.csv'
    $fixedOutput = Join-Path $script:ConfigTestRoot 'runner-fixed.masked.dat'
    $sampleOutput = Join-Path $script:ConfigTestRoot 'runner-sample.masked.csv'
    [System.IO.File]::WriteAllText($badInput, 'UNMAPPED-PRIVATE-VALUE', [System.Text.Encoding]::GetEncoding(932))
    [System.IO.File]::WriteAllBytes($fixedInput, [System.Text.Encoding]::ASCII.GetBytes('1234567890'))
    [System.IO.File]::WriteAllText($sampleInput, "SAFE`nSAFE2", [System.Text.Encoding]::GetEncoding(932))
    $badHash = [Convert]::ToBase64String([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.IO.File]::ReadAllBytes($badInput)))
    $fixedHash = [Convert]::ToBase64String([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.IO.File]::ReadAllBytes($fixedInput)))
    $sampleHash = [Convert]::ToBase64String([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.IO.File]::ReadAllBytes($sampleInput)))
    $config = [pscustomobject]@{
        Mappings = [pscustomobject]@{
            MemberNumber = [pscustomobject]@{ Path = $memberMapPath; Encoding = 'UTF-8'; KeyTrimMode = 'None' }
            Values = [pscustomobject]@{ Path = $valueMapPath; Encoding = 'UTF-8'; KeyTrimMode = 'None' }
        }
        Jobs = @(
            [pscustomobject]@{
                InputPath = $badInput; OutputPath = $badOutput; Format = 'Csv'; Mode = 'All'; Encoding = '932'
                Layout = [pscustomobject]@{ Delimiter = ','; Quote = '"'; Header = $false; ContainerMode = 'Records' }
                Fields = @([pscustomobject]@{ Name = 'PrivateValue'; ColumnIndex = 1; Mapping = 'Values'; TrimSpaces = 'None' })
            },
            [pscustomobject]@{
                InputPath = $fixedInput; OutputPath = $fixedOutput; Format = 'FixedWidth'; Mode = 'All'; Encoding = '932'
                Layout = [pscustomobject]@{ RecordWidthBytes = 10; RecordTerminator = 'None' }
                Fields = @([pscustomobject]@{
                    Name = 'MemberNumber'; StartByte = 1; ByteLength = 10; Mapping = 'MemberNumber'
                    TrimSpaces = 'None'; PadSpace = 'Half'
                })
            },
            [pscustomobject]@{
                InputPath = $sampleInput; OutputPath = $sampleOutput; Format = 'Csv'; Mode = 'Sample'
                SampleCount = 1; Seed = 22; Encoding = '932'
                Layout = [pscustomobject]@{ Delimiter = ','; Quote = '"'; Header = $false; ContainerMode = 'Records' }
                Fields = @([pscustomobject]@{ Name = 'Value'; ColumnIndex = 1; Mapping = 'Values'; TrimSpaces = 'None' })
            }
        )
    }

    $runOutput = @(Invoke-MaskingRun $config 6>&1)
    $exitCode = [int]@($runOutput | Where-Object { $_ -is [int] })[-1]
    $reportText = @($runOutput | Where-Object { $_ -is [System.Management.Automation.InformationRecord] } |
        ForEach-Object { [string]$_.MessageData }) -join ''
    Assert-Equal 1 $exitCode 'Aggregate partial-failure exit code'
    if ($reportText -match 'UNMAPPED-PRIVATE-VALUE') { throw 'Value-free report contains source data' }
    if ($reportText -notmatch 'Mapping key not found') { throw 'Report omitted the sanitized job failure' }
    Assert-Equal $false (Test-Path -LiteralPath $badOutput) 'Failed job output is not published'
    Assert-Equal $true (Test-Path -LiteralPath $fixedOutput) 'Later successful job is published'
    Assert-Equal $true (Test-Path -LiteralPath $sampleOutput) 'Sample job is published after earlier failure'
    Assert-Equal '0987654321' ([System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($fixedOutput))) 'Fixed-width job output'
    Assert-Equal $badHash ([Convert]::ToBase64String([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.IO.File]::ReadAllBytes($badInput)))) 'Failed CSV source unchanged'
    Assert-Equal $fixedHash ([Convert]::ToBase64String([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.IO.File]::ReadAllBytes($fixedInput)))) 'Fixed-width source unchanged'
    Assert-Equal $sampleHash ([Convert]::ToBase64String([System.Security.Cryptography.SHA256]::Create().ComputeHash([System.IO.File]::ReadAllBytes($sampleInput)))) 'Sample CSV source unchanged'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $script:ConfigTestRoot -Filter '*.masking.*.tmp' -File).Count 'Temporary job output cleanup'
}

Remove-Item -LiteralPath $script:ConfigTestRoot -Recurse -Force

Write-Output "Tests passed: $script:PassedTests"
Write-Output "Tests failed: $($script:FailedTests.Count)"
if ($script:FailedTests.Count -gt 0) {
    exit 1
}

exit 0
