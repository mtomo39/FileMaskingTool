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

Invoke-TestCase 'configuration rejects non-positive sample counts' {
    $config = New-ValidTestConfig
    $config.Jobs[0].SampleCount = 0
    Assert-Throws { Test-MaskingConfig $config } -ExpectedMessage 'SampleCount must be positive'
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

Remove-Item -LiteralPath $script:ConfigTestRoot -Recurse -Force

Write-Output "Tests passed: $script:PassedTests"
Write-Output "Tests failed: $($script:FailedTests.Count)"
if ($script:FailedTests.Count -gt 0) {
    exit 1
}

exit 0
