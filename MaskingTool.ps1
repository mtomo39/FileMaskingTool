[CmdletBinding()]
param(
    [string] $Config,
    [switch] $ValidateOnly,
    [switch] $Help
)

if ($Help) {
    Write-Output 'Usage: .\MaskingTool.ps1 -Config <path> [-ValidateOnly]'
    Write-Output 'See README.md for supported formats, configuration examples, and operating guidance.'
    exit 0
}

if ([string]::IsNullOrWhiteSpace($Config)) {
    Write-Error 'Specify -Config <path>.'
    exit 2
}

$configModulePath = Join-Path $PSScriptRoot 'src\Masking\Config.psm1'
Import-Module -Force -ErrorAction Stop -Name $configModulePath

try {
    $validatedConfig = Import-MaskingConfig -Path $Config
}
catch {
    Write-Error "Configuration invalid: $($_.Exception.Message)"
    exit 2
}

if ($ValidateOnly) {
    Write-Output "Configuration valid: $(@($validatedConfig.Jobs).Count) job(s)."
    exit 0
}

$executionModules = @(
    'Mapping.psm1',
    'Transform.psm1',
    'CompiledCore.psm1',
    'FixedWidth.psm1',
    'Csv.psm1',
    'Excel.psm1',
    'Report.psm1',
    'Runner.psm1'
)
try {
    foreach ($moduleName in $executionModules) {
        $modulePath = Join-Path $PSScriptRoot ('src\Masking\' + $moduleName)
        Import-Module -Force -ErrorAction Stop -Name $modulePath
    }
}
catch {
    Write-Error "Execution modules could not be loaded: $($_.Exception.Message)"
    exit 2
}

try {
    Import-MaskingCore
}
catch {
    $jobResult = [pscustomobject]@{
        Job = 0
        Format = 'Preflight'
        Status = 'Failed'
        RecordsRead = 0
        RecordsWritten = 0
        ElapsedMilliseconds = 0
        Error = ConvertTo-ValueFreeErrorMessage $_.Exception @()
        Warnings = @()
    }
    Write-MaskingRunReport @($jobResult) 2 'PreflightFailed'
    exit 2
}

$exitCode = Invoke-MaskingRun -Config $validatedConfig
exit [int]$exitCode
