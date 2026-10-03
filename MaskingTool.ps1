[CmdletBinding()]
param(
    [string] $Config,
    [switch] $ValidateOnly,
    [switch] $Help
)

if ($Help) {
    Write-Output 'Usage: .\MaskingTool.ps1 -Config <path> [-ValidateOnly]'
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

Write-Error 'Job execution is not implemented yet; use -ValidateOnly.'
exit 2
