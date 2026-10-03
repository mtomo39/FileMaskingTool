[CmdletBinding()]
param(
    [switch] $Help
)

if ($Help) {
    Write-Output 'Usage: .\MaskingTool.ps1 -Config <path> [-ValidateOnly]'
    exit 0
}

Write-Error 'Specify -Help until job execution is implemented.'
exit 2
