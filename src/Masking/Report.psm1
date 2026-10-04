function ConvertTo-ValueFreeErrorMessage {
    param(
        [System.Exception] $Exception,
        [string[]] $SensitivePaths
    )

    if ($null -eq $Exception) { return 'Unknown processing failure' }
    $message = [string]$Exception.Message
    foreach ($path in $SensitivePaths) {
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            $message = [System.Text.RegularExpressions.Regex]::Replace(
                $message,
                [System.Text.RegularExpressions.Regex]::Escape($path),
                '[path]',
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
        }
    }
    if ([string]::IsNullOrWhiteSpace($message)) {
        return $Exception.GetType().Name
    }
    return $message
}

function Write-MaskingRunReport {
    param(
        [object[]] $JobResults,
        [int] $ExitCode,
        [string] $Status
    )

    $report = [pscustomobject]@{
        Status = $Status
        ExitCode = $ExitCode
        Jobs = @($JobResults)
        Limitations = @(
            'Only configured fields and ranges were transformed; unconfigured content remains unchanged.',
            'Excel auxiliary-content warnings do not prove that a workbook contains no personal information.'
        )
    }
    Write-Host (ConvertTo-Json -InputObject $report -Depth 8 -Compress)
}

Export-ModuleMember -Function ConvertTo-ValueFreeErrorMessage, Write-MaskingRunReport
