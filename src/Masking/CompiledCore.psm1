function Get-LoadedMaskingCoreType {
    foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
        $type = $assembly.GetType('FileMaskingTool.Core.MaskingCore', $false)
        if ($null -ne $type) {
            return $type
        }
    }
    return $null
}

function Import-MaskingCore {
    if ($null -ne (Get-LoadedMaskingCoreType)) {
        return
    }

    $sourcePath = Join-Path $PSScriptRoot 'MaskingCore.cs'
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw 'Compiled masking core source file is missing'
    }
    try {
        [void](Add-Type -Path $sourcePath -ErrorAction Stop)
    }
    catch {
        throw "Compiled masking core could not be loaded: $($_.Exception.Message)"
    }
    if ($null -eq (Get-LoadedMaskingCoreType)) {
        throw 'Compiled masking core type was not registered'
    }
}

Export-ModuleMember -Function Import-MaskingCore
