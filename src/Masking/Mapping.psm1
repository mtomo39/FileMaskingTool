function Normalize-MappingKey {
    param(
        [string] $Value,
        [string] $TrimMode
    )

    switch ($TrimMode) {
        'None' { return $Value }
        'Half' { return $Value.Trim([char[]]@([char]0x20)) }
        'Full' { return $Value.Trim([char[]]@([char]0x3000)) }
        'HalfAndFull' { return $Value.Trim([char[]]@([char]0x20, [char]0x3000)) }
        default { throw 'Unsupported mapping key trim mode' }
    }
}

function Get-StrictMappingEncoding {
    param([System.Text.Encoding] $Encoding)

    $encoderFallback = New-Object System.Text.EncoderExceptionFallback
    $decoderFallback = New-Object System.Text.DecoderExceptionFallback
    return [System.Text.Encoding]::GetEncoding(
        $Encoding.CodePage,
        $encoderFallback,
        $decoderFallback
    )
}

function Read-MappingTable {
    param(
        [string] $Path,
        [System.Text.Encoding] $Encoding,
        [string] $KeyTrimMode
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw 'Mapping file does not exist'
    }
    if ($null -eq $Encoding) {
        throw 'Mapping encoding is required'
    }

    $index = [System.Collections.Generic.Dictionary[string,string]]::new([System.StringComparer]::Ordinal)
    $strictEncoding = Get-StrictMappingEncoding $Encoding
    $reader = [System.IO.StreamReader]::new($Path, $strictEncoding, $true)
    $lineNumber = 0
    try {
        while ($null -ne ($line = $reader.ReadLine())) {
            $lineNumber++
            $firstTab = $line.IndexOf([char]9)
            $lastTab = $line.LastIndexOf([char]9)
            if ($firstTab -lt 0 -or $firstTab -ne $lastTab) {
                throw "Mapping row $lineNumber must contain exactly two tab-separated columns"
            }

            $source = $line.Substring(0, $firstTab)
            $replacement = $line.Substring($firstTab + 1)
            $key = Normalize-MappingKey $source $KeyTrimMode
            if ([string]::IsNullOrEmpty($key)) {
                throw "Mapping row $lineNumber has an empty normalized key"
            }

            if ($index.ContainsKey($key)) {
                if (-not [string]::Equals($index[$key], $replacement, [System.StringComparison]::Ordinal)) {
                    throw "Conflicting duplicate mapping key at row $lineNumber"
                }
                continue
            }

            $index.Add($key, $replacement)
        }
    }
    catch [System.Text.DecoderFallbackException] {
        throw "Mapping file contains bytes invalid for the configured encoding near row $($lineNumber + 1)"
    }
    finally {
        $reader.Dispose()
    }

    return $index
}

function Get-MappedValue {
    param(
        [System.Collections.Generic.Dictionary[string,string]] $Index,
        [string] $InputValue,
        [string] $KeyTrimMode
    )

    if ($null -eq $Index) {
        throw 'Mapping index is required'
    }

    $key = Normalize-MappingKey $InputValue $KeyTrimMode
    $replacement = $null
    if (-not $Index.TryGetValue($key, [ref]$replacement)) {
        throw 'Mapping key not found'
    }

    return [string]$replacement
}

Export-ModuleMember -Function Read-MappingTable, Get-MappedValue
