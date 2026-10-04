$script:StrictTransformEncodingCache = [System.Collections.Generic.Dictionary[int,System.Text.Encoding]]::new()

function Get-StrictTransformEncoding {
    param([System.Text.Encoding] $Encoding)

    if ($null -eq $Encoding) {
        throw 'Field encoding is required'
    }
    if ($script:StrictTransformEncodingCache.ContainsKey($Encoding.CodePage)) {
        return $script:StrictTransformEncodingCache[$Encoding.CodePage]
    }

    $encoderFallback = New-Object System.Text.EncoderExceptionFallback
    $decoderFallback = New-Object System.Text.DecoderExceptionFallback
    $strictEncoding = [System.Text.Encoding]::GetEncoding(
        $Encoding.CodePage,
        $encoderFallback,
        $decoderFallback
    )
    $script:StrictTransformEncodingCache.Add($Encoding.CodePage, $strictEncoding)
    return $strictEncoding
}

function Convert-FieldValue {
    param(
        [string] $Value,
        [object] $FieldRule,
        [System.Collections.Generic.Dictionary[string,string]] $MappingIndex,
        [System.Text.Encoding] $Encoding
    )

    if ($null -eq $FieldRule -or [int]$FieldRule.ByteLength -le 0) {
        throw 'Field rule has an invalid byte length'
    }
    if ([string]::IsNullOrWhiteSpace([string]$FieldRule.Name)) {
        throw 'Field rule has no name'
    }

    $fieldName = [string]$FieldRule.Name
    $replacement = Get-MappedValue $MappingIndex $Value ([string]$FieldRule.TrimSpaces)
    $strictEncoding = Get-StrictTransformEncoding $Encoding

    try {
        $valueBytes = $strictEncoding.GetBytes($replacement)
    }
    catch [System.Text.EncoderFallbackException] {
        throw "Field $fieldName replacement cannot be represented in the configured encoding"
    }

    $fieldWidth = [int]$FieldRule.ByteLength
    $remaining = $fieldWidth - $valueBytes.Length
    if ($remaining -lt 0) {
        throw "Field $fieldName mapped value exceeds field width"
    }

    switch ([string]$FieldRule.PadSpace) {
        'Half' { $padCharacter = [char]0x20 }
        'Full' { $padCharacter = [char]0x3000 }
        default { throw "Field $fieldName PadSpace must be Half or Full" }
    }

    try {
        $padBytes = $strictEncoding.GetBytes([string]$padCharacter)
    }
    catch [System.Text.EncoderFallbackException] {
        throw "Field $fieldName padding character cannot be represented in the configured encoding"
    }

    if ($padBytes.Length -eq 0 -or ($remaining % $padBytes.Length) -ne 0) {
        throw "Field $fieldName padding cannot be filled exactly"
    }

    $result = New-Object byte[] $fieldWidth
    if ($valueBytes.Length -gt 0) {
        [System.Buffer]::BlockCopy($valueBytes, 0, $result, 0, $valueBytes.Length)
    }
    for ($offset = $valueBytes.Length; $offset -lt $fieldWidth; $offset += $padBytes.Length) {
        [System.Buffer]::BlockCopy($padBytes, 0, $result, $offset, $padBytes.Length)
    }

    return ,$result
}

Export-ModuleMember -Function Convert-FieldValue
