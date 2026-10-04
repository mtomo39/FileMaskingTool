function Get-FixedWidthEncoding {
    param([string] $EncodingName)

    $codePage = 0
    if ([int]::TryParse($EncodingName, [ref]$codePage)) {
        $baseEncoding = [System.Text.Encoding]::GetEncoding($codePage)
    }
    else {
        $baseEncoding = [System.Text.Encoding]::GetEncoding($EncodingName)
    }

    $encoderFallback = New-Object System.Text.EncoderExceptionFallback
    $decoderFallback = New-Object System.Text.DecoderExceptionFallback
    return [System.Text.Encoding]::GetEncoding(
        $baseEncoding.CodePage,
        $encoderFallback,
        $decoderFallback
    )
}

function Read-FixedBytes {
    param(
        [System.IO.Stream] $Stream,
        [int] $Count,
        [string] $Context
    )

    $buffer = New-Object byte[] $Count
    $offset = 0
    while ($offset -lt $Count) {
        $read = $Stream.Read($buffer, $offset, $Count - $offset)
        if ($read -le 0) {
            throw "Unexpected end of file while reading $Context"
        }
        $offset += $read
    }
    return ,$buffer
}

function Test-FixedBytesEqual {
    param(
        [byte[]] $Left,
        [byte[]] $Right
    )

    if ($Left.Length -ne $Right.Length) {
        return $false
    }
    for ($i = 0; $i -lt $Left.Length; $i++) {
        if ($Left[$i] -ne $Right[$i]) {
            return $false
        }
    }
    return $true
}

function Get-FixedWidthTerminator {
    param([string] $Name)

    switch ($Name) {
        'None' { return ,([byte[]]@()) }
        'LF' { return ,([byte[]]@(10)) }
        'CRLF' { return ,([byte[]]@(13, 10)) }
        default { throw 'Unsupported fixed-width record terminator' }
    }
}

function Get-FixedWidthRanges {
    param(
        [object[]] $Fields,
        [int] $RecordWidthBytes
    )

    $ranges = New-Object 'System.Collections.Generic.List[object]'
    foreach ($field in $Fields) {
        if ([string]::IsNullOrWhiteSpace([string]$field.Name)) {
            throw 'Fixed-width field has no name'
        }

        if ($field.PSObject.Properties['Segments']) {
            $segments = @($field.Segments)
            if ($segments.Count -lt 2) {
                throw "Composite field $($field.Name) requires at least two segments"
            }
            $segmentWidth = 0
            $segmentNumber = 0
            foreach ($segment in $segments) {
                $segmentNumber++
                $start = [int]$segment.StartByte
                $width = [int]$segment.ByteLength
                if ($start -lt 1 -or $width -lt 1 -or ($start - 1 + $width) -gt $RecordWidthBytes) {
                    throw "Composite field $($field.Name) has an invalid range"
                }
                $segmentWidth += $width
                $ranges.Add([pscustomobject]@{
                    Name  = "$($field.Name).$segmentNumber"
                    Start = $start
                    End   = $start + $width - 1
                })
            }
            if ($segmentWidth -ne [int]$field.ByteLength) {
                throw "Composite field $($field.Name) byte length does not match its segments"
            }
        }
        else {
            $start = [int]$field.StartByte
            $width = [int]$field.ByteLength
            if ($start -lt 1 -or $width -lt 1 -or ($start - 1 + $width) -gt $RecordWidthBytes) {
                throw "Field $($field.Name) has an invalid byte range"
            }
            $ranges.Add([pscustomobject]@{
                Name  = [string]$field.Name
                Start = $start
                End   = $start + $width - 1
            })
        }
    }

    for ($i = 0; $i -lt $ranges.Count; $i++) {
        for ($j = $i + 1; $j -lt $ranges.Count; $j++) {
            if ($ranges[$i].Start -le $ranges[$j].End -and $ranges[$j].Start -le $ranges[$i].End) {
                throw "Fixed-width fields overlap: $($ranges[$i].Name) and $($ranges[$j].Name)"
            }
        }
    }

    return ,$ranges.ToArray()
}

function Get-FixedFieldBytes {
    param(
        [byte[]] $Record,
        [int] $StartByte,
        [int] $ByteLength
    )

    $bytes = New-Object byte[] $ByteLength
    [System.Buffer]::BlockCopy($Record, $StartByte - 1, $bytes, 0, $ByteLength)
    return ,$bytes
}

function Set-FixedFieldBytes {
    param(
        [byte[]] $Record,
        [int] $StartByte,
        [byte[]] $Value
    )

    [System.Buffer]::BlockCopy($Value, 0, $Record, $StartByte - 1, $Value.Length)
}

function Get-FieldMappingIndex {
    param(
        [object] $MappingIndexes,
        [string] $MappingName,
        [string] $FieldName
    )

    if ($null -eq $MappingIndexes -or -not $MappingIndexes.ContainsKey($MappingName)) {
        throw "Field $FieldName refers to an unavailable mapping index"
    }
    return $MappingIndexes[$MappingName]
}

function Transform-FixedWidthRecord {
    param(
        [byte[]] $Record,
        [object[]] $Fields,
        [object] $MappingIndexes,
        [System.Text.Encoding] $Encoding,
        [int] $RecordNumber
    )

    foreach ($preparedField in $Fields) {
        $field = $preparedField.Rule
        $mappingIndex = $preparedField.MappingIndex
        if ($field.PSObject.Properties['Segments']) {
            $parts = New-Object 'System.Collections.Generic.List[string]'
            foreach ($segment in @($field.Segments)) {
                try {
                    $part = $Encoding.GetString($Record, ([int]$segment.StartByte - 1), [int]$segment.ByteLength)
                }
                catch [System.Text.DecoderFallbackException] {
                    throw "Record $RecordNumber field $($field.Name) contains invalid encoded bytes"
                }
                if ($part -cnotmatch '^[0-9]+$') {
                    throw "Record $RecordNumber field $($field.Name) must contain digits only"
                }
                $parts.Add($part)
            }

            $logicalValue = [string]::Concat($parts.ToArray())
            if ($logicalValue.Length -ne 10) {
                throw "Record $RecordNumber field $($field.Name) must contain exactly 10 digits"
            }
            try {
                $replacement = Get-MappedValue $mappingIndex $logicalValue ([string]$field.TrimSpaces)
            }
            catch {
                throw "Record $RecordNumber field $($field.Name) mapping failed"
            }
            if ($replacement -cnotmatch '^[0-9]{10}$') {
                throw "Record $RecordNumber field $($field.Name) replacement must contain exactly 10 digits"
            }

            $cursor = 0
            foreach ($segment in @($field.Segments)) {
                $segmentWidth = [int]$segment.ByteLength
                $part = $replacement.Substring($cursor, $segmentWidth)
                try {
                    $partBytes = $Encoding.GetBytes($part)
                }
                catch [System.Text.EncoderFallbackException] {
                    throw "Record $RecordNumber field $($field.Name) replacement cannot be encoded"
                }
                if ($partBytes.Length -ne $segmentWidth) {
                    throw "Record $RecordNumber field $($field.Name) replacement segment has an invalid byte width"
                }
                Set-FixedFieldBytes $Record ([int]$segment.StartByte) $partBytes
                $cursor += $segmentWidth
            }
        }
        else {
            try {
                $value = $Encoding.GetString($Record, ([int]$field.StartByte - 1), [int]$field.ByteLength)
                $replacementBytes = Convert-FieldValue $value $field $mappingIndex $Encoding
            }
            catch [System.Text.DecoderFallbackException] {
                throw "Record $RecordNumber field $($field.Name) contains invalid encoded bytes"
            }
            catch {
                throw "Record $RecordNumber field $($field.Name) transformation failed: $($_.Exception.Message)"
            }
            if ($replacementBytes.Length -ne [int]$field.ByteLength) {
                throw "Record $RecordNumber field $($field.Name) output has an invalid byte width"
            }
            Set-FixedFieldBytes $Record ([int]$field.StartByte) $replacementBytes
        }
    }

    return ,$Record
}

function Invoke-FixedWidthJob {
    param(
        [object] $Job,
        [object] $MappingIndexes,
        [int[]] $SelectedRecordIndexes,
        [string] $TemporaryOutputPath
    )

    Import-MaskingCore
    $coreType = [FileMaskingTool.Core.MaskingCore]
    $fields = @($Job.Fields)
    $fieldType = [FileMaskingTool.Core.FieldRule]
    $fieldRules = [Array]::CreateInstance($fieldType, $fields.Count)
    for ($i = 0; $i -lt $fields.Count; $i++) {
        $sourceField = $fields[$i]
        $fieldRule = [Activator]::CreateInstance($fieldType)
        $fieldRule.Name = [string]$sourceField.Name
        $fieldRule.StartByte = [int]$sourceField.StartByte
        $fieldRule.ByteLength = [int]$sourceField.ByteLength
        $fieldRule.MappingName = [string]$sourceField.Mapping
        $fieldRule.TrimSpaces = [string]$sourceField.TrimSpaces
        $fieldRule.PadSpace = [string]$sourceField.PadSpace

        $sourceSegments = @()
        if ($sourceField.PSObject.Properties['Segments']) {
            $sourceSegments = @($sourceField.Segments)
        }
        if ($sourceSegments.Count -gt 0) {
            $segmentType = [FileMaskingTool.Core.SegmentRule]
            $segments = [Array]::CreateInstance($segmentType, $sourceSegments.Count)
            for ($segmentIndex = 0; $segmentIndex -lt $sourceSegments.Count; $segmentIndex++) {
                $segment = [Activator]::CreateInstance($segmentType)
                $segment.StartByte = [int]$sourceSegments[$segmentIndex].StartByte
                $segment.ByteLength = [int]$sourceSegments[$segmentIndex].ByteLength
                $segments.SetValue($segment, $segmentIndex)
            }
            $fieldRule.Segments = $segments
        }
        $fieldRules.SetValue($fieldRule, $i)
    }

    $terminator = Get-FixedWidthTerminator ([string]$Job.Layout.RecordTerminator)
    if ($Job.Layout.PSObject.Properties['OutputTerminator']) {
        $outputTerminator = Get-FixedWidthTerminator ([string]$Job.Layout.OutputTerminator)
    }
    else {
        $outputTerminator = $terminator
    }
    $options = [Activator]::CreateInstance([FileMaskingTool.Core.FixedWidthOptions])
    $options.InputPath = [string]$Job.InputPath
    $options.OutputPath = $TemporaryOutputPath
    $options.RecordWidthBytes = [int]$Job.Layout.RecordWidthBytes
    $options.InputTerminator = [byte[]]$terminator
    $options.OutputTerminator = [byte[]]$outputTerminator
    $options.Encoding = Get-FixedWidthEncoding ([string]$Job.Encoding)
    $options.Fields = $fieldRules
    $options.Mode = if ([string]::IsNullOrWhiteSpace([string]$Job.Mode)) { 'All' } else { [string]$Job.Mode }
    $options.SampleCount = if ($null -ne $Job.PSObject.Properties['SampleCount']) { [int]$Job.SampleCount } else { 0 }
    $options.Seed = if ($null -ne $Job.PSObject.Properties['Seed']) { [int]$Job.Seed } else { 0 }
    if ($null -ne $SelectedRecordIndexes) {
        $options.SelectedRecordIndexes = [int[]]$SelectedRecordIndexes
    }

    $result = $coreType::ProcessFixedWidth($options, $MappingIndexes)
    return [pscustomobject]@{
        RecordsRead         = $result.RecordsRead
        RecordsWritten      = $result.RecordsWritten
        RecordsInSource     = $result.RecordsInSource
        FieldsTransformed   = $result.FieldsTransformed
        ElapsedMilliseconds = $result.ElapsedMilliseconds
    }
}

Export-ModuleMember -Function Invoke-FixedWidthJob
