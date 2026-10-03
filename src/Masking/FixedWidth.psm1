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

    $transformed = New-Object byte[] $Record.Length
    [System.Buffer]::BlockCopy($Record, 0, $transformed, 0, $Record.Length)

    foreach ($field in $Fields) {
        $mappingIndex = Get-FieldMappingIndex $MappingIndexes ([string]$field.Mapping) ([string]$field.Name)
        if ($field.PSObject.Properties['Segments']) {
            $parts = New-Object 'System.Collections.Generic.List[string]'
            foreach ($segment in @($field.Segments)) {
                $partBytes = Get-FixedFieldBytes $Record ([int]$segment.StartByte) ([int]$segment.ByteLength)
                try {
                    $part = $Encoding.GetString($partBytes)
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
                Set-FixedFieldBytes $transformed ([int]$segment.StartByte) $partBytes
                $cursor += $segmentWidth
            }
        }
        else {
            $fieldBytes = Get-FixedFieldBytes $Record ([int]$field.StartByte) ([int]$field.ByteLength)
            try {
                $value = $Encoding.GetString($fieldBytes)
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
            Set-FixedFieldBytes $transformed ([int]$field.StartByte) $replacementBytes
        }
    }

    return ,$transformed
}

function Invoke-FixedWidthJob {
    param(
        [object] $Job,
        [object] $MappingIndexes,
        [int[]] $SelectedRecordIndexes,
        [string] $TemporaryOutputPath
    )

    $recordWidth = [int]$Job.Layout.RecordWidthBytes
    if ($recordWidth -le 0) {
        throw 'RecordWidthBytes must be positive'
    }
    $fields = @($Job.Fields)
    if ($fields.Count -eq 0) {
        throw 'Fixed-width job must define fields'
    }
    $ranges = Get-FixedWidthRanges $fields $recordWidth
    foreach ($field in $fields) {
        [void](Get-FieldMappingIndex $MappingIndexes ([string]$field.Mapping) ([string]$field.Name))
    }

    $terminator = Get-FixedWidthTerminator ([string]$Job.Layout.RecordTerminator)
    if ($Job.Layout.PSObject.Properties['OutputTerminator']) {
        $outputTerminator = Get-FixedWidthTerminator ([string]$Job.Layout.OutputTerminator)
    }
    else {
        $outputTerminator = $terminator
    }
    $stride = [long]$recordWidth + $terminator.Length
    $fileLength = (Get-Item -LiteralPath $Job.InputPath -ErrorAction Stop).Length
    if (($fileLength % $stride) -ne 0) {
        throw 'File size is not divisible by record stride'
    }
    $recordCount = [int]($fileLength / $stride)
    if ($recordCount -eq 0) {
        throw 'Fixed-width input contains no records'
    }

    $encodingName = [string]$Job.Encoding
    $encoding = Get-FixedWidthEncoding $encodingName

    $inputStream = [System.IO.File]::Open(
        $Job.InputPath,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read
    )
    $outputStream = $null
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        if ($terminator.Length -gt 0) {
            for ($recordNumber = 0; $recordNumber -lt $recordCount; $recordNumber++) {
                $inputStream.Position = ([long]$recordNumber * $stride) + $recordWidth
                $actualTerminator = Read-FixedBytes $inputStream $terminator.Length "record $($recordNumber + 1) terminator"
                if (-not (Test-FixedBytesEqual $actualTerminator $terminator)) {
                    throw "Record $($recordNumber + 1) record terminator mismatch"
                }
            }
        }
        $inputStream.Position = 0

        if ($null -eq $SelectedRecordIndexes) {
            $selected = $null
            $recordsToWrite = $recordCount
        }
        else {
            if ($SelectedRecordIndexes.Count -eq 0) {
                throw 'SelectedRecordIndexes cannot be empty'
            }
            $seen = New-Object 'System.Collections.Generic.HashSet[int]'
            foreach ($recordIndex in $SelectedRecordIndexes) {
                if ($recordIndex -lt 0 -or $recordIndex -ge $recordCount) {
                    throw 'Selected record index is outside the input file'
                }
                if (-not $seen.Add($recordIndex)) {
                    throw 'Selected record indexes contain duplicates'
                }
            }
            $selected = [int[]]$SelectedRecordIndexes.Clone()
            [System.Array]::Sort($selected)
            $recordsToWrite = $selected.Length
        }

        $outputStream = [System.IO.File]::Open(
            $TemporaryOutputPath,
            [System.IO.FileMode]::CreateNew,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::None
        )

        if ($null -eq $selected) {
            for ($recordIndex = 0; $recordIndex -lt $recordCount; $recordIndex++) {
                $record = Read-FixedBytes $inputStream $recordWidth "record $($recordIndex + 1)"
                if ($terminator.Length -gt 0) {
                    $actualTerminator = Read-FixedBytes $inputStream $terminator.Length "record $($recordIndex + 1) terminator"
                    if (-not (Test-FixedBytesEqual $actualTerminator $terminator)) {
                        throw "Record $($recordIndex + 1) record terminator mismatch"
                    }
                }
                $transformed = Transform-FixedWidthRecord $record $fields $MappingIndexes $encoding ($recordIndex + 1)
                $outputStream.Write($transformed, 0, $transformed.Length)
                if ($outputTerminator.Length -gt 0) {
                    $outputStream.Write($outputTerminator, 0, $outputTerminator.Length)
                }
            }
        }
        else {
            foreach ($recordIndex in $selected) {
                $inputStream.Position = [long]$recordIndex * $stride
                $record = Read-FixedBytes $inputStream $recordWidth "record $($recordIndex + 1)"
                if ($terminator.Length -gt 0) {
                    $actualTerminator = Read-FixedBytes $inputStream $terminator.Length "record $($recordIndex + 1) terminator"
                    if (-not (Test-FixedBytesEqual $actualTerminator $terminator)) {
                        throw "Record $($recordIndex + 1) record terminator mismatch"
                    }
                }
                $transformed = Transform-FixedWidthRecord $record $fields $MappingIndexes $encoding ($recordIndex + 1)
                $outputStream.Write($transformed, 0, $transformed.Length)
                if ($outputTerminator.Length -gt 0) {
                    $outputStream.Write($outputTerminator, 0, $outputTerminator.Length)
                }
            }
        }

        $outputStream.Flush()
        $stopwatch.Stop()
        return [pscustomobject]@{
            RecordsRead       = $recordsToWrite
            RecordsWritten    = $recordsToWrite
            RecordsInSource   = $recordCount
            FieldsTransformed = $recordsToWrite * $fields.Count
            ElapsedMilliseconds = $stopwatch.ElapsedMilliseconds
        }
    }
    finally {
        if ($null -ne $outputStream) {
            $outputStream.Dispose()
        }
        $inputStream.Dispose()
    }
}

Export-ModuleMember -Function Invoke-FixedWidthJob
