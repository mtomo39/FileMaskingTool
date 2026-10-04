function Read-CsvRecord {
    param(
        [System.IO.TextReader] $Reader,
        [char] $Delimiter,
        [char] $Quote
    )

    if ($null -eq $Reader -or $Delimiter -eq $Quote -or $Delimiter -eq "`r" -or $Delimiter -eq "`n" -or $Quote -eq "`r" -or $Quote -eq "`n") {
        throw 'CSV reader or dialect is invalid'
    }

    $fields = New-Object 'System.Collections.Generic.List[object]'
    $value = New-Object System.Text.StringBuilder
    $wasQuoted = $false
    $inQuotes = $false
    $afterQuote = $false
    $recordStarted = $false

    while ($true) {
        $code = $Reader.Read()
        if ($code -lt 0) {
            if ($inQuotes) { throw 'CSV contains an unterminated quoted field' }
            if (-not $recordStarted -and $fields.Count -eq 0 -and $value.Length -eq 0) { return $null }
            $fields.Add([pscustomobject]@{ Value = $value.ToString(); WasQuoted = $wasQuoted })
            return [pscustomobject]@{ Fields = $fields.ToArray(); Terminator = '' }
        }

        $character = [char]$code
        $recordStarted = $true
        if ($inQuotes) {
            if ($character -eq $Quote) {
                if ($Reader.Peek() -eq [int]$Quote) {
                    [void]$Reader.Read()
                    [void]$value.Append($Quote)
                }
                else {
                    $inQuotes = $false
                    $afterQuote = $true
                }
            }
            else { [void]$value.Append($character) }
            continue
        }

        if ($afterQuote -and $character -ne $Delimiter -and $character -ne "`r" -and $character -ne "`n") {
            throw 'CSV has data after a closing quote'
        }
        if ($character -eq $Delimiter) {
            $fields.Add([pscustomobject]@{ Value = $value.ToString(); WasQuoted = $wasQuoted })
            $value = New-Object System.Text.StringBuilder
            $wasQuoted = $false
            $afterQuote = $false
            continue
        }
        if ($character -eq "`r" -or $character -eq "`n") {
            $terminator = [string]$character
            if ($character -eq "`r" -and $Reader.Peek() -eq 10) {
                [void]$Reader.Read()
                $terminator = "`r`n"
            }
            $fields.Add([pscustomobject]@{ Value = $value.ToString(); WasQuoted = $wasQuoted })
            return [pscustomobject]@{ Fields = $fields.ToArray(); Terminator = $terminator }
        }
        if ($character -eq $Quote) {
            if ($afterQuote -or $value.Length -ne 0) { throw 'CSV quote occurs inside an unquoted field' }
            $wasQuoted = $true
            $inQuotes = $true
            continue
        }
        [void]$value.Append($character)
    }
}

function Read-CsvRecords {
    param(
        [System.IO.TextReader] $Reader,
        [char] $Delimiter,
        [char] $Quote
    )

    while ($null -ne ($record = Read-CsvRecord $Reader $Delimiter $Quote)) {
        Write-Output -NoEnumerate $record
    }
}

function Write-CsvRecord {
    param(
        [object] $Record,
        [char] $Delimiter,
        [char] $Quote
    )

    if ($null -eq $Record -or $null -eq $Record.Fields -or $Delimiter -eq $Quote) {
        throw 'CSV record or dialect is invalid'
    }

    $parts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($field in $Record.Fields) {
        $value = [string]$field.Value
        $requiresQuotes = $value.IndexOf($Delimiter) -ge 0 -or
            $value.IndexOf($Quote) -ge 0 -or
            $value.IndexOf("`r") -ge 0 -or
            $value.IndexOf("`n") -ge 0

        if ([bool]$field.WasQuoted) {
            $escaped = $value.Replace([string]$Quote, ([string]$Quote + [string]$Quote))
            $parts.Add(([string]$Quote + $escaped + [string]$Quote))
        }
        elseif ($requiresQuotes) {
            throw 'Unquoted CSV field requires quoting'
        }
        else {
            $parts.Add($value)
        }
    }

    return [string]::Join([string]$Delimiter, $parts.ToArray())
}

function Get-CsvEncoding {
    param([string] $Name)

    $codePage = 0
    if ([int]::TryParse($Name, [ref]$codePage)) {
        $baseEncoding = [System.Text.Encoding]::GetEncoding($codePage)
    }
    else {
        $baseEncoding = [System.Text.Encoding]::GetEncoding($Name)
    }

    $encoderFallback = New-Object System.Text.EncoderExceptionFallback
    $decoderFallback = New-Object System.Text.DecoderExceptionFallback
    return [System.Text.Encoding]::GetEncoding($baseEncoding.CodePage, $encoderFallback, $decoderFallback)
}

function Get-CsvOutputEncoding {
    param([System.Text.Encoding] $Encoding)

    switch ($Encoding.CodePage) {
        65001 { return New-Object System.Text.UTF8Encoding($false, $true) }
        1200 { return New-Object System.Text.UnicodeEncoding($false, $false, $true) }
        1201 { return New-Object System.Text.UnicodeEncoding($true, $false, $true) }
        12000 { return New-Object System.Text.UTF32Encoding($false, $false, $true) }
        12001 { return New-Object System.Text.UTF32Encoding($true, $false, $true) }
        default { return $Encoding }
    }
}

function Test-CsvPreamble {
    param(
        [string] $Path,
        [System.Text.Encoding] $Encoding
    )

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        if ($stream.Length -eq 0) { return $false }
        $readLength = [int][Math]::Min(4, $stream.Length)
        $actual = New-Object byte[] $readLength
        $read = $stream.Read($actual, 0, $actual.Length)
        if ($read -ne $actual.Length) { return $false }

        $detectedCodePage = 0
        if ($actual.Length -ge 4 -and $actual[0] -eq 0xFF -and $actual[1] -eq 0xFE -and $actual[2] -eq 0 -and $actual[3] -eq 0) {
            $detectedCodePage = 12000
        }
        elseif ($actual.Length -ge 4 -and $actual[0] -eq 0 -and $actual[1] -eq 0 -and $actual[2] -eq 0xFE -and $actual[3] -eq 0xFF) {
            $detectedCodePage = 12001
        }
        elseif ($actual.Length -ge 3 -and $actual[0] -eq 0xEF -and $actual[1] -eq 0xBB -and $actual[2] -eq 0xBF) {
            $detectedCodePage = 65001
        }
        elseif ($actual.Length -ge 2 -and $actual[0] -eq 0xFE -and $actual[1] -eq 0xFF) {
            $detectedCodePage = 1201
        }
        elseif ($actual.Length -ge 2 -and $actual[0] -eq 0xFF -and $actual[1] -eq 0xFE) {
            $detectedCodePage = 1200
        }

        if ($detectedCodePage -ne 0 -and $detectedCodePage -ne $Encoding.CodePage) {
            throw 'CSV byte-order mark does not match the configured encoding'
        }
        return $detectedCodePage -ne 0
    }
    finally {
        $stream.Dispose()
    }
}

function Open-CsvTextReader {
    param(
        [string] $Path,
        [System.Text.Encoding] $Encoding
    )

    $hasPreamble = Test-CsvPreamble $Path $Encoding
    $stream = [System.IO.File]::OpenRead($Path)
    try {
        if ($hasPreamble) {
            $stream.Position = $Encoding.GetPreamble().Length
        }
        return [System.IO.StreamReader]::new($stream, $Encoding, $false, 65536)
    }
    catch {
        $stream.Dispose()
        throw
    }
}

function Get-CsvTerminatorBytes {
    param([string] $Name)

    switch ($Name) {
        'None' { return ,([byte[]]@()) }
        'LF' { return ,([byte[]]@(10)) }
        'CRLF' { return ,([byte[]]@(13, 10)) }
        default { throw 'Unsupported CSV record terminator' }
    }
}

function Read-CsvBlock {
    param(
        [System.IO.Stream] $Stream,
        [int] $Width,
        [byte[]] $Terminator,
        [int] $RecordNumber,
        [System.Text.Encoding] $Encoding,
        [char] $Delimiter,
        [char] $Quote
    )

    $block = New-Object byte[] $Width
    $offset = 0
    while ($offset -lt $Width) {
        $read = $Stream.Read($block, $offset, $Width - $offset)
        if ($read -le 0) {
            if ($offset -eq 0) { return $null }
            throw "CSV fixed-width record $RecordNumber is truncated"
        }
        $offset += $read
    }

    if ($Terminator.Length -gt 0) {
        $actualTerminator = New-Object byte[] $Terminator.Length
        $termOffset = 0
        while ($termOffset -lt $Terminator.Length) {
            $read = $Stream.Read($actualTerminator, $termOffset, $Terminator.Length - $termOffset)
            if ($read -le 0) { throw "CSV fixed-width record $RecordNumber has an invalid terminator" }
            $termOffset += $read
        }
        for ($i = 0; $i -lt $Terminator.Length; $i++) {
            if ($actualTerminator[$i] -ne $Terminator[$i]) {
                throw "CSV fixed-width record $RecordNumber has an invalid terminator"
            }
        }
    }

    $payloadLength = $Width
    while ($payloadLength -gt 0 -and $block[$payloadLength - 1] -eq 0) { $payloadLength-- }
    for ($i = 0; $i -lt $payloadLength; $i++) {
        if ($block[$i] -eq 0) { throw 'NUL byte occurs inside CSV payload' }
    }

    try {
        $payload = $Encoding.GetString($block, 0, $payloadLength)
    }
    catch [System.Text.DecoderFallbackException] {
        throw "CSV fixed-width record $RecordNumber contains invalid encoded bytes"
    }

    $reader = [System.IO.StringReader]::new($payload)
    try {
        $record = Read-CsvRecord $reader $Delimiter $Quote
        $extraRecord = Read-CsvRecord $reader $Delimiter $Quote
    }
    finally {
        $reader.Dispose()
    }
    if ($null -eq $record -or $null -ne $extraRecord -or $record.Terminator -ne '') {
        throw "CSV fixed-width record $RecordNumber contains more than one CSV record"
    }
    return $record
}

function Get-CsvTargetColumns {
    param(
        [object[]] $Fields,
        [object] $HeaderRecord
    )

    $targets = New-Object 'System.Collections.Generic.List[object]'
    $usedColumns = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($field in $Fields) {
        if ($field.PSObject.Properties['ColumnName'] -and -not [string]::IsNullOrWhiteSpace([string]$field.ColumnName)) {
            if ($null -eq $HeaderRecord) { throw "CSV field $($field.Name) requires a header row" }
            $columnMatches = New-Object 'System.Collections.Generic.List[int]'
            for ($i = 0; $i -lt $HeaderRecord.Fields.Count; $i++) {
                if ([string]::Equals([string]$HeaderRecord.Fields[$i].Value, [string]$field.ColumnName, [System.StringComparison]::Ordinal)) {
                    $columnMatches.Add($i)
                }
            }
            if ($columnMatches.Count -ne 1) {
                throw "CSV header must contain exactly one column named $($field.ColumnName)"
            }
            $column = $columnMatches[0]
        }
        else {
            $column = [int]$field.ColumnIndex - 1
        }
        if ($null -ne $HeaderRecord -and $column -ge $HeaderRecord.Fields.Count) {
            throw "CSV header has no column $($column + 1) required by field $($field.Name)"
        }
        if (-not $usedColumns.Add($column)) {
            throw "CSV fields target the same column index $($column + 1)"
        }
        $targets.Add([pscustomobject]@{ Rule = $field; ColumnIndex = $column })
    }
    return ,$targets.ToArray()
}

function Convert-CsvDataRecord {
    param(
        [object] $Record,
        [object[]] $Targets,
        [object] $MappingIndexes
    )

    $transformed = New-Object 'System.Collections.Generic.List[object]'
    foreach ($field in $Record.Fields) {
        $transformed.Add([pscustomobject]@{ Value = [string]$field.Value; WasQuoted = [bool]$field.WasQuoted })
    }
    foreach ($target in $Targets) {
        $index = [int]$target.ColumnIndex
        if ($index -lt 0 -or $index -ge $transformed.Count) {
            throw "CSV record has no column $($index + 1) required by field $($target.Rule.Name)"
        }
        $rule = $target.Rule
        $mappingName = [string]$rule.Mapping
        if ($null -eq $MappingIndexes -or -not $MappingIndexes.ContainsKey($mappingName)) {
            throw "CSV field $($rule.Name) refers to an unavailable mapping index"
        }
        $oldField = $transformed[$index]
        $mapped = Get-MappedValue $MappingIndexes[$mappingName] ([string]$oldField.Value) ([string]$rule.TrimSpaces)
        $transformed[$index] = [pscustomobject]@{ Value = $mapped; WasQuoted = [bool]$oldField.WasQuoted }
    }
    return [pscustomobject]@{ Fields = $transformed.ToArray() }
}

function Get-CsvSampleIndexes {
    param(
        [object] $Job,
        [System.Text.Encoding] $Encoding
    )

    $selected = New-Object 'System.Collections.Generic.HashSet[int]'
    $headerRead = -not [bool]$Job.Layout.Header
    $dataIndex = 0
    $expectedColumnCount = -1
    $sampleTargets = $null
    $random = [System.Random]::new([int]$Job.Seed)
    $sampleSize = [int]$Job.SampleCount
    $delimiter = [char][string]$Job.Layout.Delimiter
    $quote = [char][string]$Job.Layout.Quote

    if ([string]$Job.Layout.ContainerMode -eq 'FixedWidthNullPadded') {
        $width = [int]$Job.Layout.RecordWidthBytes
        $terminator = Get-CsvTerminatorBytes ([string]$Job.Layout.RecordTerminator)
        $stride = $width + $terminator.Length
        $stream = [System.IO.File]::OpenRead([string]$Job.InputPath)
        try {
            if (($stream.Length % $stride) -ne 0) { throw 'CSV fixed-width file length is not a whole number of records' }
            $recordNumber = 0
            while ($stream.Position -lt $stream.Length) {
                $recordNumber++
                $record = Read-CsvBlock $stream $width $terminator $recordNumber $Encoding $delimiter $quote
                if (-not $headerRead) {
                    $headerRead = $true
                    $expectedColumnCount = $record.Fields.Count
                    $sampleTargets = Get-CsvTargetColumns @($Job.Fields) $record
                    continue
                }
                if ($null -eq $sampleTargets) { $sampleTargets = Get-CsvTargetColumns @($Job.Fields) $null }
                if ($expectedColumnCount -lt 0) { $expectedColumnCount = $record.Fields.Count }
                if ($record.Fields.Count -ne $expectedColumnCount) { throw "CSV record $($dataIndex + 1) has an inconsistent column count" }
                foreach ($target in $sampleTargets) {
                    if ([int]$target.ColumnIndex -ge $record.Fields.Count) {
                        throw "CSV record $($dataIndex + 1) has no column $([int]$target.ColumnIndex + 1) required by field $($target.Rule.Name)"
                    }
                }
                if ($selected.Count -lt $sampleSize) { [void]$selected.Add($dataIndex) }
                else {
                    $replacementIndex = $random.Next(0, $dataIndex + 1)
                    if ($replacementIndex -lt $sampleSize) {
                        [void]$selected.Remove($replacementIndex)
                        [void]$selected.Add($dataIndex)
                    }
                }
                $dataIndex++
            }
        }
        finally { $stream.Dispose() }
    }
    else {
        $reader = Open-CsvTextReader ([string]$Job.InputPath) $Encoding
        try {
            while ($null -ne ($record = Read-CsvRecord $reader $delimiter $quote)) {
                if (-not $headerRead) {
                    $headerRead = $true
                    $expectedColumnCount = $record.Fields.Count
                    $sampleTargets = Get-CsvTargetColumns @($Job.Fields) $record
                    continue
                }
                if ($null -eq $sampleTargets) { $sampleTargets = Get-CsvTargetColumns @($Job.Fields) $null }
                if ($expectedColumnCount -lt 0) { $expectedColumnCount = $record.Fields.Count }
                if ($record.Fields.Count -ne $expectedColumnCount) { throw "CSV record $($dataIndex + 1) has an inconsistent column count" }
                foreach ($target in $sampleTargets) {
                    if ([int]$target.ColumnIndex -ge $record.Fields.Count) {
                        throw "CSV record $($dataIndex + 1) has no column $([int]$target.ColumnIndex + 1) required by field $($target.Rule.Name)"
                    }
                }
                if ($selected.Count -lt $sampleSize) { [void]$selected.Add($dataIndex) }
                else {
                    $replacementIndex = $random.Next(0, $dataIndex + 1)
                    if ($replacementIndex -lt $sampleSize) {
                        [void]$selected.Remove($replacementIndex)
                        [void]$selected.Add($dataIndex)
                    }
                }
                $dataIndex++
            }
        }
        finally { $reader.Dispose() }
    }
    if ([bool]$Job.Layout.Header -and -not $headerRead) { throw 'CSV header row is missing' }
    return ,@($selected | Sort-Object)
}

function Invoke-CsvJob {
    param(
        [object] $Job,
        [object] $MappingIndexes,
        [string] $OutputPath
    )

    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'CSV output path is required' }
    if (-not (Test-Path -LiteralPath ([string]$Job.InputPath) -PathType Leaf)) { throw 'CSV input file does not exist' }
    Import-MaskingCore
    $fields = @($Job.Fields)
    if ($fields.Count -eq 0) { throw 'CSV job must define fields' }
    $fieldType = [FileMaskingTool.Core.CsvFieldRule]
    $fieldRules = [Array]::CreateInstance($fieldType, $fields.Count)
    for ($i = 0; $i -lt $fields.Count; $i++) {
        $sourceField = $fields[$i]
        $fieldRule = [Activator]::CreateInstance($fieldType)
        $fieldRule.Name = [string]$sourceField.Name
        $fieldRule.ColumnIndex = [int]$sourceField.ColumnIndex
        if ($sourceField.PSObject.Properties['ColumnName']) {
            $fieldRule.ColumnName = [string]$sourceField.ColumnName
        }
        $fieldRule.MappingName = [string]$sourceField.Mapping
        $fieldRule.TrimSpaces = [string]$sourceField.TrimSpaces
        $fieldRules.SetValue($fieldRule, $i)
    }

    $fixedWidth = [string]$Job.Layout.ContainerMode -eq 'FixedWidthNullPadded'
    $terminator = [byte[]]@()
    $recordWidth = 0
    $nullPadByte = [byte]0
    if ($fixedWidth) {
        $recordWidth = [int]$Job.Layout.RecordWidthBytes
        $terminator = Get-CsvTerminatorBytes ([string]$Job.Layout.RecordTerminator)
        if ($Job.Layout.PSObject.Properties['NullPadByte']) {
            $nullPadByte = [byte]$Job.Layout.NullPadByte
        }
    }
    $encoding = Get-CsvEncoding ([string]$Job.Encoding)
    $temporaryPath = $OutputPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        $options = [Activator]::CreateInstance([FileMaskingTool.Core.CsvOptions])
        $options.InputPath = [string]$Job.InputPath
        $options.OutputPath = $temporaryPath
        $options.Delimiter = [char][string]$Job.Layout.Delimiter
        $options.Quote = [char][string]$Job.Layout.Quote
        $options.Header = [bool]$Job.Layout.Header
        $options.ContainerMode = if ([string]::IsNullOrWhiteSpace([string]$Job.Layout.ContainerMode)) { 'Records' } else { [string]$Job.Layout.ContainerMode }
        $options.RecordWidthBytes = $recordWidth
        $options.NullPadByte = $nullPadByte
        $options.RecordTerminator = $terminator
        $options.Encoding = $encoding
        $options.Fields = $fieldRules
        $options.Mode = if ([string]::IsNullOrWhiteSpace([string]$Job.Mode)) { 'All' } else { [string]$Job.Mode }
        $options.SampleCount = if ($null -ne $Job.PSObject.Properties['SampleCount']) { [int]$Job.SampleCount } else { 0 }
        $options.Seed = if ($null -ne $Job.PSObject.Properties['Seed']) { [int]$Job.Seed } else { 0 }
        $result = [FileMaskingTool.Core.MaskingCore]::ProcessCsv($options, $MappingIndexes)
        if (Test-Path -LiteralPath $OutputPath -PathType Leaf) {
            [System.IO.File]::Replace($temporaryPath, $OutputPath, $null)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $OutputPath)
        }
        return [pscustomobject]@{
            RecordsRead = $result.RecordsRead
            RecordsWritten = if ($fixedWidth) { $result.RecordsWritten } else { $result.DataRecordsWritten }
            DataRecordsWritten = $result.DataRecordsWritten
            ElapsedMilliseconds = $result.ElapsedMilliseconds
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            [System.IO.File]::Delete($temporaryPath)
        }
    }
}

function Invoke-CsvJobLegacy {
    param(
        [object] $Job,
        [object] $MappingIndexes,
        [string] $OutputPath
    )

    $encoding = Get-CsvEncoding ([string]$Job.Encoding)
    $outputEncoding = Get-CsvOutputEncoding $encoding
    $delimiter = [char][string]$Job.Layout.Delimiter
    $quote = [char][string]$Job.Layout.Quote
    $fixedWidth = [string]$Job.Layout.ContainerMode -eq 'FixedWidthNullPadded'
    if ($fixedWidth -and @(1200, 1201, 12000, 12001) -contains $encoding.CodePage) {
        throw 'Fixed-width NUL-padded CSV does not support UTF-16 or UTF-32 encodings'
    }
    $selectedIndexes = $null
    if ([string]$Job.Mode -eq 'Sample') {
        $selectedIndexes = New-Object 'System.Collections.Generic.HashSet[int]'
        foreach ($index in (Get-CsvSampleIndexes $Job $encoding)) { [void]$selectedIndexes.Add([int]$index) }
    }
    elseif ([string]$Job.Mode -ne 'All') {
        throw 'CSV Mode must be Sample or All'
    }

    $headerRecord = $null
    $targets = $null
    $expectedColumnCount = -1
    $dataIndex = 0
    $recordsWritten = 0
    $recordsRead = 0
    $tempPath = $OutputPath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    $inputHasPreamble = $false
    if (-not $fixedWidth) {
        $inputHasPreamble = Test-CsvPreamble ([string]$Job.InputPath) $encoding
    }

    try {
        if ($fixedWidth) {
            $width = [int]$Job.Layout.RecordWidthBytes
            $terminator = Get-CsvTerminatorBytes ([string]$Job.Layout.RecordTerminator)
            $stride = $width + $terminator.Length
            $inputStream = [System.IO.File]::OpenRead([string]$Job.InputPath)
            try {
                if (($inputStream.Length % $stride) -ne 0) {
                    throw 'CSV fixed-width file length is not a whole number of records'
                }
                $outputStream = [System.IO.File]::Create($tempPath)
                try {
                    $recordNumber = 0
                    while ($inputStream.Position -lt $inputStream.Length) {
                        $recordNumber++
                        $record = Read-CsvBlock $inputStream $width $terminator $recordNumber $encoding $delimiter $quote
                        $recordsRead++
                        if ([bool]$Job.Layout.Header -and $recordNumber -eq 1) {
                            $headerRecord = $record
                            $expectedColumnCount = $record.Fields.Count
                            $targets = Get-CsvTargetColumns @($Job.Fields) $headerRecord
                            $outputRecord = $record
                        }
                        else {
                            if ($null -eq $targets) {
                                $targets = Get-CsvTargetColumns @($Job.Fields) $null
                            }
                            if ($expectedColumnCount -lt 0) { $expectedColumnCount = $record.Fields.Count }
                            if ($record.Fields.Count -ne $expectedColumnCount) {
                                throw "CSV record $recordNumber has an inconsistent column count"
                            }
                            $shouldWrite = $null -eq $selectedIndexes -or $selectedIndexes.Contains($dataIndex)
                            if ($shouldWrite) {
                                $outputRecord = Convert-CsvDataRecord $record $targets $MappingIndexes
                            }
                            $dataIndex++
                            if (-not $shouldWrite) { continue }
                        }

                        $text = Write-CsvRecord $outputRecord $delimiter $quote
                        try {
                            $payload = $encoding.GetBytes($text)
                        }
                        catch [System.Text.EncoderFallbackException] {
                            throw "CSV record $recordNumber cannot be represented in the configured encoding"
                        }
                        if ($payload.Length -gt $width) {
                            throw "CSV record $recordNumber exceeds fixed-width CSV record payload"
                        }
                        $outputStream.Write($payload, 0, $payload.Length)
                        for ($padding = $payload.Length; $padding -lt $width; $padding++) {
                            $outputStream.WriteByte(0)
                        }
                        if ($terminator.Length -gt 0) { $outputStream.Write($terminator, 0, $terminator.Length) }
                        $recordsWritten++
                    }
                }
                finally {
                    $outputStream.Dispose()
                }
            }
            finally {
                $inputStream.Dispose()
            }
        }
        else {
            $reader = Open-CsvTextReader ([string]$Job.InputPath) $encoding
            try {
                $outputStream = [System.IO.File]::Create($tempPath)
                try {
                    if ($inputHasPreamble) {
                        $preamble = $outputEncoding.GetPreamble()
                        if ($preamble.Length -gt 0) { $outputStream.Write($preamble, 0, $preamble.Length) }
                    }
                    $writer = [System.IO.StreamWriter]::new($outputStream, $outputEncoding, 65536, $true)
                    try {
                        while ($null -ne ($record = Read-CsvRecord $reader $delimiter $quote)) {
                            $recordsRead++
                            if ([bool]$Job.Layout.Header -and $null -eq $headerRecord) {
                                $headerRecord = $record
                                $expectedColumnCount = $record.Fields.Count
                                $targets = Get-CsvTargetColumns @($Job.Fields) $headerRecord
                                $writer.Write((Write-CsvRecord $record $delimiter $quote))
                                $writer.Write([string]$record.Terminator)
                                continue
                            }
                            if ($null -eq $targets) { $targets = Get-CsvTargetColumns @($Job.Fields) $null }
                            if ($expectedColumnCount -lt 0) { $expectedColumnCount = $record.Fields.Count }
                            if ($record.Fields.Count -ne $expectedColumnCount) {
                                throw "CSV record $($dataIndex + 1) has an inconsistent column count"
                            }

                            $shouldWrite = $null -eq $selectedIndexes -or $selectedIndexes.Contains($dataIndex)
                            if ($shouldWrite) {
                                $outputRecord = Convert-CsvDataRecord $record $targets $MappingIndexes
                                $writer.Write((Write-CsvRecord $outputRecord $delimiter $quote))
                                $writer.Write([string]$record.Terminator)
                                $recordsWritten++
                            }
                            $dataIndex++
                        }
                    }
                    finally {
                        $writer.Dispose()
                    }
                }
                finally {
                    $outputStream.Dispose()
                }
            }
            finally {
                $reader.Dispose()
            }
        }

        if ([bool]$Job.Layout.Header -and $null -eq $headerRecord) {
            throw 'CSV header row is missing'
        }
        if (Test-Path -LiteralPath $OutputPath -PathType Leaf) {
            [System.IO.File]::Replace($tempPath, $OutputPath, $null)
        }
        else {
            [System.IO.File]::Move($tempPath, $OutputPath)
        }
        return [pscustomobject]@{
            RecordsRead = $recordsRead
            RecordsWritten = $recordsWritten
            DataRecordsWritten = if ([bool]$Job.Layout.Header) { [Math]::Max(0, $recordsWritten - 1) } else { $recordsWritten }
        }
    }
    finally {
        if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
            [System.IO.File]::Delete($tempPath)
        }
    }
}

Export-ModuleMember -Function Read-CsvRecords, Write-CsvRecord, Invoke-CsvJob
