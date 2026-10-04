[CmdletBinding()]
param(
    [ValidateRange(1, 300000)]
    [int] $RecordCount = 300000,
    [ValidateRange(0, 300000)]
    [int] $MappingCount = 0
)

$ErrorActionPreference = 'Stop'

function Release-BenchmarkComObject {
    param([object] $Object)

    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Write-BenchmarkRate {
    param(
        [string] $Label,
        [long] $RecordCount,
        [long] $ElapsedMilliseconds
    )

    $seconds = [Math]::Max($ElapsedMilliseconds / 1000.0, 0.001)
    $rate = [Math]::Round($RecordCount / $seconds, 2)
    Write-Output ('{0}: records={1}; seconds={2:N3}; records_per_second={3:N2}' -f $Label, $RecordCount, $seconds, $rate)
}

$root = Join-Path $env:TEMP ('MaskingBenchmark-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root | Out-Null

try {
    $moduleDirectory = Join-Path $PSScriptRoot '..\src\Masking'
    foreach ($moduleName in @('Mapping.psm1', 'Transform.psm1', 'CompiledCore.psm1', 'FixedWidth.psm1', 'Csv.psm1', 'Excel.psm1')) {
        Import-Module -Force -ErrorAction Stop -Name (Join-Path $moduleDirectory $moduleName)
    }

    $compilationTimer = [System.Diagnostics.Stopwatch]::StartNew()
    Import-MaskingCore
    $compilationTimer.Stop()
    Write-Output ('Core compilation: seconds={0:N3}; status=Loaded' -f ($compilationTimer.ElapsedMilliseconds / 1000.0))

    $recordCount = $RecordCount
    $mappingCount = if ($MappingCount -eq 0) { $recordCount } else { $MappingCount }
    if ($mappingCount -lt $recordCount) {
        throw 'MappingCount must be at least RecordCount'
    }
    $mapPath = Join-Path $root 'synthetic-map.tsv'
    $fixedInputPath = Join-Path $root 'synthetic-fixed.dat'
    $csvInputPath = Join-Path $root 'synthetic.csv'
    $fixedOutputPath = Join-Path $root 'synthetic-fixed.masked.dat'
    $csvOutputPath = Join-Path $root 'synthetic.masked.csv'
    $utf8 = [System.Text.UTF8Encoding]::new($false)
    $ascii = [System.Text.Encoding]::ASCII
    $generationTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $mapWriter = $null
    $fixedStream = $null
    $csvWriter = $null
    try {
        $mapWriter = [System.IO.StreamWriter]::new($mapPath, $false, $utf8)
        $fixedStream = [System.IO.FileStream]::new($fixedInputPath, [System.IO.FileMode]::CreateNew)
        $csvWriter = [System.IO.StreamWriter]::new($csvInputPath, $false, $utf8)
        for ($i = 0; $i -lt $mappingCount; $i++) {
            $sourceValue = '{0:D10}' -f ([long]1000000000 + $i)
            $maskedValue = '{0:D10}' -f ([long]2000000000 + $i)
            $mapWriter.WriteLine($sourceValue + "`t" + $maskedValue)
        }
        for ($i = 0; $i -lt $recordCount; $i++) {
            $sourceValue = '{0:D10}' -f ([long]1000000000 + $i)
            $fixedBytes = $ascii.GetBytes($sourceValue)
            $fixedStream.Write($fixedBytes, 0, $fixedBytes.Length)
            $csvWriter.WriteLine($sourceValue)
        }
    }
    finally {
        if ($null -ne $csvWriter) { $csvWriter.Dispose() }
        if ($null -ne $fixedStream) { $fixedStream.Dispose() }
        if ($null -ne $mapWriter) { $mapWriter.Dispose() }
    }
    $generationTimer.Stop()

    $mappingTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $mappingIndex = Read-MappingTable $mapPath $utf8 'None'
    $mappingTimer.Stop()
    $mappingIndexes = @{ MemberNumber = $mappingIndex }
    Write-Output ('Synthetic setup: records={0}; generation_seconds={1:N3}' -f
        $recordCount, ($generationTimer.ElapsedMilliseconds / 1000.0))
    Write-Output ('Mapping load: entries={0}; seconds={1:N3}' -f
        $mappingIndex.Count, ($mappingTimer.ElapsedMilliseconds / 1000.0))

    $fixedJob = [pscustomobject]@{
        InputPath = $fixedInputPath
        Mode = 'All'
        Encoding = '932'
        Layout = [pscustomobject]@{ RecordWidthBytes = 10; RecordTerminator = 'None' }
        Fields = @([pscustomobject]@{
            Name = 'MemberNumber'; StartByte = 1; ByteLength = 10
            Mapping = 'MemberNumber'; TrimSpaces = 'None'; PadSpace = 'Half'
        })
    }
    $fixedTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $fixedResult = Invoke-FixedWidthJob $fixedJob $mappingIndexes $null $fixedOutputPath
    $fixedTimer.Stop()
    if ([long]$fixedResult.RecordsWritten -ne $recordCount) { throw 'Fixed-width benchmark wrote an unexpected record count' }
    Write-BenchmarkRate 'FixedWidth' ([long]$fixedResult.RecordsWritten) $fixedTimer.ElapsedMilliseconds

    $csvJob = [pscustomobject]@{
        InputPath = $csvInputPath
        Mode = 'All'
        Encoding = '932'
        Layout = [pscustomobject]@{ Delimiter = ','; Quote = '"'; Header = $false; ContainerMode = 'Records' }
        Fields = @([pscustomobject]@{
            Name = 'MemberNumber'; ColumnIndex = 1; Mapping = 'MemberNumber'; TrimSpaces = 'None'
        })
    }
    $csvTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $csvResult = Invoke-CsvJob $csvJob $mappingIndexes $csvOutputPath
    $csvTimer.Stop()
    if ([long]$csvResult.RecordsWritten -ne $recordCount) { throw 'CSV benchmark wrote an unexpected record count' }
    Write-BenchmarkRate 'CSV' ([long]$csvResult.RecordsWritten) $csvTimer.ElapsedMilliseconds

    if (-not (Test-ExcelComAvailable)) {
        Write-Output 'Excel separate timing: BLOCKED (Office Excel COM is not registered; no Excel benchmark run).'
    }
    else {
        $excelInputPath = Join-Path $root 'synthetic.xlsx'
        $excelOutputPath = Join-Path $root 'synthetic.masked.xlsx'
        $excelRecordCount = [Math]::Min(1000, $recordCount)
        $excel = $null
        $workbook = $null
        $worksheet = $null
        $cells = $null
        $firstCell = $null
        $lastCell = $null
        $range = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.EnableEvents = $false
            $workbook = $excel.Workbooks.Add(-4167)
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Name = 'Members'
            $values = [object[,]]::new($excelRecordCount + 1, 1)
            $values[0, 0] = 'MemberNumber'
            for ($i = 0; $i -lt $excelRecordCount; $i++) {
                $values[$i + 1, 0] = '{0:D10}' -f ([long]1000000000 + $i)
            }
            $cells = $worksheet.Cells
            $firstCell = $cells.Item(1, 1)
            $lastCell = $cells.Item($excelRecordCount + 1, 1)
            $range = $worksheet.Range($firstCell, $lastCell)
            $range.NumberFormat = '@'
            $range.Value2 = $values
            $workbook.SaveAs($excelInputPath, 51)
            $workbook.Close($false)
            Release-BenchmarkComObject $range
            $range = $null
            Release-BenchmarkComObject $lastCell
            $lastCell = $null
            Release-BenchmarkComObject $firstCell
            $firstCell = $null
            Release-BenchmarkComObject $cells
            $cells = $null
            Release-BenchmarkComObject $worksheet
            $worksheet = $null
            Release-BenchmarkComObject $workbook
            $workbook = $null

            $excelJob = [pscustomobject]@{
                InputPath = $excelInputPath
                Mode = 'All'
                Layout = [pscustomobject]@{
                    Sheets = @([pscustomobject]@{
                        Name = 'Members'; Role = 'Data'; HeaderRow = 1; FirstDataRow = 2
                        Fields = @([pscustomobject]@{
                            Name = 'MemberNumber'; ColumnIndex = 1
                            Mapping = 'MemberNumber'; TrimSpaces = 'None'
                        })
                    })
                }
            }
            $excelTimer = [System.Diagnostics.Stopwatch]::StartNew()
            $excelResult = Invoke-ExcelJob $excelJob $mappingIndexes $excelOutputPath
            $excelTimer.Stop()
            $written = [long]$excelResult.Sheets[0].RecordsWritten
            if ($written -ne $excelRecordCount) { throw 'Excel benchmark wrote an unexpected record count' }
            Write-BenchmarkRate ('Excel separate (synthetic {0}-row workbook)' -f $excelRecordCount) $written $excelTimer.ElapsedMilliseconds
        }
        finally {
            if ($null -ne $workbook) { $workbook.Close($false) }
            if ($null -ne $excel) { $excel.Quit() }
            foreach ($comObject in @($range, $lastCell, $firstCell, $cells, $worksheet, $workbook, $excel)) {
                Release-BenchmarkComObject $comObject
            }
        }
    }
}
finally {
    Remove-Item -LiteralPath $root -Recurse -Force
}
