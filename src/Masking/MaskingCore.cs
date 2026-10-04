using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace FileMaskingTool.Core
{
    public sealed class SegmentRule
    {
        public int StartByte { get; set; }
        public int ByteLength { get; set; }
    }

    public sealed class FieldRule
    {
        public string Name { get; set; }
        public int StartByte { get; set; }
        public int ByteLength { get; set; }
        public string MappingName { get; set; }
        public string TrimSpaces { get; set; }
        public string PadSpace { get; set; }
        public SegmentRule[] Segments { get; set; }
    }

    public sealed class FixedWidthOptions
    {
        public string InputPath { get; set; }
        public string OutputPath { get; set; }
        public int RecordWidthBytes { get; set; }
        public byte[] InputTerminator { get; set; }
        public byte[] OutputTerminator { get; set; }
        public Encoding Encoding { get; set; }
        public FieldRule[] Fields { get; set; }
        public string Mode { get; set; }
        public int SampleCount { get; set; }
        public int Seed { get; set; }
        public int[] SelectedRecordIndexes { get; set; }
    }

    public sealed class FixedWidthResult
    {
        public long RecordsRead { get; set; }
        public long RecordsWritten { get; set; }
        public long RecordsInSource { get; set; }
        public long FieldsTransformed { get; set; }
        public long ElapsedMilliseconds { get; set; }
    }

    public sealed class CsvFieldRule
    {
        public string Name { get; set; }
        public int ColumnIndex { get; set; }
        public string ColumnName { get; set; }
        public string MappingName { get; set; }
        public string TrimSpaces { get; set; }
    }

    public sealed class CsvOptions
    {
        public string InputPath { get; set; }
        public string OutputPath { get; set; }
        public char Delimiter { get; set; }
        public char Quote { get; set; }
        public bool Header { get; set; }
        public string ContainerMode { get; set; }
        public int RecordWidthBytes { get; set; }
        public byte NullPadByte { get; set; }
        public byte[] RecordTerminator { get; set; }
        public Encoding Encoding { get; set; }
        public CsvFieldRule[] Fields { get; set; }
        public string Mode { get; set; }
        public int SampleCount { get; set; }
        public int Seed { get; set; }
    }

    public sealed class CsvResult
    {
        public long RecordsRead { get; set; }
        public long RecordsWritten { get; set; }
        public long DataRecordsWritten { get; set; }
        public long ElapsedMilliseconds { get; set; }
    }

    public static class MaskingCore
    {
        private sealed class PreparedField
        {
            public FieldRule Rule;
            public IDictionary Mapping;
            public byte[] PadBytes;
        }

        private sealed class ByteRange
        {
            public string Name;
            public int Start;
            public int End;
        }

        private sealed class CsvField
        {
            public string Value;
            public bool WasQuoted;
        }

        private sealed class CsvRecord
        {
            public List<CsvField> Fields = new List<CsvField>();
            public string Terminator;
        }

        private sealed class CsvInput
        {
            public FileStream Stream;
            public StreamReader Reader;
            public byte[] Preamble;
        }

        private sealed class PreparedCsvField
        {
            public CsvFieldRule Rule;
            public IDictionary Mapping;
            public int ColumnIndex;
        }

        public static FixedWidthResult ProcessFixedWidth(FixedWidthOptions options, IDictionary mappingIndexes)
        {
            ValidateOptions(options, mappingIndexes);
            Encoding encoding = Encoding.GetEncoding(
                options.Encoding.CodePage,
                EncoderFallback.ExceptionFallback,
                DecoderFallback.ExceptionFallback
            );
            PreparedField[] fields = PrepareFields(options, mappingIndexes, encoding);
            byte[] inputTerminator = options.InputTerminator ?? new byte[0];
            byte[] outputTerminator = options.OutputTerminator ?? new byte[0];
            long stride = (long)options.RecordWidthBytes + inputTerminator.Length;
            if (stride <= 0)
            {
                throw new InvalidDataException("Record stride is invalid");
            }
            if (String.Equals(
                Path.GetFullPath(options.InputPath),
                Path.GetFullPath(options.OutputPath),
                StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Input and output paths must differ");
            }

            FileStream input = null;
            FileStream output = null;
            Stopwatch stopwatch = Stopwatch.StartNew();
            bool completed = false;
            bool outputCreated = false;
            try
            {
                input = new FileStream(options.InputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                if ((input.Length % stride) != 0)
                {
                    throw new InvalidDataException("File size is not divisible by record stride");
                }
                long sourceCountLong = input.Length / stride;
                if (sourceCountLong == 0)
                {
                    throw new InvalidDataException("Fixed-width input contains no records");
                }
                if (sourceCountLong > Int32.MaxValue)
                {
                    throw new InvalidDataException("Fixed-width record count exceeds the supported limit");
                }
                int sourceCount = (int)sourceCountLong;

                ValidateInputTerminators(input, sourceCount, options.RecordWidthBytes, stride, inputTerminator);
                int[] selectedIndexes = options.SelectedRecordIndexes;
                if (selectedIndexes != null)
                {
                    if (selectedIndexes.Length == 0)
                    {
                        throw new InvalidDataException("SelectedRecordIndexes cannot be empty");
                    }
                    selectedIndexes = (int[])selectedIndexes.Clone();
                    Array.Sort(selectedIndexes);
                    int previousIndex = -1;
                    for (int i = 0; i < selectedIndexes.Length; i++)
                    {
                        if (selectedIndexes[i] < 0 || selectedIndexes[i] >= sourceCount)
                        {
                            throw new InvalidDataException("Selected record index is outside the input file");
                        }
                        if (selectedIndexes[i] == previousIndex)
                        {
                            throw new InvalidDataException("Selected record indexes contain duplicates");
                        }
                        previousIndex = selectedIndexes[i];
                    }
                }
                else if (String.Equals(options.Mode, "Sample", StringComparison.Ordinal))
                {
                    selectedIndexes = SelectSampleIndexes(sourceCount, options.SampleCount, options.Seed);
                }
                else if (!String.Equals(options.Mode, "All", StringComparison.Ordinal))
                {
                    throw new InvalidDataException("Unsupported fixed-width mode");
                }

                int recordsToWrite = selectedIndexes == null ? sourceCount : selectedIndexes.Length;
                long expectedLength = (long)recordsToWrite * ((long)options.RecordWidthBytes + outputTerminator.Length);
                if (expectedLength < 0)
                {
                    throw new InvalidDataException("Fixed-width output size is invalid");
                }

                output = new FileStream(options.OutputPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                outputCreated = true;
                byte[] record = new byte[options.RecordWidthBytes];
                if (selectedIndexes == null)
                {
                    input.Position = 0;
                    for (int recordIndex = 0; recordIndex < sourceCount; recordIndex++)
                    {
                        ReadExactly(input, record, "record " + (recordIndex + 1));
                        if (inputTerminator.Length > 0)
                        {
                            ReadAndValidateTerminator(input, inputTerminator, recordIndex + 1);
                        }
                        TransformRecord(record, fields, encoding, recordIndex + 1);
                        output.Write(record, 0, record.Length);
                        if (outputTerminator.Length > 0)
                        {
                            output.Write(outputTerminator, 0, outputTerminator.Length);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < selectedIndexes.Length; i++)
                    {
                        int recordIndex = selectedIndexes[i];
                        input.Position = (long)recordIndex * stride;
                        ReadExactly(input, record, "record " + (recordIndex + 1));
                        if (inputTerminator.Length > 0)
                        {
                            ReadAndValidateTerminator(input, inputTerminator, recordIndex + 1);
                        }
                        TransformRecord(record, fields, encoding, recordIndex + 1);
                        output.Write(record, 0, record.Length);
                        if (outputTerminator.Length > 0)
                        {
                            output.Write(outputTerminator, 0, outputTerminator.Length);
                        }
                    }
                }

                output.Flush();
                if (output.Length != expectedLength)
                {
                    throw new InvalidDataException("Fixed-width output size validation failed");
                }
                stopwatch.Stop();
                completed = true;
                return new FixedWidthResult
                {
                    RecordsRead = recordsToWrite,
                    RecordsWritten = recordsToWrite,
                    RecordsInSource = sourceCount,
                    FieldsTransformed = (long)recordsToWrite * options.Fields.Length,
                    ElapsedMilliseconds = stopwatch.ElapsedMilliseconds
                };
            }
            finally
            {
                if (output != null)
                {
                    output.Dispose();
                }
                if (input != null)
                {
                    input.Dispose();
                }
                if (outputCreated && !completed && File.Exists(options.OutputPath))
                {
                    File.Delete(options.OutputPath);
                }
            }
        }

        public static CsvResult ProcessCsv(CsvOptions options, IDictionary mappingIndexes)
        {
            ValidateCsvOptions(options, mappingIndexes);
            Encoding encoding = Encoding.GetEncoding(
                options.Encoding.CodePage,
                EncoderFallback.ExceptionFallback,
                DecoderFallback.ExceptionFallback
            );
            byte[] preamble = GetCsvPreamble(options.InputPath, encoding);
            bool fixedWidth = String.Equals(options.ContainerMode, "FixedWidthNullPadded", StringComparison.Ordinal);
            HashSet<int> selectedIndexes = null;
            try
            {
                if (String.Equals(options.Mode, "Sample", StringComparison.Ordinal))
                {
                    selectedIndexes = new HashSet<int>(SelectCsvSampleIndexes(options, encoding, mappingIndexes));
                }
            }
            catch (DecoderFallbackException)
            {
                throw new InvalidDataException("CSV input contains bytes invalid for the configured encoding");
            }

            FileStream output = null;
            StreamWriter writer = null;
            Stopwatch stopwatch = Stopwatch.StartNew();
            bool completed = false;
            bool outputCreated = false;
            long recordsRead = 0;
            long recordsWritten = 0;
            long dataRecordsWritten = 0;
            try
            {
                output = new FileStream(options.OutputPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                outputCreated = true;
                CsvRecord header = null;
                PreparedCsvField[] fields = null;
                int expectedColumnCount = -1;
                int dataIndex = 0;

                if (fixedWidth)
                {
                    long stride = (long)options.RecordWidthBytes + options.RecordTerminator.Length;
                    using (FileStream input = new FileStream(options.InputPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        if ((input.Length % stride) != 0)
                        {
                            throw new InvalidDataException("CSV fixed-width file length is not a whole number of records");
                        }
                        int recordNumber = 0;
                        while (input.Position < input.Length)
                        {
                            recordNumber++;
                            CsvRecord record = ReadFixedCsvRecord(input, options, encoding, recordNumber);
                            recordsRead++;
                            bool isHeader = options.Header && recordNumber == 1;
                            if (isHeader)
                            {
                                header = record;
                                expectedColumnCount = record.Fields.Count;
                                fields = PrepareCsvFields(options, mappingIndexes, header);
                                WriteFixedCsvRecord(output, record, options, encoding, recordNumber);
                                recordsWritten++;
                                continue;
                            }

                            if (fields == null)
                            {
                                fields = PrepareCsvFields(options, mappingIndexes, null);
                            }
                            ValidateCsvDataRecord(record, fields, ref expectedColumnCount, recordNumber);
                            bool shouldWrite = selectedIndexes == null || selectedIndexes.Contains(dataIndex);
                            if (shouldWrite)
                            {
                                TransformCsvRecord(record, fields, options);
                                WriteFixedCsvRecord(output, record, options, encoding, recordNumber);
                                recordsWritten++;
                                dataRecordsWritten++;
                            }
                            dataIndex++;
                        }
                    }
                    if (options.Header && header == null)
                    {
                        throw new InvalidDataException("CSV header row is missing");
                    }
                }
                else
                {
                    CsvInput csvInput = OpenCsvInput(options.InputPath, encoding);
                    try
                    {
                        if (csvInput.Preamble.Length > 0)
                        {
                            output.Write(csvInput.Preamble, 0, csvInput.Preamble.Length);
                        }
                        writer = new StreamWriter(output, GetCsvWriterEncoding(encoding), 65536, true);
                        CsvRecord record;
                        while ((record = ReadCsvRecord(csvInput.Reader, options.Delimiter, options.Quote)) != null)
                        {
                            recordsRead++;
                            bool isHeader = options.Header && header == null;
                            if (isHeader)
                            {
                                header = record;
                                expectedColumnCount = record.Fields.Count;
                                fields = PrepareCsvFields(options, mappingIndexes, header);
                                WriteCsvRecord(writer, record, options);
                                recordsWritten++;
                                continue;
                            }

                            if (fields == null)
                            {
                                fields = PrepareCsvFields(options, mappingIndexes, null);
                            }
                            ValidateCsvDataRecord(record, fields, ref expectedColumnCount, dataIndex + 1);
                            bool shouldWrite = selectedIndexes == null || selectedIndexes.Contains(dataIndex);
                            if (shouldWrite)
                            {
                                TransformCsvRecord(record, fields, options);
                                WriteCsvRecord(writer, record, options);
                                recordsWritten++;
                                dataRecordsWritten++;
                            }
                            dataIndex++;
                        }
                    }
                    finally
                    {
                        csvInput.Reader.Dispose();
                        csvInput.Stream.Dispose();
                    }
                    if (options.Header && header == null)
                    {
                        throw new InvalidDataException("CSV header row is missing");
                    }
                }

                if (writer != null)
                {
                    writer.Flush();
                    writer.Dispose();
                    writer = null;
                }
                output.Flush();
                output.Dispose();
                output = null;
                stopwatch.Stop();
                completed = true;
                return new CsvResult
                {
                    RecordsRead = recordsRead,
                    RecordsWritten = recordsWritten,
                    DataRecordsWritten = dataRecordsWritten,
                    ElapsedMilliseconds = stopwatch.ElapsedMilliseconds
                };
            }
            catch (DecoderFallbackException)
            {
                throw new InvalidDataException("CSV input contains bytes invalid for the configured encoding");
            }
            catch (EncoderFallbackException)
            {
                throw new InvalidDataException("CSV output cannot be represented in the configured encoding");
            }
            finally
            {
                if (writer != null)
                {
                    writer.Dispose();
                }
                if (output != null)
                {
                    output.Dispose();
                }
                if (outputCreated && !completed && File.Exists(options.OutputPath))
                {
                    File.Delete(options.OutputPath);
                }
            }
        }

        private static void ValidateCsvOptions(CsvOptions options, IDictionary mappingIndexes)
        {
            if (options == null)
            {
                throw new ArgumentNullException("options");
            }
            if (String.IsNullOrWhiteSpace(options.InputPath) || !File.Exists(options.InputPath))
            {
                throw new FileNotFoundException("CSV input file does not exist");
            }
            if (String.IsNullOrWhiteSpace(options.OutputPath))
            {
                throw new InvalidDataException("CSV output path is required");
            }
            if (String.Equals(
                Path.GetFullPath(options.InputPath),
                Path.GetFullPath(options.OutputPath),
                StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Input and output paths must differ");
            }
            if (options.Encoding == null)
            {
                throw new InvalidDataException("CSV encoding is required");
            }
            if (options.Encoding.CodePage != 932)
            {
                throw new InvalidDataException("CSV job Encoding must be cp932 (code page 932)");
            }
            if (options.Delimiter == options.Quote || options.Delimiter == '\r' || options.Delimiter == '\n' ||
                options.Quote == '\r' || options.Quote == '\n')
            {
                throw new InvalidDataException("CSV dialect is invalid");
            }
            if (options.Fields == null || options.Fields.Length == 0)
            {
                throw new InvalidDataException("CSV job must define fields");
            }
            if (mappingIndexes == null)
            {
                throw new InvalidDataException("Mapping indexes are required");
            }
            if (!String.Equals(options.Mode, "All", StringComparison.Ordinal) &&
                !String.Equals(options.Mode, "Sample", StringComparison.Ordinal))
            {
                throw new InvalidDataException("CSV Mode must be Sample or All");
            }
            if (String.Equals(options.Mode, "Sample", StringComparison.Ordinal) && options.SampleCount <= 0)
            {
                throw new InvalidDataException("SampleCount must be positive");
            }
            if (!String.Equals(options.ContainerMode, "Records", StringComparison.Ordinal) &&
                !String.Equals(options.ContainerMode, "FixedWidthNullPadded", StringComparison.Ordinal))
            {
                throw new InvalidDataException("Unsupported CSV container mode");
            }
            if (String.Equals(options.ContainerMode, "FixedWidthNullPadded", StringComparison.Ordinal))
            {
                if (options.RecordWidthBytes <= 0)
                {
                    throw new InvalidDataException("CSV fixed-width record size must be positive");
                }
                if (options.NullPadByte != 0)
                {
                    throw new InvalidDataException("CSV NullPadByte must be zero");
                }
                if (options.RecordTerminator == null)
                {
                    throw new InvalidDataException("CSV fixed-width record terminator is required");
                }
                if (options.RecordTerminator.Length > 2 ||
                    (options.RecordTerminator.Length == 1 && options.RecordTerminator[0] != 10) ||
                    (options.RecordTerminator.Length == 2 &&
                    (options.RecordTerminator[0] != 13 || options.RecordTerminator[1] != 10)))
                {
                    throw new InvalidDataException("Unsupported CSV fixed-width record terminator");
                }
                if (options.Encoding.CodePage == 1200 || options.Encoding.CodePage == 1201 ||
                    options.Encoding.CodePage == 12000 || options.Encoding.CodePage == 12001)
                {
                    throw new InvalidDataException("Fixed-width NUL-padded CSV does not support UTF-16 or UTF-32 encodings");
                }
            }

            for (int i = 0; i < options.Fields.Length; i++)
            {
                CsvFieldRule field = options.Fields[i];
                if (field == null || String.IsNullOrWhiteSpace(field.Name) ||
                    String.IsNullOrWhiteSpace(field.MappingName))
                {
                    throw new InvalidDataException("CSV field rule is invalid");
                }
                if (String.IsNullOrWhiteSpace(field.ColumnName) && field.ColumnIndex < 1)
                {
                    throw new InvalidDataException("CSV ColumnIndex must be positive");
                }
                if (!mappingIndexes.Contains(field.MappingName) || !(mappingIndexes[field.MappingName] is IDictionary))
                {
                    throw new InvalidDataException("CSV field " + field.Name + " refers to an unavailable mapping index");
                }
                switch (field.TrimSpaces)
                {
                    case "None":
                    case "Half":
                    case "Full":
                    case "HalfAndFull":
                        break;
                    default:
                        throw new InvalidDataException("CSV field " + field.Name + " has an unsupported trim mode");
                }
            }
        }

        private static int[] SelectCsvSampleIndexes(CsvOptions options, Encoding encoding, IDictionary mappingIndexes)
        {
            List<int> selected = new List<int>(Math.Min(options.SampleCount, 1024));
            Random random = new Random(options.Seed);
            int dataIndex = 0;
            int expectedColumnCount = -1;
            PreparedCsvField[] fields = null;
            CsvRecord header = null;
            bool fixedWidth = String.Equals(options.ContainerMode, "FixedWidthNullPadded", StringComparison.Ordinal);
            if (fixedWidth)
            {
                long stride = (long)options.RecordWidthBytes + options.RecordTerminator.Length;
                using (FileStream input = new FileStream(options.InputPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    if ((input.Length % stride) != 0)
                    {
                        throw new InvalidDataException("CSV fixed-width file length is not a whole number of records");
                    }
                    int recordNumber = 0;
                    while (input.Position < input.Length)
                    {
                        recordNumber++;
                        CsvRecord record = ReadFixedCsvRecord(input, options, encoding, recordNumber);
                        if (options.Header && recordNumber == 1)
                        {
                            header = record;
                            expectedColumnCount = record.Fields.Count;
                            fields = PrepareCsvFields(options, mappingIndexes, header);
                            continue;
                        }
                        if (fields == null)
                        {
                            fields = PrepareCsvFields(options, mappingIndexes, null);
                        }
                        ValidateCsvDataRecord(record, fields, ref expectedColumnCount, dataIndex + 1);
                        ReservoirAdd(selected, dataIndex, options.SampleCount, random);
                        dataIndex++;
                    }
                }
            }
            else
            {
                CsvInput csvInput = OpenCsvInput(options.InputPath, encoding);
                try
                {
                    CsvRecord record;
                    int recordNumber = 0;
                    while ((record = ReadCsvRecord(csvInput.Reader, options.Delimiter, options.Quote)) != null)
                    {
                        recordNumber++;
                        if (options.Header && recordNumber == 1)
                        {
                            header = record;
                            expectedColumnCount = record.Fields.Count;
                            fields = PrepareCsvFields(options, mappingIndexes, header);
                            continue;
                        }
                        if (fields == null)
                        {
                            fields = PrepareCsvFields(options, mappingIndexes, null);
                        }
                        ValidateCsvDataRecord(record, fields, ref expectedColumnCount, dataIndex + 1);
                        ReservoirAdd(selected, dataIndex, options.SampleCount, random);
                        dataIndex++;
                    }
                }
                finally
                {
                    csvInput.Reader.Dispose();
                    csvInput.Stream.Dispose();
                }
            }
            if (options.Header && header == null)
            {
                throw new InvalidDataException("CSV header row is missing");
            }
            selected.Sort();
            return selected.ToArray();
        }

        private static void ReservoirAdd(List<int> selected, int index, int sampleCount, Random random)
        {
            if (selected.Count < sampleCount)
            {
                selected.Add(index);
            }
            else
            {
                int replacementIndex = random.Next(0, index + 1);
                if (replacementIndex < sampleCount)
                {
                    selected[replacementIndex] = index;
                }
            }
        }

        private static PreparedCsvField[] PrepareCsvFields(
            CsvOptions options,
            IDictionary mappingIndexes,
            CsvRecord header)
        {
            PreparedCsvField[] fields = new PreparedCsvField[options.Fields.Length];
            HashSet<int> usedColumns = new HashSet<int>();
            for (int i = 0; i < options.Fields.Length; i++)
            {
                CsvFieldRule rule = options.Fields[i];
                int columnIndex;
                if (!String.IsNullOrWhiteSpace(rule.ColumnName))
                {
                    if (header == null)
                    {
                        throw new InvalidDataException("CSV field " + rule.Name + " requires a header row");
                    }
                    int matches = 0;
                    columnIndex = -1;
                    for (int j = 0; j < header.Fields.Count; j++)
                    {
                        if (String.Equals(header.Fields[j].Value, rule.ColumnName, StringComparison.Ordinal))
                        {
                            matches++;
                            columnIndex = j;
                        }
                    }
                    if (matches != 1)
                    {
                        throw new InvalidDataException("CSV header must contain exactly one configured column name");
                    }
                }
                else
                {
                    columnIndex = rule.ColumnIndex - 1;
                }
                if (header != null && columnIndex >= header.Fields.Count)
                {
                    throw new InvalidDataException("CSV header does not contain a configured column");
                }
                if (!usedColumns.Add(columnIndex))
                {
                    throw new InvalidDataException("CSV fields target the same column index");
                }
                fields[i] = new PreparedCsvField
                {
                    Rule = rule,
                    Mapping = mappingIndexes[rule.MappingName] as IDictionary,
                    ColumnIndex = columnIndex
                };
            }
            return fields;
        }

        private static void ValidateCsvDataRecord(
            CsvRecord record,
            PreparedCsvField[] fields,
            ref int expectedColumnCount,
            int recordNumber)
        {
            if (expectedColumnCount < 0)
            {
                expectedColumnCount = record.Fields.Count;
            }
            if (record.Fields.Count != expectedColumnCount)
            {
                throw new InvalidDataException("CSV record " + recordNumber + " has an inconsistent column count");
            }
            for (int i = 0; i < fields.Length; i++)
            {
                if (fields[i].ColumnIndex < 0 || fields[i].ColumnIndex >= record.Fields.Count)
                {
                    throw new InvalidDataException(
                        "CSV record " + recordNumber + " has no configured column " + (fields[i].ColumnIndex + 1)
                    );
                }
            }
        }

        private static void TransformCsvRecord(CsvRecord record, PreparedCsvField[] fields, CsvOptions options)
        {
            for (int i = 0; i < fields.Length; i++)
            {
                PreparedCsvField field = fields[i];
                CsvField original = record.Fields[field.ColumnIndex];
                string replacement = GetMappedValue(
                    field.Mapping,
                    original.Value,
                    field.Rule.TrimSpaces,
                    field.Rule.Name
                );
                record.Fields[field.ColumnIndex] = new CsvField
                {
                    Value = replacement,
                    WasQuoted = original.WasQuoted
                };
            }
        }

        private static CsvInput OpenCsvInput(string path, Encoding encoding)
        {
            FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            try
            {
                byte[] preamble = GetCsvPreamble(stream, encoding);
                stream.Position = preamble.Length;
                StreamReader reader = new StreamReader(stream, encoding, false, 65536, true);
                return new CsvInput
                {
                    Stream = stream,
                    Reader = reader,
                    Preamble = preamble
                };
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        private static byte[] GetCsvPreamble(string path, Encoding encoding)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                return GetCsvPreamble(stream, encoding);
            }
        }

        private static byte[] GetCsvPreamble(Stream stream, Encoding encoding)
        {
            long originalPosition = stream.Position;
            stream.Position = 0;
            byte[] first = new byte[4];
            int read = 0;
            while (read < first.Length)
            {
                int bytesRead = stream.Read(first, read, first.Length - read);
                if (bytesRead <= 0)
                {
                    break;
                }
                read += bytesRead;
            }
            stream.Position = originalPosition;
            int codePage = 0;
            int length = 0;
            if (read >= 4 && first[0] == 0xFF && first[1] == 0xFE && first[2] == 0 && first[3] == 0)
            {
                codePage = 12000;
                length = 4;
            }
            else if (read >= 4 && first[0] == 0 && first[1] == 0 && first[2] == 0xFE && first[3] == 0xFF)
            {
                codePage = 12001;
                length = 4;
            }
            else if (read >= 3 && first[0] == 0xEF && first[1] == 0xBB && first[2] == 0xBF)
            {
                codePage = 65001;
                length = 3;
            }
            else if (read >= 2 && first[0] == 0xFE && first[1] == 0xFF)
            {
                codePage = 1201;
                length = 2;
            }
            else if (read >= 2 && first[0] == 0xFF && first[1] == 0xFE)
            {
                codePage = 1200;
                length = 2;
            }
            if (codePage != 0 && codePage != encoding.CodePage)
            {
                throw new InvalidDataException("CSV byte-order mark does not match the configured encoding");
            }
            if (length == 0)
            {
                return new byte[0];
            }
            byte[] preamble = new byte[length];
            Buffer.BlockCopy(first, 0, preamble, 0, length);
            return preamble;
        }

        private static Encoding GetCsvWriterEncoding(Encoding encoding)
        {
            switch (encoding.CodePage)
            {
                case 65001:
                    return new UTF8Encoding(false, true);
                case 1200:
                    return new UnicodeEncoding(false, false, true);
                case 1201:
                    return new UnicodeEncoding(true, false, true);
                case 12000:
                    return new UTF32Encoding(false, false, true);
                case 12001:
                    return new UTF32Encoding(true, false, true);
                default:
                    return Encoding.GetEncoding(
                        encoding.CodePage,
                        EncoderFallback.ExceptionFallback,
                        DecoderFallback.ExceptionFallback
                    );
            }
        }

        private static CsvRecord ReadCsvRecord(TextReader reader, char delimiter, char quote)
        {
            List<CsvField> fields = new List<CsvField>();
            StringBuilder value = new StringBuilder();
            bool wasQuoted = false;
            bool inQuotes = false;
            bool afterQuote = false;
            bool recordStarted = false;
            while (true)
            {
                int code = reader.Read();
                if (code < 0)
                {
                    if (inQuotes)
                    {
                        throw new InvalidDataException("CSV contains an unterminated quoted field");
                    }
                    if (!recordStarted && fields.Count == 0 && value.Length == 0)
                    {
                        return null;
                    }
                    fields.Add(new CsvField { Value = value.ToString(), WasQuoted = wasQuoted });
                    return new CsvRecord { Fields = fields, Terminator = String.Empty };
                }

                char character = (char)code;
                recordStarted = true;
                if (inQuotes)
                {
                    if (character == quote)
                    {
                        if (reader.Peek() == (int)quote)
                        {
                            reader.Read();
                            value.Append(quote);
                        }
                        else
                        {
                            inQuotes = false;
                            afterQuote = true;
                        }
                    }
                    else
                    {
                        value.Append(character);
                    }
                    continue;
                }

                if (afterQuote && character != delimiter && character != '\r' && character != '\n')
                {
                    throw new InvalidDataException("CSV has data after a closing quote");
                }
                if (character == delimiter)
                {
                    fields.Add(new CsvField { Value = value.ToString(), WasQuoted = wasQuoted });
                    value.Length = 0;
                    wasQuoted = false;
                    afterQuote = false;
                    continue;
                }
                if (character == '\r' || character == '\n')
                {
                    string terminator = character.ToString();
                    if (character == '\r' && reader.Peek() == 10)
                    {
                        reader.Read();
                        terminator = "\r\n";
                    }
                    fields.Add(new CsvField { Value = value.ToString(), WasQuoted = wasQuoted });
                    return new CsvRecord { Fields = fields, Terminator = terminator };
                }
                if (character == quote)
                {
                    if (afterQuote || value.Length != 0)
                    {
                        throw new InvalidDataException("CSV quote occurs inside an unquoted field");
                    }
                    wasQuoted = true;
                    inQuotes = true;
                    continue;
                }
                value.Append(character);
            }
        }

        private static CsvRecord ReadFixedCsvRecord(
            Stream stream,
            CsvOptions options,
            Encoding encoding,
            int recordNumber)
        {
            byte[] block = new byte[options.RecordWidthBytes];
            ReadExactly(stream, block, "CSV fixed-width record " + recordNumber);
            byte[] terminator = options.RecordTerminator ?? new byte[0];
            if (terminator.Length > 0)
            {
                byte[] actual = new byte[terminator.Length];
                ReadExactly(stream, actual, "CSV fixed-width record " + recordNumber + " terminator");
                for (int i = 0; i < terminator.Length; i++)
                {
                    if (actual[i] != terminator[i])
                    {
                        throw new InvalidDataException("CSV fixed-width record " + recordNumber + " has an invalid terminator");
                    }
                }
            }
            int payloadLength = block.Length;
            while (payloadLength > 0 && block[payloadLength - 1] == options.NullPadByte)
            {
                payloadLength--;
            }
            for (int i = 0; i < payloadLength; i++)
            {
                if (block[i] == options.NullPadByte)
                {
                    throw new InvalidDataException("NUL byte occurs inside CSV payload");
                }
            }
            string payload = encoding.GetString(block, 0, payloadLength);
            using (StringReader reader = new StringReader(payload))
            {
                CsvRecord record = ReadCsvRecord(reader, options.Delimiter, options.Quote);
                CsvRecord extra = ReadCsvRecord(reader, options.Delimiter, options.Quote);
                if (record == null || extra != null || record.Terminator.Length != 0)
                {
                    throw new InvalidDataException("CSV fixed-width record " + recordNumber + " contains more than one CSV record");
                }
                return record;
            }
        }

        private static string SerializeCsvRecord(CsvRecord record, char delimiter, char quote)
        {
            StringBuilder output = new StringBuilder();
            for (int i = 0; i < record.Fields.Count; i++)
            {
                if (i > 0)
                {
                    output.Append(delimiter);
                }
                CsvField field = record.Fields[i];
                string value = field.Value;
                bool requiresQuotes = value.IndexOf(delimiter) >= 0 || value.IndexOf(quote) >= 0 ||
                    value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0;
                if (field.WasQuoted)
                {
                    output.Append(quote);
                    for (int j = 0; j < value.Length; j++)
                    {
                        if (value[j] == quote)
                        {
                            output.Append(quote);
                            output.Append(quote);
                        }
                        else
                        {
                            output.Append(value[j]);
                        }
                    }
                    output.Append(quote);
                }
                else if (requiresQuotes)
                {
                    throw new InvalidDataException("Unquoted CSV field requires quoting");
                }
                else
                {
                    output.Append(value);
                }
            }
            return output.ToString();
        }

        private static void WriteCsvRecord(StreamWriter writer, CsvRecord record, CsvOptions options)
        {
            writer.Write(SerializeCsvRecord(record, options.Delimiter, options.Quote));
            writer.Write(record.Terminator);
        }

        private static void WriteFixedCsvRecord(
            Stream output,
            CsvRecord record,
            CsvOptions options,
            Encoding encoding,
            int recordNumber)
        {
            string text = SerializeCsvRecord(record, options.Delimiter, options.Quote);
            byte[] payload = encoding.GetBytes(text);
            if (payload.Length > options.RecordWidthBytes)
            {
                throw new InvalidDataException("CSV record " + recordNumber + " exceeds fixed-width CSV record payload");
            }
            output.Write(payload, 0, payload.Length);
            for (int i = payload.Length; i < options.RecordWidthBytes; i++)
            {
                output.WriteByte(options.NullPadByte);
            }
            byte[] terminator = options.RecordTerminator ?? new byte[0];
            if (terminator.Length > 0)
            {
                output.Write(terminator, 0, terminator.Length);
            }
        }

        private static void ValidateOptions(FixedWidthOptions options, IDictionary mappingIndexes)
        {
            if (options == null)
            {
                throw new ArgumentNullException("options");
            }
            if (String.IsNullOrWhiteSpace(options.InputPath) || !File.Exists(options.InputPath))
            {
                throw new FileNotFoundException("Fixed-width input file does not exist");
            }
            if (String.IsNullOrWhiteSpace(options.OutputPath))
            {
                throw new InvalidDataException("Fixed-width output path is required");
            }
            if (options.RecordWidthBytes <= 0)
            {
                throw new InvalidDataException("RecordWidthBytes must be positive");
            }
            if (options.Encoding == null)
            {
                throw new InvalidDataException("Fixed-width encoding is required");
            }
            if (options.Encoding.CodePage != 932)
            {
                throw new InvalidDataException("Fixed-width job Encoding must be cp932 (code page 932)");
            }
            if (options.Fields == null || options.Fields.Length == 0)
            {
                throw new InvalidDataException("Fixed-width job must define fields");
            }
            if (mappingIndexes == null)
            {
                throw new InvalidDataException("Mapping indexes are required");
            }

            List<ByteRange> ranges = new List<ByteRange>();
            for (int i = 0; i < options.Fields.Length; i++)
            {
                FieldRule field = options.Fields[i];
                if (field == null || String.IsNullOrWhiteSpace(field.Name) || field.ByteLength <= 0)
                {
                    throw new InvalidDataException("Fixed-width field rule is invalid");
                }
                if (String.IsNullOrWhiteSpace(field.MappingName) || !mappingIndexes.Contains(field.MappingName))
                {
                    throw new InvalidDataException("Field " + field.Name + " refers to an unavailable mapping index");
                }
                if (!(mappingIndexes[field.MappingName] is IDictionary))
                {
                    throw new InvalidDataException("Field " + field.Name + " mapping index is invalid");
                }

                SegmentRule[] segments = field.Segments ?? new SegmentRule[0];
                if (segments.Length == 0)
                {
                    AddRange(ranges, field.Name, field.StartByte, field.ByteLength, options.RecordWidthBytes);
                }
                else
                {
                    if (segments.Length < 2)
                    {
                        throw new InvalidDataException("Composite field " + field.Name + " requires at least two segments");
                    }
                    int combinedLength = 0;
                    for (int segmentIndex = 0; segmentIndex < segments.Length; segmentIndex++)
                    {
                        SegmentRule segment = segments[segmentIndex];
                        if (segment == null)
                        {
                            throw new InvalidDataException("Composite field " + field.Name + " contains an invalid segment");
                        }
                        AddRange(
                            ranges,
                            field.Name + "." + (segmentIndex + 1),
                            segment.StartByte,
                            segment.ByteLength,
                            options.RecordWidthBytes
                        );
                        combinedLength += segment.ByteLength;
                    }
                    if (combinedLength != field.ByteLength)
                    {
                        throw new InvalidDataException("Composite field " + field.Name + " byte length does not match its segments");
                    }
                }
            }

            for (int i = 0; i < ranges.Count; i++)
            {
                for (int j = i + 1; j < ranges.Count; j++)
                {
                    if (ranges[i].Start <= ranges[j].End && ranges[j].Start <= ranges[i].End)
                    {
                        throw new InvalidDataException("Fixed-width fields overlap");
                    }
                }
            }
        }

        private static PreparedField[] PrepareFields(
            FixedWidthOptions options,
            IDictionary mappingIndexes,
            Encoding encoding)
        {
            PreparedField[] prepared = new PreparedField[options.Fields.Length];
            for (int i = 0; i < options.Fields.Length; i++)
            {
                FieldRule field = options.Fields[i];
                IDictionary mapping = mappingIndexes[field.MappingName] as IDictionary;
                byte[] padBytes = new byte[0];
                if (field.Segments == null || field.Segments.Length == 0)
                {
                    char padCharacter;
                    if (String.Equals(field.PadSpace, "Half", StringComparison.Ordinal))
                    {
                        padCharacter = ' ';
                    }
                    else if (String.Equals(field.PadSpace, "Full", StringComparison.Ordinal))
                    {
                        padCharacter = '\u3000';
                    }
                    else
                    {
                        throw new InvalidDataException("Field " + field.Name + " PadSpace must be Half or Full");
                    }
                    padBytes = encoding.GetBytes(new char[] { padCharacter });
                    if (padBytes.Length == 0)
                    {
                        throw new InvalidDataException("Field " + field.Name + " has an invalid padding character");
                    }
                }
                prepared[i] = new PreparedField
                {
                    Rule = field,
                    Mapping = mapping,
                    PadBytes = padBytes
                };
            }
            return prepared;
        }

        private static void AddRange(List<ByteRange> ranges, string name, int startByte, int byteLength, int recordWidth)
        {
            if (startByte < 1 || byteLength < 1 || ((long)startByte - 1 + byteLength) > recordWidth)
            {
                throw new InvalidDataException("Field " + name + " exceeds record width");
            }
            ranges.Add(new ByteRange
            {
                Name = name,
                Start = startByte,
                End = startByte + byteLength - 1
            });
        }

        private static void ValidateInputTerminators(
            FileStream input,
            int recordCount,
            int recordWidth,
            long stride,
            byte[] terminator)
        {
            if (terminator.Length == 0)
            {
                return;
            }
            byte[] actual = new byte[terminator.Length];
            for (int i = 0; i < recordCount; i++)
            {
                input.Position = ((long)i * stride) + recordWidth;
                ReadExactly(input, actual, "record " + (i + 1) + " terminator");
                for (int j = 0; j < terminator.Length; j++)
                {
                    if (actual[j] != terminator[j])
                    {
                        throw new InvalidDataException("Record " + (i + 1) + " record terminator mismatch");
                    }
                }
            }
        }

        private static int[] SelectSampleIndexes(int recordCount, int sampleCount, int seed)
        {
            if (sampleCount <= 0)
            {
                throw new InvalidDataException("SampleCount must be positive");
            }
            int size = Math.Min(recordCount, sampleCount);
            List<int> selected = new List<int>(size);
            Random random = new Random(seed);
            for (int i = 0; i < recordCount; i++)
            {
                if (selected.Count < size)
                {
                    selected.Add(i);
                }
                else
                {
                    int replacementIndex = random.Next(0, i + 1);
                    if (replacementIndex < size)
                    {
                        selected[replacementIndex] = i;
                    }
                }
            }
            selected.Sort();
            return selected.ToArray();
        }

        private static void ReadExactly(Stream input, byte[] buffer, string context)
        {
            int offset = 0;
            while (offset < buffer.Length)
            {
                int read = input.Read(buffer, offset, buffer.Length - offset);
                if (read <= 0)
                {
                    throw new EndOfStreamException("Unexpected end of file while reading " + context);
                }
                offset += read;
            }
        }

        private static void ReadAndValidateTerminator(Stream input, byte[] expected, int recordNumber)
        {
            byte[] actual = new byte[expected.Length];
            ReadExactly(input, actual, "record " + recordNumber + " terminator");
            for (int i = 0; i < expected.Length; i++)
            {
                if (actual[i] != expected[i])
                {
                    throw new InvalidDataException("Record " + recordNumber + " record terminator mismatch");
                }
            }
        }

        private static void TransformRecord(byte[] record, PreparedField[] fields, Encoding encoding, int recordNumber)
        {
            for (int i = 0; i < fields.Length; i++)
            {
                PreparedField prepared = fields[i];
                FieldRule field = prepared.Rule;
                SegmentRule[] segments = field.Segments ?? new SegmentRule[0];
                try
                {
                    if (segments.Length == 0)
                    {
                        TransformSimpleField(record, prepared, encoding);
                    }
                    else
                    {
                        TransformCompositeField(record, prepared, encoding);
                    }
                }
                catch (DecoderFallbackException)
                {
                    throw new InvalidDataException(
                        "Record " + recordNumber + " field " + field.Name + " contains invalid encoded bytes"
                    );
                }
                catch (EncoderFallbackException)
                {
                    throw new InvalidDataException(
                        "Record " + recordNumber + " field " + field.Name + " replacement cannot be represented in the configured encoding"
                    );
                }
                catch (InvalidDataException)
                {
                    throw;
                }
                catch
                {
                    throw new InvalidDataException(
                        "Record " + recordNumber + " field " + field.Name + " transformation failed"
                    );
                }
            }
        }

        private static void TransformSimpleField(byte[] record, PreparedField prepared, Encoding encoding)
        {
            FieldRule field = prepared.Rule;
            string value = encoding.GetString(record, field.StartByte - 1, field.ByteLength);
            string replacement = GetMappedValue(prepared.Mapping, value, field.TrimSpaces, field.Name);
            byte[] replacementBytes = encoding.GetBytes(replacement);
            int remaining = field.ByteLength - replacementBytes.Length;
            if (remaining < 0)
            {
                throw new InvalidDataException("Field " + field.Name + " mapped value exceeds field width");
            }
            if ((remaining % prepared.PadBytes.Length) != 0)
            {
                throw new InvalidDataException("Field " + field.Name + " padding cannot be filled exactly");
            }
            Buffer.BlockCopy(replacementBytes, 0, record, field.StartByte - 1, replacementBytes.Length);
            for (int offset = replacementBytes.Length; offset < field.ByteLength; offset += prepared.PadBytes.Length)
            {
                Buffer.BlockCopy(prepared.PadBytes, 0, record, field.StartByte - 1 + offset, prepared.PadBytes.Length);
            }
        }

        private static void TransformCompositeField(byte[] record, PreparedField prepared, Encoding encoding)
        {
            FieldRule field = prepared.Rule;
            SegmentRule[] segments = field.Segments;
            StringBuilder logicalValue = new StringBuilder();
            for (int i = 0; i < segments.Length; i++)
            {
                SegmentRule segment = segments[i];
                string part = encoding.GetString(record, segment.StartByte - 1, segment.ByteLength);
                if (!ContainsDigitsOnly(part))
                {
                    throw new InvalidDataException("Record field " + field.Name + " must contain digits only");
                }
                logicalValue.Append(part);
            }
            if (logicalValue.Length != 10)
            {
                throw new InvalidDataException("Record field " + field.Name + " must contain exactly 10 digits");
            }
            string replacement = GetMappedValue(
                prepared.Mapping,
                logicalValue.ToString(),
                field.TrimSpaces,
                field.Name
            );
            if (replacement.Length != 10 || !ContainsDigitsOnly(replacement))
            {
                throw new InvalidDataException("Field " + field.Name + " replacement must contain exactly 10 digits");
            }

            byte[] replacementBytes = encoding.GetBytes(replacement);
            if (replacementBytes.Length != field.ByteLength)
            {
                throw new InvalidDataException("Field " + field.Name + " replacement has an invalid byte width");
            }
            int cursor = 0;
            for (int i = 0; i < segments.Length; i++)
            {
                SegmentRule segment = segments[i];
                Buffer.BlockCopy(replacementBytes, cursor, record, segment.StartByte - 1, segment.ByteLength);
                cursor += segment.ByteLength;
            }
        }

        private static bool ContainsDigitsOnly(string value)
        {
            if (String.IsNullOrEmpty(value))
            {
                return false;
            }
            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] < '0' || value[i] > '9')
                {
                    return false;
                }
            }
            return true;
        }

        private static string GetMappedValue(IDictionary mapping, string value, string trimMode, string fieldName)
        {
            string key = Normalize(value, trimMode, fieldName);
            if (!mapping.Contains(key))
            {
                throw new InvalidDataException("Mapping key not found for field " + fieldName);
            }
            object replacement = mapping[key];
            if (replacement == null)
            {
                throw new InvalidDataException("Mapping replacement is invalid for field " + fieldName);
            }
            return Convert.ToString(replacement, System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string Normalize(string value, string trimMode, string fieldName)
        {
            switch (trimMode)
            {
                case "None":
                    return value;
                case "Half":
                    return value.Trim(new char[] { ' ' });
                case "Full":
                    return value.Trim(new char[] { '\u3000' });
                case "HalfAndFull":
                    return value.Trim(new char[] { ' ', '\u3000' });
                default:
                    throw new InvalidDataException("Field " + fieldName + " has an unsupported trim mode");
            }
        }
    }
}
