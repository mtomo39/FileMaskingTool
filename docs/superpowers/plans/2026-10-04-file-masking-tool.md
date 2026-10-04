# Multi-Format File Masking Tool Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local PowerShell tool that masks configured fields in fixed-width text, CSV (including fixed-width NUL-padded records), `.xls`, and `.xlsx` files, producing either deterministic samples or full-file outputs.

**Architecture:** A PowerShell 5.1 CLI loads and validates a JSON job manifest, indexes tab-separated mapping files in ordinal dictionaries, and dispatches each job to a streaming format handler. Fixed-width and CSV handlers preserve byte/CSV structure; Excel uses Office 2024 COM on a staged copy. Every job writes to a temporary output, validates it, and publishes it only on success; logs never include data values.

**Tech Stack:** Windows PowerShell 5.1, .NET Framework APIs, Office 2024 Excel COM, JSON, self-contained PowerShell test harness. No Python or third-party runtime packages.

**Spec:** [../specs/2026-10-04-masked-file-tool-design.md](../specs/2026-10-04-masked-file-tool-design.md)

## Global Constraints

- Run on Windows 11 using Windows PowerShell 5.1 and installed Office 2024 for Excel automation.
- Do not install Python or add third-party product dependencies.
- Keep all input, mapping, temporary, and output files local to the controlled environment; do not add network sending.
- Never modify source files or log source values, replacement values, mapping values, or sample cell values.
- Fail the affected file on unknown mapping keys, invalid layouts, malformed records, or output-width overflow; never pass the original value through as a success fallback.
- Preserve original Excel extension and CSV field quote state; fixed-width CSV NUL padding is per-record trailing padding.

## Review Focus

- Multibyte encodings where byte offsets split a character or padding cannot exactly fill a field — test strict decode/encode failure and byte-width validation in Task 4.
- Split member IDs with noncontiguous ranges and leading zeroes — test logical concatenation and scattering in Task 4.
- CSV fields with quotes, escaped quotes, embedded delimiters/newlines, or replacements that require quoting — test quote-state preservation and fail-closed behavior in Task 5.
- Fixed-width CSV records with terminal NUL padding, record terminators, and transformations that grow beyond the record payload — test exact-width repadding and rejection in Task 5.
- Excel formulas, hidden sheets, comments/hyperlinks, legacy `.xls` saving, and COM cleanup — test configured sheets, value-only sampled output, warning status, extension retention, and cleanup in Task 6.

---

## File Structure

- Create: `MaskingTool.ps1` — CLI entry point, parameters, exit codes, and module imports.
- Create: `src\Masking\Config.psm1` — JSON loading and complete job/layout validation.
- Create: `src\Masking\Mapping.psm1` — mapping-table parsing and ordinal lookup dictionaries.
- Create: `src\Masking\Transform.psm1` — normalization, mapped-value lookup, byte-width padding, and field transformation.
- Create: `src\Masking\FixedWidth.psm1` — fixed-width byte-stream records and composite fixed-width fields.
- Create: `src\Masking\Csv.psm1` — CSV parser/serializer with quote-state preservation and optional NUL-padded fixed blocks.
- Create: `src\Masking\Excel.psm1` — `.xls`/`.xlsx` processing through Excel COM.
- Create: `src\Masking\Runner.psm1` — sampling, job orchestration, staging, finalization, and per-file failure handling.
- Create: `src\Masking\Report.psm1` — value-free job reports and timing/count summaries.
- Create: `config.example.json` — documented example job manifest covering each input format.
- Create: `tests\Run-Tests.ps1` — dependency-free assertion helpers and synthetic unit/integration tests.
- Create: `README.md` — local setup, configuration, operation, safety limitations, and test instructions.

### Task 1: Creating the PowerShell project skeleton and test harness

**Files:**
- Create: `MaskingTool.ps1`
- Create: `src\Masking\Config.psm1`
- Create: `src\Masking\Mapping.psm1`
- Create: `src\Masking\Transform.psm1`
- Create: `src\Masking\FixedWidth.psm1`
- Create: `src\Masking\Csv.psm1`
- Create: `src\Masking\Excel.psm1`
- Create: `src\Masking\Runner.psm1`
- Create: `src\Masking\Report.psm1`
- Create: `tests\Run-Tests.ps1`

**Interfaces:**
- Produces: A PowerShell 5.1 project layout with a test harness exposing `Assert-Equal`, `Assert-Throws`, and `Invoke-TestCase`; module files can be imported without running jobs.

- [ ] **Step 1: Write the failing harness smoke test**

Add tests asserting `MaskingTool.ps1 -Help` exits successfully and each module can be imported in Windows PowerShell 5.1 without side effects.

- [ ] **Step 2: Run the test to verify it fails**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: FAIL because the CLI, modules, and harness do not yet exist.

- [ ] **Step 3: Create the CLI/module skeleton and assertion helpers**

Make the CLI import modules relative to `$PSScriptRoot`; add only help handling at this stage. Keep module imports side-effect free.

- [ ] **Step 4: Run the test to verify it passes**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: PASS for help and clean imports.

### Task 2: Loading and validating the JSON job manifest

**Files:**
- Modify: `src\Masking\Config.psm1`
- Modify: `MaskingTool.ps1`
- Create: `config.example.json`
- Modify: `tests\Run-Tests.ps1`

**Interfaces:**
- Produces: `Import-MaskingConfig([string] $Path) -> PSCustomObject`
- Produces: `Test-MaskingConfig([PSCustomObject] $Config) -> void`, throwing a sanitized validation exception on invalid configuration.
- CLI supports `MaskingTool.ps1 -Config <path> -ValidateOnly`; this validates paths and configuration without creating outputs.
- Config job properties: `InputPath`, `OutputPath`, `Format`, `Mode`, `SampleCount`, `Seed`, `Encoding`, plus format-specific `Layout` and `Fields`.

- [ ] **Step 1: Write failing configuration tests**

Cover a valid job; validation-only mode creating no output; missing input path; source/output path collision; unknown format/mode; invalid sample count; overlapping fixed-width fields; a field beyond `recordWidthBytes`; and missing configuration for any workbook sheet.

- [x] **Step 2: Run the tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: FAIL because configuration functions are not implemented.

- [ ] **Step 3: Implement JSON import and preflight validation**

Support `FixedWidth`, `Csv`, and `Excel` formats; `Sample` and `All` modes; 1-based `startByte`; explicit field widths, encodings, record terminators, CSV dialect, Excel sheets, and mapping references. Add `-ValidateOnly` to run preflight without output creation. Do not include or print actual data values in errors.

- [ ] **Step 4: Run the configuration tests**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: PASS for accepted and rejected configurations.

### Task 3: Indexing mapping tables and applying safe field transformations

**Files:**
- Modify: `src\Masking\Mapping.psm1`
- Modify: `src\Masking\Transform.psm1`
- Modify: `tests\Run-Tests.ps1`

**Interfaces:**
- Produces: `Read-MappingTable([string] $Path, [System.Text.Encoding] $Encoding, [string] $KeyTrimMode) -> System.Collections.Generic.Dictionary[string,string]`
- Produces: `Get-MappedValue($Index, [string] $InputValue, [string] $KeyTrimMode) -> string`, throwing on missing keys.
- Produces: `Convert-FieldValue([string] $Value, $FieldRule, $MappingIndex, [System.Text.Encoding] $Encoding) -> byte[]`
- Mapping indexes use `StringComparer.Ordinal`; mapping rows contain exactly two tab-separated columns.

- [ ] **Step 1: Write failing mapping/transform tests**

Cover 300,000 synthetic mapping rows, duplicate identical and conflicting keys, two-column validation, half-/full-width space trimming on both lookup values and keys, missing-key failure, half-/full-width padding, exact-fit output, overflow, and non-divisible padding remainder.

- [x] **Step 2: Run the tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: FAIL because mapping and transformation functions are not implemented.

- [ ] **Step 3: Implement mapping indexes and byte-safe transforms**

Load each table once into an ordinal dictionary. Encode with strict encoder/decoder fallbacks. Never truncate transformed values; fail when the encoded value exceeds the field width or the chosen pad character cannot fill the remaining bytes exactly.

- [ ] **Step 4: Run mapping/transform tests**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: PASS, including the 300,000-key synthetic lookup test.

### Task 4: Processing fixed-width files and split logical identifiers

**Files:**
- Modify: `src\Masking\FixedWidth.psm1`
- Modify: `src\Masking\Transform.psm1`
- Modify: `tests\Run-Tests.ps1`

**Interfaces:**
- Produces: `Invoke-FixedWidthJob($Job, $MappingIndexes, [int[]] $SelectedRecordIndexes, [string] $TemporaryOutputPath) -> PSCustomObject`
- Fixed-width field rules use 1-based `startByte`, `byteLength`, `mapping`, `trimSpaces`, and `padSpace`.
- Composite rules use ordered `segments`, each with `startByte` and `byteLength`; the concatenated logical value is transformed once and scattered back by segment widths.

- [ ] **Step 1: Write failing fixed-width tests**

Cover a 3,000-byte no-newline file, a 150-byte record with LF and CRLF terminators, malformed sizes/terminators, FILLER byte preservation, field overlap, deterministic samples, and a split 10-digit member number made from prefix 3 + suffix 7 with leading zeroes.

- [ ] **Step 2: Run the tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: FAIL because the fixed-width job handler is not implemented.

- [ ] **Step 3: Implement binary fixed-width processing**

Use `FileStream` and exact byte offsets; do not use `ReadLine()` or decode the whole file. For line-terminated fixed-width data, read the configured record width and then validate the exact terminator bytes. For split member numbers, concatenate segments in configured logical order, map once, validate 10 digits, and scatter 3 + 7 digits to their original ranges.

- [ ] **Step 4: Run fixed-width tests**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: PASS with every output record preserving configured byte width and terminator behavior.

### Task 5: Parsing and writing CSV with quote-state and NUL-padding preservation

**Files:**
- Modify: `src\Masking\Csv.psm1`
- Modify: `src\Masking\Config.psm1`
- Modify: `tests\Run-Tests.ps1`

**Interfaces:**
- Produces: `Read-CsvRecords([System.IO.TextReader] $Reader, [char] $Delimiter, [char] $Quote)` emits one `PSCustomObject` logical record at a time; each field has `Value` and `WasQuoted`.
- Produces: `Write-CsvRecord($Record, [char] $Delimiter, [char] $Quote) -> string`; throws if an originally unquoted field must be quoted to remain valid.
- Produces: `Invoke-CsvJob($Job, $MappingIndexes, [string] $TemporaryOutputPath) -> PSCustomObject`
- CSV fixed-width mode uses `recordWidthBytes`, `recordTerminator`, and `nullPadByte=0`; `recordWidthBytes` excludes the terminator.

- [x] **Step 1: Write failing CSV tests**

Cover quoted/unquoted fields, commas and line breaks inside quotes, doubled quotes, header preservation, exact quote-state preservation, error when an unquoted transformed value needs quoting, fixed-width NUL-padded blocks, CRLF/LF/no terminator, and output overflow.

- [x] **Step 2: Run the tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: FAIL because the CSV parser and serializer are not implemented.

- [x] **Step 3: Implement a quote-aware CSV state machine and record writer**

Retain each field's `WasQuoted` state while parsing and stream logical records rather than loading the entire CSV. In NUL-padded mode, read one exact-width byte block, validate the terminator, strip only terminal `0x00` padding, decode and parse the CSV payload, then encode and re-pad the serialized payload to the same width. Reject NUL inside the payload and never silently add quotes.

- [x] **Step 4: Run CSV tests**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: PASS with quote state, column count, terminator, and fixed record width preserved.

### Task 6: Processing `.xls` and `.xlsx` workbooks through Excel COM

**Files:**
- Modify: `src\Masking\Excel.psm1`
- Modify: `src\Masking\Config.psm1`
- Modify: `tests\Run-Tests.ps1`

**Interfaces:**
- Produces: `Invoke-ExcelJob($Job, $MappingIndexes, [string] $TemporaryOutputPath) -> PSCustomObject`
- Job sheet rules identify every worksheet and designate its data range, sample/all behavior, target columns, and explicitly non-personal ranges.
- The result reports per-sheet row counts and detected auxiliary-information warnings without cell values.

- [x] **Step 1: Write failing Excel integration tests**

Generate synthetic `.xlsx` and `.xls` workbooks with multiple sheets, headers, formulas, formatting, hidden sheets, comments, hyperlinks, and target fields. Assert sample rows, formula-to-value conversion, masked cells, preserved sheet names/extension, warning reporting, and no source modification.

- [x] **Step 2: Run the tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: FAIL because Excel processing is not implemented.

- [x] **Step 3: Implement guarded Excel COM processing**

Create a separate Excel COM instance, disable macros/events, calculate before sampling, process only a staged copy, convert formulas in retained output ranges to values, preserve configured basic formatting, detect/report supported auxiliary content, save with the input workbook format, and close/release every COM object in `finally`.

- [ ] **Step 4: Run Excel integration tests** — BLOCKED in this host because Office Excel COM is not registered; the test reports BLOCKED and is not counted as passed.

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: PASS with Office 2024 installed; otherwise tests report an explicit environment prerequisite failure, not a pass.

### Task 7: Orchestrating jobs, deterministic sampling, staging, and reports

**Files:**
- Modify: `src\Masking\Runner.psm1`
- Modify: `src\Masking\Report.psm1`
- Modify: `MaskingTool.ps1`
- Modify: `tests\Run-Tests.ps1`

**Interfaces:**
- Produces: `Invoke-MaskingRun([PSCustomObject] $Config) -> int` where `0` means all jobs completed and validated, `1` means one or more jobs failed, and `2` means configuration/preflight failed.
- Each format handler writes only to the provided temporary path and returns counts, elapsed time, and warnings without values.
- Runner publishes the temporary output only after handler validation; failures remove only that job's temporary output.

- [x] **Step 1: Write failing runner tests**

Cover multi-job configurations, per-file sample/all modes, deterministic sample selection from seed, a failed file followed by a successful file, no source overwrite, temporary output cleanup on failure, final output publication only after validation, nonzero aggregate exit status after a partial failure, and sanitized logs/reports.

- [ ] **Step 2: Run the tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: FAIL because runner and reporting functions are not implemented.

- [x] **Step 3: Implement orchestration and value-free reporting**

Preflight all configured jobs before processing. Load referenced maps once, dispatch each job, continue to subsequent jobs after a per-file failure, stage output under a unique same-volume temporary path, verify and publish each file independently, and record counts/timing/errors without any source, replacement, or cell values. Excel auxiliary warnings must be reported as warnings and must not be represented as proof of privacy.

- [x] **Step 4: Run runner tests**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: PASS with failed jobs returning nonzero exit status and no incomplete final output.

### Task 8: Completing the example configuration, operating guide, and performance verification

**Files:**
- Modify: `config.example.json`
- Create: `README.md`
- Modify: `tests\Run-Tests.ps1`

**Interfaces:**
- Produces: An example manifest for each supported input format and documented CLI invocation, exit codes, report fields, and Excel prerequisites.
- Produces: A synthetic performance test that reports fixed-width/CSV throughput for 300,000 records and Excel timing separately.

- [x] **Step 1: Write failing documentation/config smoke tests**

Validate the example JSON against `Test-MaskingConfig`; assert help and README document no-newline fixed-width, 150-byte terminated records, quote-state behavior, NUL-padded CSV, split fixed-width member IDs, sample/all modes, and Excel `.xls`/`.xlsx` warnings.

- [ ] **Step 2: Run the smoke tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: FAIL until the example and operating guide cover each required format and rule.

- [x] **Step 3: Complete examples, safety guidance, and synthetic benchmark command**

Document that the example contains synthetic paths/data only; warn that FILLER and unconfigured cells remain unchanged and that Excel auxiliary warnings do not guarantee absence of personal data.

- [x] **Step 4: Run the full test suite and synthetic benchmark** — After moving the record loops into the in-memory compiled core, the full suite passes (Excel COM integration is BLOCKED in this environment) and the 300,000-record synthetic benchmark completes: fixed-width 0.379 s and CSV 0.303 s. These are measurements, not an SLA.

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: PASS for all non-Excel tests; Excel integration passes with Office 2024 available. The benchmark prints record counts and elapsed time without reading or emitting production data.

## Final Verification

- Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
- Run the example manifest in validation-only mode and confirm it makes no output files.
- Run one synthetic job per format and compare source hashes before/after to confirm inputs remain unchanged.
- Inspect generated reports to confirm they contain no data values.
- Measure fixed-width and CSV throughput at 300,000 synthetic records and record Excel timing separately; do not claim a performance SLA until the user selects one.
