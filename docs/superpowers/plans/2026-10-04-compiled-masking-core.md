# Compiled Masking Core Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the slow PowerShell per-record loops for fixed-width and CSV jobs with an in-memory compiled .NET core without changing output or failure behavior.

**Architecture:** Keep configuration, mapping-table loading, CLI, orchestration, reporting, staging, and Excel COM in PowerShell. Add a C# source file loaded once per process with `Add-Type`; retain existing handler signatures as adapters that translate validated job objects and dispatch to compiled streaming processors.

**Tech Stack:** Windows PowerShell 5.1, .NET Framework, C# compiled in memory with `Add-Type`; no Python, NuGet, or separate executable.

**Spec:** [2026-10-04-compiled-masking-core-design.md](../specs/2026-10-04-compiled-masking-core-design.md)

## Global Constraints

- Windows 11 and Windows PowerShell 5.1; no Python, external service, NuGet package, or separately deployed executable.
- Runtime-compile an adjacent C# source file in memory with PowerShell `Add-Type`.
- Use C# syntax supported by the .NET Framework compiler available to Windows PowerShell 5.1; do not depend on modern C# language features.
- Leave the CLI, JSON configuration, configuration validation, mapping-file loading, report schema, job orchestration, staged-output publication, and Excel COM handler in PowerShell.
- Preserve the source files, deterministic sampling, record order, line endings, quote state, CSV BOM behavior, NUL-padded CSV behavior, fixed-width FILLER bytes, split fixed-width fields, and fail-closed mapping behavior. Fixed-width and CSV input files are cp932-only; mapping-table encoding remains separately configurable.
- Never include input/replacement values in exceptions or reports.
- Do not silently fall back to the slower PowerShell implementation. If compilation is unavailable or fails, report an explicit preflight failure before processing.
- This change has no fixed time SLA. Report synthetic 300,000-record throughput; do not claim a speed guarantee.

## Review Focus

- cp932 multibyte fields and malformed encoded bytes: fixed-width task must preserve boundaries and fail without a published output.
- Composite fields with non-contiguous segments and leading zeroes: fixed-width tests must map once and scatter by configured logical order.
- CSV quoted CRLF/newlines and escaped quotes: CSV tests must retain logical record and quote state.
- CSV transformed text that needs quoting when source was unquoted: CSV tests must fail rather than alter quote state.
- Fixed-width NUL-padded CSV with interior NUL or UTF-16/UTF-32 configuration: CSV tests must reject ambiguous or malformed payloads.

---

### Task 1: Compiled core loader and fixed-width processing

**Files:**
- Create: `src/Masking/MaskingCore.cs`
- Create: `src/Masking/CompiledCore.psm1`
- Modify: `src/Masking/FixedWidth.psm1`
- Modify: `src/Masking/Runner.psm1`
- Modify: `tests/Run-Tests.ps1`

**Interfaces:**
- Consumes: existing `Invoke-FixedWidthJob([object] $Job, [object] $MappingIndexes, [int[]] $SelectedRecordIndexes, [string] $TemporaryOutputPath)` and mapping indexes produced by `Read-MappingTable`.
- Produces: `Import-MaskingCore()` that compiles the stable C# type once per process; `MaskingCore.FixedWidthOptions` with `InputPath`, `OutputPath`, `RecordWidthBytes`, input/output terminator bytes, strict encoding, `Fields`, mode, sample count, and seed; `MaskingCore.FieldRule` with `Name`, `StartByte`, `ByteLength`, `MappingName`, `TrimSpaces`, `PadSpace`, and optional `Segments`; `MaskingCore.SegmentRule` with `StartByte` and `ByteLength`; `MaskingCore.FixedWidthResult`; and `MaskingCore.ProcessFixedWidth(FixedWidthOptions, IDictionary) -> FixedWidthResult`, whose counts/timing expose `RecordsRead`, `RecordsWritten`, `RecordsInSource`, `FieldsTransformed`, and `ElapsedMilliseconds`.
- Preserves: the existing PowerShell handler signature and return properties consumed by `Runner.psm1`.

- [x] **Step 1: Write failing loader and fixed-width parity tests**

Add these `Invoke-TestCase` cases:

- `compiled core import is idempotent`: import twice and assert one stable `MaskingCore` type is usable.
- `compiled fixed-width core preserves bytes and terminators`: assert exact output for a 3,000-byte record, unchanged filler, cp932 field padding, and None/LF/CRLF terminators; assert malformed cp932 input fails.
- `compiled fixed-width core maps split member number once`: assert 3+7 segments are combined, mapped with leading zeroes intact, and scattered to the original offsets.
- `compiled fixed-width sampling is deterministic`: same seed selects the same source-ordered records.
- `compiled fixed-width mapping failure does not publish output`: an unknown mapping fails and leaves no completed output.

- [x] **Step 2: Run tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: FAIL because the compiled core loader/type and fixed-width dispatch are not implemented.

- [x] **Step 3: Implement the compiled core loader and fixed-width adapter**

Compile `MaskingCore.cs` from `CompiledCore.psm1` only when its fully qualified type is absent. Import the loader before the format handler in test/runtime execution; CLI integration is completed in Task 3. Adapt validated fixed-width fields once per job; process records in a reusable byte buffer, validate input/output terminators and widths, apply strict encoding, map composite fields once, and write only to the runner-provided staged path. Update `Runner.psm1` to let the compiled handler perform seeded sampling instead of selecting every fixed-width sample index in PowerShell.

- [x] **Step 4: Run fixed-width and loader tests**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: all fixed-width and loader tests pass; reports and Runner behavior remain unchanged; Excel integration remains explicitly BLOCKED without COM.

- [ ] **Step 5: Commit the fixed-width compiled core** (pending Git author identity; do not change Git configuration without user approval)

```powershell
git add src/Masking/MaskingCore.cs src/Masking/CompiledCore.psm1 src/Masking/FixedWidth.psm1 src/Masking/Runner.psm1 tests/Run-Tests.ps1
git commit -m "perf: compile fixed-width masking core"
```

### Task 2: Compiled CSV processing

**Files:**
- Modify: `src/Masking/MaskingCore.cs`
- Modify: `src/Masking/Csv.psm1`
- Modify: `tests/Run-Tests.ps1`

**Interfaces:**
- Consumes: Task 1 `Import-MaskingCore()` and the outer mapping index dictionary loaded by `Runner.psm1`.
- Produces: `MaskingCore.CsvOptions` with input/output paths, delimiter, quote, header flag, container mode, record width, NUL pad byte, terminator, strict encoding, BOM flag, fields, mode, sample count, and seed; `MaskingCore.CsvFieldRule` with `Name`, column index/name, `MappingName`, and `TrimSpaces`; `MaskingCore.CsvResult`; and `MaskingCore.ProcessCsv(CsvOptions, IDictionary) -> CsvResult`, whose counts expose `RecordsRead`, `RecordsWritten`, and `DataRecordsWritten`; compiled record-based and fixed-width NUL-padded CSV paths behind the existing `Invoke-CsvJob([object] $Job, [object] $MappingIndexes, [string] $OutputPath)` interface; result properties remain compatible with `Runner.psm1`.

- [x] **Step 1: Write failing CSV parity tests**

Add these `Invoke-TestCase` cases, calling the compiled C# path directly as well as the PowerShell handler adapter:

- `compiled CSV preserves quote state and embedded records`: assert quoted/unquoted states, delimiters and CRLF inside quotes, escaped `""`, and exact LF/CRLF terminators.
- `compiled CSV preserves BOM and rejects encoding mismatch`: assert unchanged BOM presence and fail on a conflicting BOM.
- `compiled CSV sample is deterministic and keeps header`: same seed selects identical rows in source order.
- `compiled fixed-width CSV repads the same byte width`: assert trailing NUL padding and record terminator bytes.
- `compiled fixed-width CSV rejects malformed payloads`: reject interior NUL, overflow, and UTF-16/UTF-32 configuration.
- `compiled CSV fails closed on invalid row or mapping`: reject inconsistent column count, unknown mapping, and a transformed unquoted field that would require quoting.

- [x] **Step 2: Run tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: the new compiled-path cases fail because `Invoke-CsvJob` still uses the PowerShell per-character loop.

- [x] **Step 3: Implement compiled streaming CSV parsing and writing**

Implement the state machine in C# without `String.Split` or whole-file buffering. Preserve `WasQuoted`, exact logical record boundaries and terminators, strict encoding and BOM behavior. For NUL-padded mode, parse exactly one payload record, allow only trailing zero padding, and restore the configured byte width. Implement seeded two-pass CSV sampling while retaining headers.

- [x] **Step 4: Run CSV parity tests and the full suite**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: all CSV, fixed-width, runner, and CLI tests pass; no source values appear in errors or reports.

- [ ] **Step 5: Commit compiled CSV processing** (pending Git author identity; do not change Git configuration without user approval)

```powershell
git add src/Masking/MaskingCore.cs src/Masking/Csv.psm1 tests/Run-Tests.ps1
git commit -m "perf: compile CSV masking core"
```

### Task 3: CLI integration, measured performance, and operating guidance

**Files:**
- Modify: `MaskingTool.ps1`
- Modify: `tests/Measure-Performance.ps1`
- Modify: `tests/Run-Tests.ps1`
- Modify: `README.md`
- Modify: `docs/superpowers/plans/2026-10-04-file-masking-tool.md`

**Interfaces:**
- Consumes: Task 1 and Task 2 compiled handlers, existing runner/report contract, and synthetic benchmark command.
- Produces: end-to-end CLI execution that surfaces compilation errors as preflight failures; a benchmark reporting compilation, mapping load, fixed-width, CSV, and separate Excel timing; operating guidance for the runtime compiler prerequisite.

- [x] **Step 1: Write failing CLI and benchmark tests**

Add these `Invoke-TestCase` cases:

- `CLI help and validation-only bypass core compilation`: assert both return zero without loading the compiled type or creating outputs.
- `CLI reports compiled-core load failure as preflight exit 2`: in an isolated temporary project copy with invalid C# source, assert explicit failure and no output publication.
- `example manifest validates through CLI without output`: substitute temporary fixture paths, invoke `-ValidateOnly`, and assert every configured output remains absent.
- `synthetic benchmark reports each processing stage`: invoke a small record count and assert compilation, mapping, fixed-width, CSV, and Excel status are reported separately.

- [x] **Step 2: Run tests to verify they fail**

Run: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1`
Expected: CLI compilation-failure handling and separate compiled benchmark measurements are absent or fail before implementation.

- [x] **Step 3: Complete CLI, benchmark, and documentation wiring**

Measure C# startup compilation separately from mapping loading and per-format processing. Map compilation failures to the existing preflight exit code `2` before job publication. Run 300,000 synthetic records for both fixed-width and CSV; do not read or emit production content. Keep Excel timing separate and explicitly BLOCKED if Excel COM is not registered. Document that no time SLA is promised.

- [x] **Step 4: Run final verification**

Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1
.\tests\Measure-Performance.ps1
git diff --check
```

Expected: all non-Excel tests pass; Excel integration is BLOCKED only when COM is unavailable; the 300,000-record benchmark completes and prints fixed-width/CSV rates and a separate Excel status; validation-only creates no output; no whitespace errors.

- [ ] **Step 5: Commit end-to-end integration** (pending Git author identity; do not change Git configuration without user approval)

```powershell
git add MaskingTool.ps1 tests/Measure-Performance.ps1 tests/Run-Tests.ps1 README.md docs/superpowers/plans/2026-10-04-file-masking-tool.md
git commit -m "perf: wire compiled masking benchmark"
```

## Final Verification

- Run the full PowerShell 5.1 test suite and read every PASS/FAIL/BLOCKED result.
- Run the default 300,000-record synthetic benchmark; report measured throughput only, with no SLA claim.
- Confirm normal execution still stages before publish, failures expose no output values, and input hashes remain unchanged.
- Confirm `-Help` and `-ValidateOnly` work without triggering runtime compilation.
- Keep Excel COM integration marked BLOCKED unless run with registered Office Excel COM.
