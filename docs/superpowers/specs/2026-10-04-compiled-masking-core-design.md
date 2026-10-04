# Compiled .NET Core for Fixed-Width and CSV Masking

**Status:** Approved by user
**Date:** 2026-10-04
**Parent design:** [Masked File Tool](./2026-10-04-masked-file-tool-design.md)

## Purpose

The PowerShell 5.1 implementation is correct on the current functional suite but performs poorly in a 1,000-record synthetic run. With a 300,000-entry mapping index, 1,000 records took 2.480 seconds for fixed-width and 3.683 seconds for CSV. A 300,000-record run remained in fixed-width processing after more than 17 minutes and was stopped. The implementation needs compiled record-processing loops while retaining the no-Python, local Windows execution model.

This change has no fixed time SLA. The 300,000-record synthetic benchmark will report input count, elapsed time, throughput, mapping-load time, and Excel status separately.

## Constraints and preserved behavior

- Windows 11 and Windows PowerShell 5.1; no Python, external service, NuGet package, or separately deployed executable.
- Runtime-compile an adjacent C# source file in memory with PowerShell `Add-Type`.
- Leave the CLI, JSON configuration, configuration validation, mapping-file loading, report schema, job orchestration, staged-output publication, and Excel COM handler in PowerShell.
- Preserve the source files, deterministic sampling, record order, line endings, quote state, CSV BOM behavior, NUL-padded CSV behavior, fixed-width FILLER bytes, split fixed-width fields, and fail-closed mapping behavior. Fixed-width text and CSV inputs are restricted to cp932 (code page 932); mapping-table encodings remain independently configurable.
- Never include input/replacement values in exceptions or reports.
- Do not silently fall back to the slower PowerShell implementation. If compilation is unavailable or fails, report an explicit preflight failure before processing.

## Architecture

### Compiled core loading

Add `src/Masking/MaskingCore.cs` and a small `CompiledCore.psm1` loader. The loader compiles the source once per PowerShell process and checks for the stable fully qualified type before compiling, so repeated `Import-Module -Force` does not create duplicate-type errors. `MaskingTool.ps1` loads the compiled core only for execution, after configuration validation; `-Help` and `-ValidateOnly` remain usable without compilation.

The handler modules remain the public PowerShell interfaces. They convert the validated PowerShell job objects into typed C# options and delegate the hot loop. The C# code does not read configuration files, choose arbitrary paths, publish final paths, or print logs.

### Fixed-width processing

The compiled method receives an input path, a runner-owned temporary output path, record width and terminator bytes, strict encoding, typed field rules, mapping indexes, and sampling options. It reads one record at a time into a reusable buffer, verifies record boundaries and terminators, transforms configured fields in place, preserves all other bytes, and writes selected records in source order. Composite `Segments` are joined, mapped once, and scattered back to their configured ranges.

`Sample` uses the configured count and seed with reservoir sampling; `All` streams every record. Output record width and configured output terminator behavior remain unchanged. Mapping misses, invalid encoded data, replacement overflow, malformed segments, and record-size errors abort the handler without exposing values.

### CSV processing

The compiled CSV path uses a streaming state machine, not `String.Split` or whole-file loading. It tracks field quote state, supports quoted delimiters, embedded CR/LF, and escaped double quotes, retains each row's terminator and configured BOM behavior, and rejects output that would require changing an originally unquoted field's quote state.

For `FixedWidthNullPadded`, it reads each configured byte block, accepts only trailing NUL padding, parses the payload as exactly one CSV record, then pads the serialized masked record back to the configured width. Fixed-width and CSV input encodings are validated as cp932 before processing; UTF-8, UTF-16, and other encodings are rejected.

CSV sampling makes a first streaming pass to select logical data-record indexes and a second pass to write the selected records in original order, retaining the header when configured. Selection uses the same seed/count semantics as the existing handler.

### PowerShell runner and reporting

The runner continues to create a unique same-volume temporary output, call the format handler, and publish only after successful completion. The compiled core only writes to the temporary path supplied by the handler. The existing report format, partial-failure behavior, exit codes, and path/value sanitization remain owned by PowerShell.

Excel remains on Office COM and is not routed through the compiled core. Its separate timing remains conditional on registered Excel COM; a missing prerequisite is reported as BLOCKED, never as a pass.

## Validation

- Keep existing module import, mapping, transform, fixed-width, CSV, runner, and CLI tests.
- Add fixed-width parity cases for cp932, preserved filler bytes, LF/CRLF/no terminator, split fields, deterministic samples, unknown mappings, and malformed widths.
- Add CSV parity cases for quote state, quoted delimiters/newlines/escaped quotes, row terminators, BOMs, fixed-width NUL padding, deterministic samples, column-count errors, and unknown mappings.
- Verify handler failures never publish output and that the runner continues other jobs with a value-free report.
- Run the default synthetic benchmark at 300,000 records for fixed-width and CSV; report timings and throughput without imposing an unapproved SLA. Measure a synthetic Excel workbook separately when COM is present.
- Run the complete suite under Windows PowerShell 5.1. Excel COM integration remains an explicit environment gate.

## Risks and trade-offs

- Runtime compilation may be disabled by endpoint policy or unavailable CodeDom components. In that case execution fails clearly; validation-only remains available.
- C# must reproduce PowerShell behavior exactly at byte and quote-state boundaries. Differential tests against existing outputs are required before replacing the existing loops.
- Runtime compilation adds startup cost once per execution process. The benchmark reports processing and mapping-load measurements separately so compilation/startup is not mistaken for per-record throughput.
- No fixed speed guarantee is made until the user chooses an SLA.
