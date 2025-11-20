# Error Model & Safety (P11)

**Status**: Error model ✅ complete, security features ⬜ not started

## Completed: Error Model ✅

Error handling is centralized in `com.tjclp.xl.error.XLError` (see `xl-core/src/com/tjclp/xl/error/XLError.scala`). Key cases:

- Addressing/validation: `InvalidCellRef`, `InvalidRange`, `InvalidReference`, `InvalidSheetName`, `InvalidColumn`, `InvalidRow`, `OutOfBounds`
- Workbook structure: `SheetNotFound`, `DuplicateSheet`, `InvalidWorkbook`, `ValueCountMismatch`
- Typing/formatting: `TypeMismatch`, `FormulaError`, `StyleError`, `NumberFormatError`, `MoneyFormatError`, `PercentFormatError`, `DateFormatError`, `AccountingFormatError`, `ColorError`, `UnsupportedType`
- IO/parse surface: `IOError`, `ParseError`, `Other`

- ✅ Public APIs are pure: return `XLResult[A] = Either[XLError, A]` (exceptions only via explicit `.unsafe` helpers)
- ✅ Platform exceptions are wrapped into `IOError` / `ParseError` in interpreter layers
- ✅ Validation: address parsing, sheet names, styles, codecs are total and never throw
- ✅ Codec errors: Type mismatches and decode failures stay in the `XLResult` channel

## Not Started: Security Features (P11) ⬜

The following security hardening features are deferred to P11:

### Threat Model
**Untrusted inputs**: Adversarial `.xlsx` files, malformed XML, formula injection vectors

**Attack vectors**:
- ZIP bombs (compression ratio attacks)
- XML entity expansion
- Formula injection (CSV injection via Excel)
- Path traversal in ZIP entries
- Resource exhaustion (memory, disk, CPU)

### ZIP Bomb Detection
- **Entry count limits**: Prevent archives with excessive file counts
- **Uncompressed size limits**: Detect compression ratio attacks (cap size/ratio)
- **Nested archive detection**: Prevent recursive bombs
- **Path traversal prevention**: Reject `..` or absolute paths in ZIP entries

### Formula Injection Guards
- **Untrusted text escaping**: Prefix with `'` when starting with `= + - @`
- **Configurable policy**: Strict (always escape) vs permissive (warn only)
- **CSV export safety**: Ensure formula injection protection on export

### XLSM Macro Handling
- **Opaque preservation**: Read/write XLSM parts without execution
- **Explicit stripping**: API to remove vbaProject.bin
- **Never execute**: No macro evaluation, ever

### File Size Limits
- **Max file size**: Configurable limit (default 100MB)
- **Max cell count**: Prevent pathological sheets (default 10M cells)
- **Max string length**: Prevent memory exhaustion (default 32KB per cell)

### XML External Entity (XXE) Prevention
- ✅ **Currently safe**: scala-xml and fs2-data-xml default to safe parsers
- ⬜ **Explicit hardening**: Disable DTDs, external entities explicitly
- **XML parser limits**: Entity expansion disabled, depth limits, attribute count limits

### Security Guidance
- **Never evaluate macros**: Do not embed active content (XLSM vbaProject.bin never executed)
- **Prefer typed formulas**: Use formula AST over string manipulation where possible
- **Formula text sanitizer**: Opt-in allowlist for untrusted text writes

## Priority

- **High**: ZIP bomb detection, formula injection guards (production use)
- **Medium**: File size limits, XLSM stripping
- **Low**: Explicit XXE hardening (already safe by default)
