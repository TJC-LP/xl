# Error Model & Safety (P11)

**Status**: Error model ✅ complete, security features ⬜ not started

## Completed: Error Model ✅

Total error handling implemented with `XLError` ADT:

```scala
enum XLError derives CanEqual:
  case Io(msg: String)                           // interpreters only
  case Parse(path: String, reason: String)
  case Semantic(reason: String)
  case Validation(reason: String)
  case InvalidCellRef(ref: String, reason: String)
  case CodecError(ref: ARef, reason: String)
  case Unsupported(feature: String)
```

- ✅ **No throws**: All public APIs return `Either[XLError, A]` (aliased as `XLResult[A]`)
- ✅ **Platform exception conversion**: IOException, XML parsing errors wrapped
- ✅ **Validation**: Cell ref parsing, range validation, style construction all total
- ✅ **Codec errors**: Type mismatches, parse failures returned as Left

## Not Started: Security Features (P11) ⬜

The following security hardening features are deferred to P11:

### ZIP Bomb Detection
- **Entry count limits**: Prevent archives with excessive file counts
- **Uncompressed size limits**: Detect compression ratio attacks
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

## Priority

- **High**: ZIP bomb detection, formula injection guards (production use)
- **Medium**: File size limits, XLSM stripping
- **Low**: Explicit XXE hardening (already safe by default)
