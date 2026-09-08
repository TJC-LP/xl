<!-- GENERATED from `xl schema --json` by DocsGenSpec; do not edit by hand.
     A diff here is a contract change. Regenerate: XL_UPDATE_DOCS=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.DocsGenSpec -->

# Error and warning codes

The complete vocabulary a `code:` line (text mode) or `error.code` (`--json`) can carry:
the CLI's own codes first, then every domain (`XLError`) code. The exit code follows from
the error code alone. Warnings are `Warning[<CODE>]: <message>` lines on stderr — or
`warnings[]` in the envelope — and never change the exit code.

## Error codes

| code | exit |
| --- | --- |
| `USAGE` | 2 |
| `UNKNOWN_VERB` | 2 |
| `OUTPUT_REQUIRED` | 2 |
| `UNSUPPORTED_IN_STREAM` | 2 |
| `BATCH_JSON_INVALID` | 2 |
| `BATCH_OP_UNKNOWN` | 2 |
| `BATCH_OP_INVALID` | 2 |
| `BATCH_OP_FAILED` | 3 |
| `RASTERIZER_UNAVAILABLE` | 3 |
| `IO_READ` | 3 |
| `IO_WRITE` | 3 |
| `RESOURCE_LIMIT` | 3 |
| `RECALC_GATE` | 1 |
| `DIFFERENCES_FOUND` | 1 |
| `LINT_FINDINGS` | 1 |
| `AUDIT_FINDINGS` | 1 |
| `INTERNAL` | 3 |
| `INVALID_CELL_REF` | 3 |
| `INVALID_RANGE` | 3 |
| `INVALID_REFERENCE` | 3 |
| `INVALID_SHEET_NAME` | 3 |
| `OUT_OF_BOUNDS` | 3 |
| `SHEET_NOT_FOUND` | 3 |
| `DUPLICATE_SHEET` | 3 |
| `DUPLICATE_CELL_REF` | 3 |
| `INVALID_COLUMN` | 3 |
| `INVALID_ROW` | 3 |
| `TYPE_MISMATCH` | 3 |
| `FORMULA_ERROR` | 3 |
| `STYLE_ERROR` | 3 |
| `NUMBER_FORMAT_ERROR` | 3 |
| `MONEY_FORMAT_ERROR` | 3 |
| `PERCENT_FORMAT_ERROR` | 3 |
| `DATE_FORMAT_ERROR` | 3 |
| `ACCOUNTING_FORMAT_ERROR` | 3 |
| `COLOR_ERROR` | 3 |
| `INVALID_WORKBOOK` | 3 |
| `IO_ERROR` | 3 |
| `PARSE_ERROR` | 3 |
| `SECURITY_ERROR` | 3 |
| `VALUE_COUNT_MISMATCH` | 3 |
| `UNSUPPORTED_TYPE` | 3 |
| `INVALID_TABLE_NAME` | 3 |
| `INVALID_TABLE_DISPLAY_NAME` | 3 |
| `INVALID_TABLE_RANGE` | 3 |
| `INVALID_TABLE_COLUMNS` | 3 |
| `OTHER` | 3 |
| `EDIT_FAILED` | 3 |
| `UNSUPPORTED_CAPABILITY` | 3 |
| `SHEET_REQUIRED` | 3 |

## Warning codes

| code |
| --- |
| `READER_WARNING` |
| `TRUNCATED` |
| `HIDDEN_OMITTED` |
| `UNKNOWN_PROPERTY` |
| `FORMAT_HINT_IGNORED` |
| `STREAM_BACKEND_ONLY` |
| `RECALC_ERRORS` |
| `SHEET_AUTOSELECTED` |
| `FLAG_IGNORED` |
| `EVAL_FAILED` |
| `MEMORY_PRESSURE` |
| `OFF_GRID_REF` |
