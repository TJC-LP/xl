
# OOXML Mapping — Parts, Relationships, and Deterministic Printing

## Main parts
- `[Content_Types].xml` — content type registry (defaults + overrides).
- `/_rels/.rels` — package relationships (to workbook).
- `/xl/workbook.xml` and `/xl/_rels/workbook.xml.rels` — sheets, shared strings, styles, theme, drawings, etc.
- `/xl/worksheets/sheet#.xml` — sheet data (`sheetData`), merges, hyperlinks.
- `/xl/sharedStrings.xml` — `sst` with `si` items (dedup).
- `/xl/styles.xml` — number formats, fonts, fills, borders, `cellXfs`.
- `/xl/theme/theme1.xml` — palette.
- `/xl/drawings/drawing#.xml` and `/xl/charts/chart#.xml` — drawings & charts.
- `/xl/tables/table#.xml` — tables.

## Deterministic printing
- **Canonical order** for lists (styles, fonts, fills, borders, xfs, merges).
- **Stable IDs** derived from structural hashes (styles) with collision‑safe tie‑breakers.
- **Whitespace & attribute order** normalized for golden tests.

## Shared strings
- SST index bi‑map with LRU cache; optional spill to a temp KV store for very large files.
- All string values are normalized to NFC and deduplicated by exact codepoints.

## Dates
- 1900/1904 epoch handled **only at IO boundaries**. Internal `DateTime` keeps precise `LocalDateTime`.
