# Fixture corpus provenance (GH-240)

Real-file test inputs produced by **foreign writers** (openpyxl, LibreOffice), so reader
tests against them cannot be self-fulfilling round-trips of XL's own XML dialect.

**All content is synthetic/invented for testing. No confidential data.**

## Regeneration

```bash
python3 scripts/generate-fixtures.py            # everything (needs soffice for -lo files)
python3 scripts/generate-fixtures.py --skip-lo  # openpyxl fixtures only
```

- Generated on macOS with **Python 3.14.3**, **openpyxl 3.1.5**, **LibreOffice 25.8.2.2**.
- openpyxl fixtures are **byte-stable** across regeneration: cell dates are fixed
  literals, docProps timestamps and zip entry timestamps are pinned by the script
  (including the `dcterms:modified` stamp openpyxl rewrites at save time).
- `*-lo.xlsx` files are **not byte-stable** across regeneration (LibreOffice embeds
  generation metadata in docProps); the committed files are canonical. Only regenerate
  them deliberately.
- Budget: total corpus must stay **< 1 MB** (script enforces; currently ~60 KB).

## Files

| File | Producer | Contents |
|------|----------|----------|
| `small-values.xlsx` | openpyxl | Sheet `Values`: strings (ASCII, accented latin, CJK, astral-plane emoji, leading/trailing spaces, internal double spaces — **inline-string dialect**, `xml:space="preserve"`), numbers (int, negative, decimal, 13-digit, `1.23e-10` lowercase exponent, `9.99e15`), booleans, date serials (45366, 45366.6046875), digit string `"0123"` |
| `styled.xlsx` | openpyxl | Sheet `Styled`: Arial 14 bold red on yellow fill, italic + thin box border, `$#,##0.00` and `0.00%` numFmts, `yyyy-mm-dd` date format, wrap + center + indent 2, thick bottom border, merge `A4:C4`; every cell carries `s=1..8` |
| `formulas.xlsx` | openpyxl | Sheets `Data` (numbers) + `Calc`: `SUM(Data!A1:A5)`, cross-sheet `Data!A1*2`, leading-plus `+Data!A2`, absolute `$A$1+A2`, mixed `SUM($A$1:A4)`; cached values are **empty** (`<f>…</f><v />`) |
| `autofilter.xlsx` | openpyxl | Sheet `Filtered`: 3×6 table with `autoFilter ref="A1:C6"` |
| `chart-bar.xlsx` | openpyxl | Sheet `ChartData`: quarter/units table + BarChart anchored at D2 → `xl/charts/chart1.xml`, `xl/drawings/drawing1.xml` (+rels). NOTE: openpyxl binds `xmlns:r` on the `<drawing>` element itself, not the worksheet root |
| `image.xlsx` | openpyxl | Sheet `HasImage`: 73-byte hand-assembled 3×3 PNG (base64 literal in the script, no Pillow) anchored at B2 → `xl/media/image1.png`, `xl/drawings/drawing1.xml` |
| `comments-hyperlinks.xlsx` | openpyxl | Sheet `Notes`: comment on A1 (author `Fixture Bot`) — stored in openpyxl's `xl/comments/comment1.xml` subdirectory dialect + VML; external hyperlink `https://example.com/xl-fixtures` on B1; internal `#Notes!A1` link on B2 |
| `image-shape.xlsx` | derived (zip surgery on image.xlsx) | Same as image.xlsx plus a rel-free `<sp>` shape in a `twoCellAnchor editAs="oneCell"` appended to the SAME `wsDr` — mixed typed-Picture + Preserved-fragment coverage for the GH-221 drawing layer. Derivation: replace `</oneCellAnchor></wsDr>` in `xl/drawings/drawing1.xml` with the shape anchor, zip timestamps pinned to 1980-01-01 |
| `small-values-lo.xlsx` | LibreOffice (from small-values.xlsx) | **SST dialect** (everything `xml:space="preserve"`), explicit `t="n"`/`s=` on every cell, booleans rewritten as cached `TRUE()`/`FALSE()` formulas, `1.23E-010` three-digit exponent, adds `docProps/custom.xml` |
| `styled-lo.xlsx` | LibreOffice (from styled.xlsx) | LibreOffice re-encoding of all styles (own styles.xml layout) |
| `formulas-lo.xlsx` | LibreOffice (from formulas.xlsx) | Same formulas with **computed cached values** (`<f aca="false">…</f><v>150</v>`) |

## Consumers

- `xl-ooxml/test/.../FixtureCorpusSpec.scala` — dialect parsing + value spot-checks + write→read stability
- `xl-ooxml/test/.../FixturePreservationSpec.scala` — chart/drawing/media preservation ground truth (GH-221/GH-222)
- `xl-cats-effect/test/.../StreamingParitySpec.scala` — streaming vs in-memory reader parity law
  (xl-cats-effect test resources include this directory via `xl-cats-effect/package.mill`)
