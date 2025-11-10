
# Styles & Themes — Canonical, Deduplicated, Mapped to OOXML

## Units & Colors
- Units: `Pt` (Double), `Px` (Double), `Emu` (Long). Pure conversions with Excel DPI assumptions.
- `Color`: `Rgb(argb: Int)` or `Theme(slot: ThemeSlot, tint: Double)`; tint ∈ [‑1, 1].

## Number Formats — grammar
```
General | 0 | 0.00 | #,##0 | #,##0.00 | 0% | 0.00% | "$"#,##0.00 | [Red]0.00 | @
```
Compile‑time `nf"..."` literal tokenizes into `NumFmt` ADT; unknown tokens fall back to `Custom` with validation.

## CellStyle & Canonicalization
- `CellStyle(font, fill, border, numFmt, align)`
- Canonicalizer generates a structural key; equivalent styles deduplicate to same index in `styles.xml`.

## Patches (monoid)
```scala
enum StylePatch:
  case SetFont(f: Font)
  case SetFill(f: Fill)
  case SetBorder(b: Border)
  case SetNumFmt(n: NumFmt)
  case SetAlign(a: Align)
  case Batch(ps: Vector[StylePatch])
```
- **Associative**; identity is `Batch(Vector.empty)`.
- **Idempotent** for repeated identical setters.

## OOXML mapping (key excerpts)
| Concept | XML Part/Node | Notes |
|---|---|---|
| Number formats | `styles.xml / numFmts / numFmt@formatCode` | Stable IDs assigned deterministically |
| Fonts/Fills/Borders | `styles.xml / fonts|fills|borders` | Dedup by canonical form |
| CellXfs | `styles.xml / cellXfs / xf` | References numberFormatId, fontId, fillId, borderId |
| Theme palette | `theme/theme1.xml` | ThemeSlot→ARGB with tint/shade |
