
# Addressing & DSL — Algorithms, Macros, and Errors

## A1 Parsing Algorithm
**Goal:** parse strings like `"BC23"` into `(col = 1‑based letters → 0‑based index, row = 1‑based → 0‑based)`.

### Column label → index (base‑26 without zero)
```
acc = 0
for each ch in label:
  digit = (ch - 'A' + 1)
  acc = acc * 26 + digit
return acc - 1  // to 0-based
```

### Macro implementation sketch
```scala
package com.tjclp.xl.dsl

import scala.quoted.*
import com.tjclp.xl.core.addr.*

extension (inline sc: StringContext)
  inline def cell(inline args: Any*): ARef = ${ cellImpl('sc, 'args) }
  inline def range(inline args: Any*): CellRange = ${ rangeImpl('sc, 'args) }

private def cellImpl(sc: Expr[StringContext], args: Expr[Seq[Any]])(using Quotes): Expr[ARef] =
  import quotes.reflect.*
  val s = LiteralExtractor.oneLiteral(sc, args, "cell")
  A1.parseCell(s) match
    case Right((c0, r0)) =>
      '{ ARef(Column.from0(${Expr(c0)}), Row.from0(${Expr(r0)})) }
    case Left(msg) => report.errorAndAbort(s"Invalid cell literal '$s': $msg")

private def rangeImpl(sc: Expr[StringContext], args: Expr[Seq[Any]])(using Quotes): Expr[CellRange] =
  import quotes.reflect.*
  val s = LiteralExtractor.oneLiteral(sc, args, "range")
  A1.parseRange(s) match
    case Right(((cs, rs), (ce, re))) =>
      '{ CellRange(ARef(Column.from0(${Expr(cs)}), Row.from0(${Expr(rs)})),
                  ARef(Column.from0(${Expr(ce)}), Row.from0(${Expr(re)}))) }
    case Left(msg) => report.errorAndAbort(s"Invalid range literal '$s': $msg")
```

### Error messages
- `cell"A0"` → **“Row must start at 1 (got 0)”**
- `cell"0A"` → **“Column label must start with A..Z (got '0')”**
- `range"C5:A2"` → **normalized** to `A2:C5` (no error).

## Additional literals
- `nf"0.00%"` → `NumFmt.Percent(2)`; CT tokenization validates pattern.
- `rgb"#AABBCC"` → `Color.Rgb(0xFFAABBCC)` (alpha defaults to FF).
- `theme"accent1(-0.25)"` → `Color.Theme(Accent1, -0.25)`.

## R1C1 (later)
- Provide separate `r1c1"R10C5"` literal if demanded. Use same validation pathway.
