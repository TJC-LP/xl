
# Domain Model (Scala 3.7, Deep Dive)

This file defines the **full algebra** for cells, sheets, workbooks, styles, and metadata. All types derive `CanEqual` and prefer **opaque types** for zero‑overhead safety.

## Addressing primitives
```scala
package com.tjclp.xl.core

object addr:
  opaque type Column = Int
  object Column:
    inline def from0(i: Int): Column = i
    extension (c: Column) inline def index0: Int = c

  opaque type Row = Int
  object Row:
    inline def from0(i: Int): Row = i
    extension (r: Row) inline def index0: Int = r

  /** Absolute cell reference packed into a Long: high 32 = row, low 32 = col. */
  opaque type ARef = Long
  object ARef:
    inline def apply(c: Column, r: Row): ARef =
      ((r.toLong) << 32) | (c.toLong & 0xffffffffL)
    extension (ref: ARef)
      inline def col: Column = (ref & 0xffffffffL).toInt
      inline def row: Row    = (ref >>> 32).toInt

  final case class CellRange(start: ARef, end: ARef) derives CanEqual

  /** Unique, validated sheet name (we validate at API boundaries). */
  opaque type SheetName = String
  object SheetName:
    inline def apply(s: String): SheetName = s
    extension (n: SheetName) inline def value: String = n
```

### Invariants
- `CellRange` is always **normalized** (`start <= end` in (row, col) lexicographic order).
- `Sheet` contains a **sparse** map of `ARef → Cell`. Absent key denotes `Empty` value.
- `Workbook.sheets` has **unique names**; vector order is meaningful for UI parity.

## Values & formulas
```scala
import java.time.*

enum CellError derives CanEqual:
  case Div0, NA, Name, Null, Num, Ref, Value

enum Validity derives CanEqual: case Unvalidated, Validated
enum Operator derives CanEqual: case Add, Sub, Mul, Div, Pow

enum CellValue derives CanEqual:
  case Text(value: String)
  case Number(value: BigDecimal)    // precision first; `Double` writer optional
  case Bool(value: Boolean)
  case DateTime(value: LocalDateTime)
  case Formula(expr: formula.Expr[Validity])
  case Empty
  case Error(err: CellError)

object formula:
  import Validity.*
  enum Expr[+V <: Validity] derives CanEqual:
    case Ref(cell: addr.ARef)
    case Range(start: addr.ARef, end: addr.ARef)
    case Fn(name: String, args: List[Expr[V]])
    case Bin(l: Expr[V], op: Operator, r: Expr[V])
    case Lit(value: CellValue)
```

## Sheet & workbook
```scala
final case class Cell(
  ref: addr.ARef,
  value: CellValue,
  style: Option[style.CellStyle] = None,
  comment: Option[Comment] = None,
  hyperlink: Option[Hyperlink] = None
) derives CanEqual

final case class Sheet(
  name: addr.SheetName,
  cells: Map[addr.ARef, Cell],
  merged: Set[addr.CellRange],
  colProps: Map[addr.Column, Any],
  rowProps: Map[addr.Row, Any]
) derives CanEqual:
  def cell(ref: addr.ARef): Option[Cell] = cells.get(ref)
  def updateCell(ref: addr.ARef)(f: Cell => Cell): Sheet =
    val c0 = cells.getOrElse(ref, Cell(ref, CellValue.Empty))
    copy(cells = cells.updated(ref, f(c0)))
  def range(r: addr.CellRange): Map[addr.ARef, Cell] =
    val (cS, rS) = (r.start.col.index0, r.start.row.index0)
    val (cE, rE) = (r.end.col.index0,   r.end.row.index0)
    cells.iterator.filter { case (k, _) =>
      val c = k.col.index0; val rr = k.row.index0
      c >= cS && c <= cE && rr >= rS && rr <= rE
    }.toMap

final case class Workbook(
  sheets: Vector[Sheet],
  styles: Any,          // concrete types defined in style module
  sharedStrings: Any,   // concrete type defined in ooxml module
  metadata: Any
) derives CanEqual
```

### Complexity
- `updateCell`: O(log n) with persistent `Map` (RB‑tree). Amortized near‑O(1) with HAMT.
- `range`: O(k) for k cells matched; uses iter + predicate.

## Comments & links
```scala
final case class Comment(author: String, text: String) derives CanEqual
final case class Hyperlink(target: String) derives CanEqual
```

## Style model (summary; full detail in 05-styles-and-themes.md)
```scala
object style:
  opaque type FormatId = Int
  object FormatId: inline def apply(id: Int): FormatId = id

  enum BorderStyle: case None, Thin, Medium, Thick, Dashed, Dotted, Double
  enum HAlign: case Left, Center, Right, Justify
  enum VAlign: case Top, Middle, Bottom

  final case class Font(name: String, sizePt: Double, bold: Boolean, italic: Boolean, underline: Boolean = false) derives CanEqual
  final case class Fill(argbOrTheme: Int) derives CanEqual
  final case class Border(left: BorderStyle, right: BorderStyle, top: BorderStyle, bottom: BorderStyle) derives CanEqual
  final case class Align(horizontal: HAlign, vertical: VAlign, wrap: Boolean) derives CanEqual

  final case class CellStyle(
    font: Font,
    fill: Fill,
    border: Border,
    numberFormat: FormatId,
    align: Align
  ) derives CanEqual
```
