
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
  case RichText(value: richtext.RichText)  // Multiple formatted runs in one cell
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
  rowProps: Map[addr.Row, Any],
  styleRegistry: Option[style.StyleRegistry] = None  // Per-sheet style management
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
final case class Comment(
  text: RichText,
  author: Option[String]
) derives CanEqual

final case class Hyperlink(target: String) derives CanEqual  // Future
```

**Storage**: Comments stored at sheet level (`Sheet.comments: Map[ARef, Comment]`), not in Cell. This is memory-efficient (most cells don't have comments) and aligns with OOXML structure.

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

  /** Per-sheet style registry for managing cell styles with deduplication.
    * Maintains bidirectional mapping between CellStyle and integer indices.
    * Ensures consistent style application across OOXML write/read cycles.
    */
  final case class StyleRegistry(
    styles: Vector[CellStyle],           // Index → Style
    styleMap: Map[String, Int]           // CanonicalKey → Index
  ) derives CanEqual:
    def register(style: CellStyle): (StyleRegistry, Int) = ...
    def get(index: Int): Option[CellStyle] = ...
    def indexOf(style: CellStyle): Option[Int] = ...
```

## RichText Model (P31)

RichText enables multiple text runs with different formatting within a single cell:

```scala
object richtext:
  /** Single formatted text segment (maps to OOXML `<r>` element). */
  final case class TextRun(
    text: String,
    font: Option[style.Font] = None,
    color: Option[style.Color] = None,
    bold: Boolean = false,
    italic: Boolean = false,
    underline: Boolean = false
  ) derives CanEqual

  /** Collection of formatted text runs. */
  final case class RichText(runs: Vector[TextRun]) derives CanEqual:
    def +(other: TextRun): RichText = copy(runs = runs :+ other)
    def +(other: RichText): RichText = copy(runs = runs ++ other.runs)
    def toPlainText: String = runs.map(_.text).mkString

  /** DSL extensions for ergonomic RichText creation. */
  extension (s: String)
    def bold: RichText = RichText(Vector(TextRun(s, bold = true)))
    def italic: RichText = RichText(Vector(TextRun(s, italic = true)))
    def underline: RichText = RichText(Vector(TextRun(s, underline = true)))
    def red: RichText = RichText(Vector(TextRun(s, color = Some(Color.fromRgb(255, 0, 0)))))
    def green: RichText = RichText(Vector(TextRun(s, color = Some(Color.fromRgb(0, 255, 0)))))
    def blue: RichText = RichText(Vector(TextRun(s, color = Some(Color.fromRgb(0, 0, 255)))))
    def size(pt: Double): RichText = RichText(Vector(TextRun(s, font = Some(Font("Calibri", pt, false, false)))))
```

### RichText Usage
```scala
import com.tjclp.xl.richtext.*

// Composition with + operator
val text = "Error: ".red.bold + "Fix this immediately!"
val header = "Q1 ".size(18.0).bold + "Financial Report".italic

// Put in cell
sheet.put(cell"A1", CellValue.RichText(text))
```

### OOXML Mapping
- `TextRun` → `<r>` element with `<rPr>` (run properties)
- Multiple runs → `<si>` (shared string item) with multiple `<r>` children
- Whitespace preserved with `xml:space="preserve"` attribute
