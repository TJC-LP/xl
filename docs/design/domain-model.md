# Domain Model (Scala 3.7, Deep Dive)

This document mirrors the **current** xl-core data model. All types derive `CanEqual` and lean on opaque types for zero-overhead safety. Package references use `com.tjclp.xl` (macros and syntax live in the same module).

## Addressing Primitives
```scala
package com.tjclp.xl.addressing

opaque type Column = Int
object Column:
  val MaxIndex0 = 16383             // A..XFD
  inline def from0(index: Int): Column = index
  inline def from1(index: Int): Column = index - 1
  def fromLetter(input: String): Either[String, Column] = ...
  extension (c: Column)
    inline def index0: Int = c
    inline def index1: Int = c + 1
    def toLetter: String = ...

opaque type Row = Int
object Row:
  inline def from0(index: Int): Row = index
  inline def from1(index: Int): Row = index - 1
  extension (r: Row)
    inline def index0: Int = r
    inline def index1: Int = r + 1

/** Absolute cell reference packed into a Long: high 32 bits = row, low 32 = col. */
opaque type ARef = Long
object ARef:
  inline def apply(col: Column, row: Row): ARef = ((row.toLong) << 32) | (col.toLong & 0xffffffffL)
  def parse(a1: String): Either[String, ARef] = ...
  extension (ref: ARef)
    inline def col: Column = (ref & 0xffffffffL).toInt
    inline def row: Row    = (ref >>> 32).toInt
    def toA1: String = ...

final case class CellRange(start: ARef, end: ARef) derives CanEqual

/** Validated sheet name */
opaque type SheetName = String
object SheetName:
  def apply(name: String): Either[String, SheetName] = ...
  extension (n: SheetName) inline def value: String = n
```

### Invariants
- `CellRange` is normalized (start <= end by row, then col).
- `Sheet` stores a sparse `Map[ARef, Cell]`; missing keys mean `CellValue.Empty`.
- `Workbook.sheets` keeps names unique and preserves ordering.

## Values & Formulas
```scala
import java.time.LocalDateTime

enum CellError derives CanEqual:
  case Div0, NA, Name, Null, Num, Ref, Value

enum CellValue derives CanEqual:
  case Text(value: String)
  case RichText(value: com.tjclp.xl.richtext.RichText)
  case Number(value: BigDecimal)
  case Bool(value: Boolean)
  case DateTime(value: LocalDateTime)
  case Formula(expression: String)        // String stored; see TExpr in xl-evaluator for typed AST
  case Empty
  case Error(error: CellError)
```

- Formula strings stored in `CellValue.Formula`; use `TExpr` in `xl-evaluator` for typed AST manipulation.
- `CellValue.from` provides a best-effort conversion from common JVM types.

### Formula AST (xl-evaluator)

For programmatic formula manipulation and future evaluation, use the typed AST:

```scala
package com.tjclp.xl.formula

/**
 * Typed expression GADT - type parameter A captures result type.
 *
 * Laws:
 *   - Round-trip: parse(print(expr)) == Right(expr)
 *   - Ring laws: Add/Mul form commutative semiring over BigDecimal
 *   - Short-circuit: And/Or respect left-to-right evaluation
 */
enum TExpr[A] derives CanEqual:
  // Core constructors
  case Lit[A](value: A) extends TExpr[A]
  case Ref[A](at: ARef, anchor: Anchor, decode: Cell => Either[CodecError, A]) extends TExpr[A]
  case PolyRef(at: ARef, anchor: Anchor = Anchor.Relative) extends TExpr[Nothing]
  case SheetRef[A](sheet: SheetName, at: ARef, anchor: Anchor, decode: Cell => Either[CodecError, A])
      extends TExpr[A]
  case SheetPolyRef(sheet: SheetName, at: ARef, anchor: Anchor = Anchor.Relative) extends TExpr[Nothing]
  case RangeRef(range: CellRange) extends TExpr[Nothing]
  case SheetRange(sheet: SheetName, range: CellRange) extends TExpr[Nothing]

  // Arithmetic (TExpr[BigDecimal])
  case Add(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case Sub(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case Mul(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case Div(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  // Comparison (TExpr[Boolean])
  case Lt(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]
  case Lte(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]
  case Gt(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]
  case Gte(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[Boolean]
  case Eq[A](x: TExpr[A], y: TExpr[A]) extends TExpr[Boolean]
  case Neq[A](x: TExpr[A], y: TExpr[A]) extends TExpr[Boolean]

  // Range aggregation + function calls
  case Aggregate(aggregatorId: String, location: TExpr.RangeLocation) extends TExpr[BigDecimal]
  case Call[A](spec: FunctionSpec[A], args: spec.Args) extends TExpr[A]

object TExpr:
  // Smart constructors
  def sum(range: CellRange): TExpr[BigDecimal] = ...
  def count(range: CellRange): TExpr[BigDecimal] = ...
  def average(range: CellRange): TExpr[BigDecimal] = ...
  def cond[A](test: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A]): TExpr[A] = ...

  // Extension methods for ergonomic construction
  extension (x: TExpr[BigDecimal])
    def +(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Add(x, y)
    def -(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Sub(x, y)
    def *(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Mul(x, y)
    def /(y: TExpr[BigDecimal]): TExpr[BigDecimal] = Div(x, y)
    def <(y: TExpr[BigDecimal]): TExpr[Boolean] = Lt(x, y)
    // ... etc
```

**Type Safety**: The GADT ensures type correctness at compile time. You cannot mix numeric and boolean operations.

**Parser/Printer**:
- `FormulaParser.parse(s: String): Either[ParseError, TExpr[?]]` — Parse formula strings
- `FormulaPrinter.print(expr: TExpr[?]): String` — Print back to Excel syntax
- Round-trip law: `parse(print(expr)) == Right(expr)` (verified by 51 property tests)

**Evaluator**: `Evaluator.eval(expr: TExpr[A], sheet: Sheet): Either[EvalError, A]` (WI-08)

## Cell, Sheet, Workbook
```scala
import com.tjclp.xl.styles.units.StyleId
import com.tjclp.xl.styles.StyleRegistry
import com.tjclp.xl.cells.Comment
import com.tjclp.xl.context.SourceContext

final case class Cell(
  ref: ARef,
  value: CellValue = CellValue.Empty,
  styleId: Option[StyleId] = None,
  comment: Option[String] = None,
  hyperlink: Option[String] = None
)

final case class Sheet(
  name: SheetName,
  cells: Map[ARef, Cell] = Map.empty,
  mergedRanges: Set[CellRange] = Set.empty,
  columnProperties: Map[Column, ColumnProperties] = Map.empty,
  rowProperties: Map[Row, RowProperties] = Map.empty,
  defaultColumnWidth: Option[Double] = None,
  defaultRowHeight: Option[Double] = None,
  styleRegistry: StyleRegistry = StyleRegistry.default,
  comments: Map[ARef, Comment] = Map.empty
):
  def apply(ref: ARef): Cell = cells.getOrElse(ref, Cell.empty(ref))
  def contains(ref: ARef): Boolean = cells.contains(ref)
  def withCellStyle(ref: ARef, style: com.tjclp.xl.styles.CellStyle): Sheet = ...
  def getRange(range: CellRange): Iterable[Cell] = ...

final case class Workbook(
  sheets: Vector[Sheet] = Vector.empty,
  metadata: WorkbookMetadata = WorkbookMetadata(),
  activeSheetIndex: Int = 0,
  sourceContext: Option[SourceContext] = None
):
  def apply(name: SheetName): XLResult[Sheet] = ...
  def put(sheet: Sheet): XLResult[Workbook] = ...
```

- `styleId` indexes into the sheet-local `StyleRegistry` (style deduplication).
- `SourceContext` attaches when reading from disk to enable surgical modification (copy unchanged ZIP parts verbatim).
- Workbook has no global style/shared-strings fields; those are constructed ad hoc inside `xl-ooxml`.

## Style Model (summary)
- Defined under `com.tjclp.xl.styles` with `CellStyle`, `Font`, `Fill`, `Border`, `Align`, `NumFmt`, and units (`Pt`, `Px`, `Emu`, `StyleId`).
- `StyleRegistry` manages bidirectional mapping `CellStyle ↔ StyleId`, with `register`/`get`/`indexOf` used by both in-memory and streaming writers.

## RichText Model
```scala
package com.tjclp.xl.richtext

final case class TextRun(
  text: String,
  font: Option[com.tjclp.xl.styles.font.Font] = None,
  color: Option[com.tjclp.xl.styles.color.Color] = None,
  bold: Boolean = false,
  italic: Boolean = false,
  underline: Boolean = false
) derives CanEqual

final case class RichText(runs: Vector[TextRun]) derives CanEqual:
  def +(other: TextRun): RichText = copy(runs = runs :+ other)
  def +(other: RichText): RichText = copy(runs = runs ++ other.runs)
  def toPlainText: String = runs.map(_.text).mkString

extension (s: String)
  def bold: RichText = RichText(Vector(TextRun(s, bold = true)))
  def italic: RichText = RichText(Vector(TextRun(s, italic = true)))
  def underline: RichText = RichText(Vector(TextRun(s, underline = true)))
  def red: RichText = RichText(Vector(TextRun(s, color = Some(com.tjclp.xl.styles.color.Color.rgb(255, 0, 0)))))
  def green: RichText = RichText(Vector(TextRun(s, color = Some(com.tjclp.xl.styles.color.Color.rgb(0, 255, 0)))))
  def blue: RichText = RichText(Vector(TextRun(s, color = Some(com.tjclp.xl.styles.color.Color.rgb(0, 0, 255)))))
  def size(pt: Double): RichText = RichText(Vector(TextRun(s, font = Some(com.tjclp.xl.styles.font.Font("Calibri", pt, bold = false, italic = false)))))
```

RichText is stored losslessly (including whitespace) via SharedStrings in `xl-ooxml` and supported in streaming write paths as inline strings.

## Comments & Hyperlinks
- Cells carry lightweight `comment: Option[String]` and `hyperlink: Option[String]`.
- Sheets also store a structured `comments: Map[ARef, Comment]` for OOXML comment parts.

## Complexity Notes
- Cell map operations are `O(log n)` with immutable `Map`; batch helpers in `Sheet.put` use local mutation for speed but return pure values.
- Range queries iterate only over matching cells (`O(k)` in number of hits).
