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
  case Formula(expression: String)        // Stored as raw string (no evaluator yet)
  case Empty
  case Error(error: CellError)
```

- No formula AST today; expressions are stored as strings (evaluator planned in `xl-evaluator`).
- `CellValue.from` provides a best-effort conversion from common JVM types.

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
