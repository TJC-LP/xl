package com.tjclp.xl

import java.time.LocalDateTime

/** Cell value types supported by Excel */
enum CellValue:
  /** Text/string value */
  case Text(value: String)

  /** Numeric value (Excel uses double internally, but we preserve precision) */
  case Number(value: BigDecimal)

  /** Boolean value */
  case Bool(value: Boolean)

  /** Date/time value */
  case DateTime(value: LocalDateTime)

  /** Formula expression (stored as string, can be typed later) */
  case Formula(expression: String)

  /** Empty cell */
  case Empty

  /** Error value */
  case Error(error: CellError)

object CellValue:
  /** Smart constructor from Any */
  def from(value: Any): CellValue = value match
    case cv: CellValue => cv  // Already a CellValue, return as-is
    case s: String => Text(s)
    case i: Int => Number(BigDecimal(i))
    case l: Long => Number(BigDecimal(l))
    case d: Double => Number(BigDecimal(d))
    case bd: BigDecimal => Number(bd)
    case b: Boolean => Bool(b)
    case dt: LocalDateTime => DateTime(dt)
    case _ => Text(value.toString)

/** Excel error types */
enum CellError:
  /** Division by zero: #DIV/0! */
  case Div0

  /** Value not available: #N/A */
  case NA

  /** Invalid name: #NAME? */
  case Name

  /** Null value: #NULL! */
  case Null

  /** Invalid number: #NUM! */
  case Num

  /** Invalid reference: #REF! */
  case Ref

  /** Invalid value type: #VALUE! */
  case Value

object CellError:
  /** Parse error from Excel notation */
  def parse(s: String): Either[String, CellError] = s match
    case "#DIV/0!" => Right(Div0)
    case "#N/A" => Right(NA)
    case "#NAME?" => Right(Name)
    case "#NULL!" => Right(Null)
    case "#NUM!" => Right(Num)
    case "#REF!" => Right(Ref)
    case "#VALUE!" => Right(Value)
    case _ => Left(s"Unknown error: $s")

  extension (error: CellError)
    /** Convert to Excel notation */
    def toExcel: String = error match
      case Div0 => "#DIV/0!"
      case NA => "#N/A"
      case Name => "#NAME?"
      case Null => "#NULL!"
      case Num => "#NUM!"
      case Ref => "#REF!"
      case Value => "#VALUE!"

/** Cell with value and optional metadata */
case class Cell(
  ref: ARef,
  value: CellValue = CellValue.Empty,
  styleId: Option[Int] = None,
  comment: Option[String] = None,
  hyperlink: Option[String] = None
):
  /** Get cell column */
  def col: Column = ref.col

  /** Get cell row */
  def row: Row = ref.row

  /** Get A1 notation */
  def toA1: String = ref.toA1

  /** Update cell value */
  def withValue(newValue: CellValue): Cell = copy(value = newValue)

  /** Update cell style */
  def withStyle(styleId: Int): Cell = copy(styleId = Some(styleId))

  /** Clear cell style */
  def clearStyle: Cell = copy(styleId = None)

  /** Add comment to cell */
  def withComment(text: String): Cell = copy(comment = Some(text))

  /** Clear comment */
  def clearComment: Cell = copy(comment = None)

  /** Add hyperlink to cell */
  def withHyperlink(url: String): Cell = copy(hyperlink = Some(url))

  /** Clear hyperlink */
  def clearHyperlink: Cell = copy(hyperlink = None)

  /** Check if cell is empty */
  def isEmpty: Boolean = value == CellValue.Empty

  /** Check if cell is not empty */
  def nonEmpty: Boolean = !isEmpty

  /** Check if cell contains a formula */
  def isFormula: Boolean = value match
    case CellValue.Formula(_) => true
    case _ => false

  /** Check if cell contains an error */
  def isError: Boolean = value match
    case CellValue.Error(_) => true
    case _ => false

object Cell:
  /** Create cell from reference and value */
  def apply(ref: ARef, value: CellValue): Cell = Cell(ref, value, None, None, None)

  /** Create cell from A1 notation and value */
  def fromA1(a1: String, value: CellValue): Either[String, Cell] =
    ARef.parse(a1).map(ref => Cell(ref, value))

  /** Create empty cell */
  def empty(ref: ARef): Cell = Cell(ref, CellValue.Empty)
