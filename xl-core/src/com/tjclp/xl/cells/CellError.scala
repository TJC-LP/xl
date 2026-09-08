package com.tjclp.xl.cells

object CellError:
  /**
   * Parse an Excel error literal (`#REF!`, `#SPILL!`, …) case-insensitively — Excel upper-cases one
   * at entry and writes the canonical spelling into a `t="e"` cell's `<v>`.
   */
  def parse(s: String): Either[String, CellError] =
    values
      .find(_.toExcel.equalsIgnoreCase(s))
      .toRight(s"Unknown error: $s")

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
      case GettingData => "#GETTING_DATA"
      case Spill => "#SPILL!"
      case Connect => "#CONNECT!"
      case Blocked => "#BLOCKED!"
      case Unknown => "#UNKNOWN!"
      case Field => "#FIELD!"
      case Calc => "#CALC!"

    /**
     * The number Excel's `ERROR.TYPE` returns for this error (GH-630) — Microsoft's table, in which
     * the classic seven are 1-7 and the modern codes follow in the order Excel introduced them.
     */
    def errorTypeNumber: Int = error match
      case Null => 1
      case Div0 => 2
      case Value => 3
      case Ref => 4
      case Name => 5
      case Num => 6
      case NA => 7
      case GettingData => 8
      case Spill => 9
      case Connect => 10
      case Blocked => 11
      case Unknown => 12
      case Field => 13
      case Calc => 14

/**
 * Excel error values: the classic seven plus the modern codes Excel 365 writes into `t="e"` cells
 * and accepts as formula literals (GH-630).
 */
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

  /** An external data query still running: #GETTING_DATA (the one code with no terminator) */
  case GettingData

  /** A dynamic array cannot spill into occupied cells: #SPILL! */
  case Spill

  /** A linked data type cannot reach its service: #CONNECT! */
  case Connect

  /** A linked data type is blocked by policy or privacy settings: #BLOCKED! */
  case Blocked

  /** A linked data type's provider does not recognise the value: #UNKNOWN! */
  case Unknown

  /** A field a linked data type does not have: #FIELD! */
  case Field

  /** A calculation the engine cannot express (an empty array, a LAMBDA misuse): #CALC! */
  case Calc
