package com.tjclp.xl.cells

object CellError:
  /** Parse errors from Excel notation */
  def parse(s: String): Either[String, CellError] = s match
    case "#DIV/0!" => Right(Div0)
    case "#N/A" => Right(NA)
    case "#NAME?" => Right(Name)
    case "#NULL!" => Right(Null)
    case "#NUM!" => Right(Num)
    case "#REF!" => Right(Ref)
    case "#VALUE!" => Right(Value)
    case _ => Left(s"Unknown errors: $s")

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

/** Excel errors types */
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
