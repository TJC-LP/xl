package com.tjclp.xl.style

/**
 * Number format codes for Excel cell values.
 *
 * Includes built-in formats and custom format strings.
 */

// ========== Number Formats ==========

/** Number format for cell values */
enum NumFmt:
  case General
  case Integer // 0
  case Decimal // 0.00
  case ThousandsSeparator // #,##0
  case ThousandsDecimal // #,##0.00
  case Currency // $#,##0.00
  case Percent // 0%
  case PercentDecimal // 0.00%
  case Scientific // 0.00E+00
  case Fraction // # ?/?
  case Date // m/d/yy
  case DateTime // m/d/yy h:mm
  case Time // h:mm:ss
  case Text // @
  case Custom(code: String) // User-defined format

object NumFmt:
  private val builtInById: Map[Int, NumFmt] = Map(
    0 -> General,
    1 -> Integer,
    2 -> Decimal,
    3 -> ThousandsSeparator,
    4 -> ThousandsDecimal,
    5 -> Currency,
    6 -> Currency,
    7 -> Currency,
    8 -> Currency,
    9 -> Percent,
    10 -> PercentDecimal,
    11 -> Scientific,
    12 -> Fraction,
    13 -> Fraction,
    14 -> Date,
    15 -> Date,
    16 -> Date,
    17 -> Date,
    18 -> Time,
    19 -> Time,
    20 -> Time,
    21 -> Time,
    22 -> DateTime,
    49 -> Text
  )

  /** Built-in format ID mapping for OOXML */
  def builtInId(fmt: NumFmt): Option[Int] = fmt match
    case General => Some(0)
    case Integer => Some(1)
    case Decimal => Some(2)
    case ThousandsSeparator => Some(3)
    case ThousandsDecimal => Some(4)
    case Percent => Some(9)
    case PercentDecimal => Some(10)
    case Scientific => Some(11)
    case Fraction => Some(12)
    case Currency => Some(7) // Simplified; actual format varies by locale
    case Date => Some(14)
    case Time => Some(21)
    case DateTime => Some(22)
    case Text => Some(49)
    case Custom(_) => None

  /** Parse format code to NumFmt */
  def parse(code: String): NumFmt = code match
    case "General" => General
    case "0" => Integer
    case "0.00" => Decimal
    case "#,##0" => ThousandsSeparator
    case "#,##0.00" => ThousandsDecimal
    case "0%" => Percent
    case "0.00%" => PercentDecimal
    case "0.00E+00" => Scientific
    case "# ?/?" => Fraction
    case "m/d/yy" => Date
    case "m/d/yy h:mm" => DateTime
    case "h:mm:ss" => Time
    case "@" => Text
    case other => Custom(other)

  /** Resolve number format by built-in ID (optionally using an explicit format code). */
  def fromId(id: Int, formatCode: Option[String] = None): Option[NumFmt] =
    builtInById.get(id).orElse(formatCode.map(parse))
