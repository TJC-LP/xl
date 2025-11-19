package com.tjclp.xl.addressing

/** Opaque sheets name with validation */
opaque type SheetName = String

object SheetName:
  private val InvalidChars = Set(':', '\\', '/', '?', '*', '[', ']')
  private val MaxLength = 31

  /** Create a validated sheets name */
  def apply(name: String): Either[String, SheetName] =
    if name.isEmpty then Left("Sheet name cannot be empty")
    else if name.length > MaxLength then Left(s"Sheet name too long (max $MaxLength): $name")
    else if name.exists(InvalidChars.contains) then
      Left(s"Sheet name contains invalid characters: $name")
    else Right(name)

  /** Create an unsafe sheets name (use only when validation is guaranteed) */
  inline def unsafe(name: String): SheetName = name

  extension (name: SheetName) inline def value: String = name

end SheetName
