package com.tjclp.xl.addressing

import java.util.Locale

import com.tjclp.xl.error.{XLError, XLResult}

/** Opaque sheet name with validation */
opaque type SheetName = String

object SheetName:
  private val InvalidChars = Set(':', '\\', '/', '?', '*', '[', ']')
  private val MaxLength = 31
  // Excel reserved names (case-insensitive)
  private val ReservedNames = Set("history")
  // R1C1-style reference shapes (R1C1, RC, r2c, ...) that Excel reads as references, not names
  private val R1C1Shaped = "(?i)^R\\d*C\\d*$".r

  /** Create a validated sheet name */
  def apply(name: String): Either[String, SheetName] =
    if name.isEmpty then Left("Sheet name cannot be empty")
    else if name.length > MaxLength then Left(s"Sheet name too long (max $MaxLength): $name")
    else if name.exists(InvalidChars.contains) then
      Left(s"Sheet name contains invalid characters: $name")
    else if ReservedNames.contains(name.toLowerCase) then
      Left(s"Sheet name '$name' is reserved by Excel")
    else Right(name)

  /** Create an unsafe sheet name (use only when validation is guaranteed) */
  def unsafe(name: String): SheetName = name

  /**
   * True when this name must be single-quoted in A1-style formulas and references (GH-263).
   *
   * Shared by every formatter that emits sheet-qualified references (RefType.toA1, FormulaPrinter,
   * defined-name print formulas). Quoting is required for names that:
   *   - are empty or start with a digit,
   *   - contain any character outside letters/digits/underscore (spaces, quotes, `-`, ...),
   *   - parse as an A1 cell reference (`A1` through `XFD1048576`, case-insensitive) — bare `Q1!B2`
   *     would be read as something else entirely by Excel,
   *   - look like an R1C1 reference (`R1C1`, `RC`, `r2c`, ...),
   *   - spell a boolean literal (formula parsers read bare TRUE/FALSE as literals).
   */
  def needsQuoting(name: String): Boolean =
    name.isEmpty
      || name.headOption.exists(_.isDigit)
      || name.exists(c => !c.isLetterOrDigit && c != '_')
      || ARef.parse(name).isRight
      || R1C1Shaped.matches(name)
      || name.equalsIgnoreCase("TRUE")
      || name.equalsIgnoreCase("FALSE")

  /**
   * Render a sheet name for use in a formula or A1 reference: bare when safe, otherwise
   * single-quoted with embedded quotes doubled (Excel escape convention).
   */
  def quoteForFormula(name: String): String =
    if needsQuoting(name) then s"'${name.replace("'", "''")}'" else name

  /**
   * The position of the sheet `name` refers to among `names`, matched as Excel matches a
   * reference's sheet qualifier (#752): the exact spelling first, else the one name equal to it
   * ignoring case ([[folded]]). A name no sheet carries is [[XLError.SheetNotFound]]; a name that
   * folds onto two or more sheets and spells none of them exactly (tabs differing only by case,
   * which Excel refuses to save) is [[XLError.InvalidReference]] rather than a guess.
   */
  def indexIn(names: IndexedSeq[SheetName], name: SheetName): XLResult[Int] =
    names.indexOf(name) match
      case -1 =>
        val key = name.folded
        names.indices.filter(i => names(i).folded == key) match
          case Seq(only) => Right(only)
          case Seq() => Left(XLError.SheetNotFound(name, names.map(_.value).toVector))
          case many =>
            val spellings = many.map(i => s"'${names(i)}'").mkString(", ")
            Left(
              XLError.InvalidReference(
                s"sheet name '$name' matches sheets that differ only by case ($spellings); " +
                  "use the exact name"
              )
            )
      case exact => Right(exact)

  extension (name: SheetName)
    def value: String = name

    /**
     * The case-folded key two sheet names share when Excel treats them as the same name: lower case
     * under the root locale, so the result does not depend on the JVM's default locale.
     */
    def folded: String = name.toLowerCase(Locale.ROOT)

end SheetName
