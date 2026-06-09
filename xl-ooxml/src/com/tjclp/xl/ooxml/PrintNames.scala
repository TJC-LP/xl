package com.tjclp.xl.ooxml

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.api.{Sheet, Workbook}
import com.tjclp.xl.sheets.PageSetup
import com.tjclp.xl.workbooks.DefinedName

/**
 * Print area / print titles defined-name mapping (GH-259).
 *
 * Excel stores a sheet's print area and repeated title rows as sheet-scoped workbook defined names
 * (`_xlnm.Print_Area` / `_xlnm.Print_Titles` with `localSheetId`). The typed home for both is
 * `PageSetup.printArea` / `PageSetup.repeatRows`; this object converts between that model and the
 * defined-name formulas.
 *
 * Only modelable formula shapes are lifted into the model on read: a single absolute range for
 * Print_Area (`Sheet1!$A$1:$D$20`) and a pure row span for Print_Titles (`Sheet1!$1:$3`).
 * Multi-range areas, column-only titles, and hidden/commented names stay in
 * `WorkbookMetadata.definedNames` verbatim, so nothing is lost on rewrite.
 */
private[ooxml] object PrintNames:

  /** Defined name Excel uses for a sheet's print area. */
  val PrintArea = "_xlnm.Print_Area"

  /** Defined name Excel uses for a sheet's repeated print titles (rows and/or columns). */
  val PrintTitles = "_xlnm.Print_Titles"

  private val AbsRangeRe = """^\$([A-Za-z]{1,3})\$([0-9]+):\$([A-Za-z]{1,3})\$([0-9]+)$""".r
  private val AbsCellRe = """^\$([A-Za-z]{1,3})\$([0-9]+)$""".r
  private val RowSpanRe = """^\$([0-9]+):\$([0-9]+)$""".r

  /**
   * Quote a sheet name for use in a defined-name formula, matching Excel's convention: simple names
   * (letters/digits/underscore, not starting with a digit) stay bare, everything else is wrapped in
   * single quotes with embedded quotes doubled.
   */
  def quoteSheetName(name: String): String =
    val simple = name.nonEmpty && !name.charAt(0).isDigit &&
      name.forall(c => c.isLetterOrDigit || c == '_')
    if simple then name else s"'${name.replace("'", "''")}'"

  /** Format a Print_Area formula, e.g. `Sheet1!$A$1:$D$20` or `'Q1 Report'!$A$1:$D$20`. */
  def printAreaFormula(sheet: SheetName, area: CellRange): String =
    val s = area.start
    val e = area.end
    s"${quoteSheetName(sheet.value)}!$$${s.col.toLetter}$$${s.row.index1}:$$${e.col.toLetter}$$${e.row.index1}"

  /** Format a row-span Print_Titles formula, e.g. `Sheet1!$1:$3`. */
  def printTitlesFormula(sheet: SheetName, rows: (Int, Int)): String =
    s"${quoteSheetName(sheet.value)}!$$${rows._1}:$$${rows._2}"

  /**
   * Split a `Sheet!rest` or `'Quoted Sheet'!rest` formula prefix, unescaping `''` to `'`. Returns
   * None when the formula has no (single) sheet qualifier.
   */
  private def splitSheetPrefix(formula: String): Option[(String, String)] =
    if formula.startsWith("'") then
      @annotation.tailrec
      def findClose(i: Int): Option[Int] =
        if i >= formula.length then None
        else if formula.charAt(i) == '\'' then
          if i + 1 < formula.length && formula.charAt(i + 1) == '\'' then findClose(i + 2)
          else Some(i)
        else findClose(i + 1)
      findClose(1).flatMap { close =>
        val name = formula.substring(1, close).replace("''", "'")
        val rest = formula.substring(close + 1)
        if rest.startsWith("!") then Some((name, rest.drop(1))) else None
      }
    else
      val bang = formula.indexOf('!')
      if bang <= 0 then None
      else Some((formula.substring(0, bang), formula.substring(bang + 1)))

  /**
   * Parse a Print_Area formula into a CellRange when it is a single absolute range (or cell) scoped
   * to the given sheet. Multi-range areas (comma-separated) are not modeled.
   */
  def parsePrintArea(formula: String, sheet: SheetName): Option[CellRange] =
    splitSheetPrefix(formula.trim).filter((name, _) => name == sheet.value).flatMap { (_, rest) =>
      rest match
        case AbsRangeRe(c1, r1, c2, r2) => CellRange.parse(s"$c1$r1:$c2$r2").toOption
        case AbsCellRe(c, r) => CellRange.parse(s"$c$r:$c$r").toOption
        case _ => None
    }

  /**
   * Parse a Print_Titles formula into a 1-based row span when it is a pure row span scoped to the
   * given sheet. Column titles (`$A:$B`) and mixed forms are not modeled.
   */
  def parsePrintTitles(formula: String, sheet: SheetName): Option[(Int, Int)] =
    splitSheetPrefix(formula.trim).filter((name, _) => name == sheet.value).flatMap { (_, rest) =>
      rest match
        case RowSpanRe(a, b) =>
          for
            start <- a.toIntOption
            end <- b.toIntOption
            // Bounds mirror PageSetup's repeatRows invariant (total parse: never throw on read)
            if start >= 1 && start <= end && end <= com.tjclp.xl.addressing.Row.MaxIndex0 + 1
          yield (start, end)
        case _ => None
    }

  /** Defined names derived from each sheet's PageSetup (write side), in sheet order. */
  def fromSheets(sheets: Vector[Sheet]): Vector[DefinedName] =
    sheets.zipWithIndex.flatMap { case (sheet, idx) =>
      val area = sheet.pageSetup.flatMap(_.printArea).map { range =>
        DefinedName(PrintArea, printAreaFormula(sheet.name, range), localSheetId = Some(idx))
      }
      val titles = sheet.pageSetup.flatMap(_.repeatRows).map { rows =>
        DefinedName(PrintTitles, printTitlesFormula(sheet.name, rows), localSheetId = Some(idx))
      }
      area.toList ++ titles.toList
    }

  /**
   * Defined names to serialize for a workbook: the metadata names (with any print name shadowed by
   * a PageSetup-derived one removed) followed by the derived print names.
   */
  def effective(wb: Workbook): Vector[DefinedName] =
    val derived = fromSheets(wb.sheets)
    val derivedKeys = derived.map(dn => (dn.name, dn.localSheetId)).toSet
    wb.metadata.definedNames.filterNot(dn =>
      derivedKeys.contains((dn.name, dn.localSheetId))
    ) ++ derived

  /**
   * Read side: lift modelable sheet-scoped print names into each Sheet's PageSetup and drop them
   * from the metadata list (the writer re-derives them from the model). Unmodelable names are left
   * in the metadata verbatim and the corresponding PageSetup fields stay None.
   *
   * @return
   *   (sheets with printArea/repeatRows populated, defined names minus the lifted ones)
   */
  def extract(
    sheets: Vector[Sheet],
    names: Vector[DefinedName]
  ): (Vector[Sheet], Vector[DefinedName]) =
    if names.isEmpty then (sheets, names)
    else
      def candidate(name: String, idx: Int): Option[DefinedName] =
        names.find(dn =>
          dn.name == name && dn.localSheetId.contains(idx) && !dn.hidden && dn.comment.isEmpty
        )

      val parsed = sheets.zipWithIndex.map { case (sheet, idx) =>
        val area = candidate(PrintArea, idx)
          .flatMap(dn => parsePrintArea(dn.formula, sheet.name).map(dn -> _))
        val titles = candidate(PrintTitles, idx)
          .flatMap(dn => parsePrintTitles(dn.formula, sheet.name).map(dn -> _))
        (area, titles)
      }

      val updatedSheets = sheets.zip(parsed).map { case (sheet, (area, titles)) =>
        if area.isEmpty && titles.isEmpty then sheet
        else
          val base = sheet.pageSetup.getOrElse(PageSetup())
          sheet.copy(pageSetup =
            Some(base.copy(printArea = area.map(_._2), repeatRows = titles.map(_._2)))
          )
      }
      val lifted = parsed.flatMap((a, t) => a.map(_._1).toList ++ t.map(_._1).toList).toSet
      (updatedSheets, names.filterNot(lifted.contains))
