package com.tjclp.xl.cli.commands

import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.cli.helpers.ValueParser
import com.tjclp.xl.styles.CellStyle

/**
 * Workbook comparison for the diff command (GH-137).
 *
 * Pure: `computeDiff` and the renderers never throw and perform no IO. The caller (Main.runDiff)
 * reads both workbooks and maps `identical` to the diff-tool exit-code convention (0 identical, 1
 * differs, 2 error).
 *
 * Comparison semantics:
 *   - Cells are compared by formula text for formulas (cached values are derived and ignored) and
 *     by typed value otherwise
 *   - `styleChanged` compares RESOLVED styles (styleId looked up in the sheet's registry; missing
 *     style = default), so identical formatting under different style ids is not a difference
 *   - A cell with Empty value, default resolved style, and no hyperlink is equivalent to a missing
 *     cell
 *   - Merges, comments, and hyperlinks are reported as separate per-sheet deltas
 */
object DiffCommands:

  /** Display snapshot of one side of a cell: formatted value + formula text when present. */
  final case class CellSnapshot(value: String, formula: Option[String]) derives CanEqual

  /** A cell present on both sides whose value, formula, or resolved style differs. */
  final case class CellChange(
    ref: ARef,
    before: CellSnapshot,
    after: CellSnapshot,
    styleChanged: Boolean
  )

  /** All differences for one sheet present in both workbooks. */
  final case class SheetDiff(
    name: String,
    added: Vector[(ARef, CellSnapshot)],
    removed: Vector[(ARef, CellSnapshot)],
    changed: Vector[CellChange],
    mergesAdded: Vector[String],
    mergesRemoved: Vector[String],
    commentsAdded: Vector[String],
    commentsRemoved: Vector[String],
    commentsChanged: Vector[String],
    hyperlinksAdded: Vector[String],
    hyperlinksRemoved: Vector[String],
    hyperlinksChanged: Vector[String]
  ):
    def isEmpty: Boolean =
      added.isEmpty && removed.isEmpty && changed.isEmpty &&
        mergesAdded.isEmpty && mergesRemoved.isEmpty &&
        commentsAdded.isEmpty && commentsRemoved.isEmpty && commentsChanged.isEmpty &&
        hyperlinksAdded.isEmpty && hyperlinksRemoved.isEmpty && hyperlinksChanged.isEmpty

  /**
   * Full workbook comparison. `sheets` contains only sheets (present in both workbooks) with at
   * least one difference, in workbook A's sheet order.
   */
  final case class WorkbookDiff(
    sheetsAdded: Vector[String],
    sheetsRemoved: Vector[String],
    sheets: Vector[SheetDiff]
  ):
    def identical: Boolean = sheetsAdded.isEmpty && sheetsRemoved.isEmpty && sheets.isEmpty

  /**
   * Compare two workbooks, optionally restricted to one sheet name.
   *
   * Left when the filtered sheet exists in neither workbook.
   */
  def computeDiff(
    wbA: Workbook,
    wbB: Workbook,
    sheetFilter: Option[String]
  ): Either[String, WorkbookDiff] =
    val namesA = wbA.sheets.map(_.name.value)
    val namesB = wbB.sheets.map(_.name.value)

    sheetFilter match
      case Some(filter) if !namesA.contains(filter) && !namesB.contains(filter) =>
        Left(s"Sheet '$filter' not found in either workbook")
      case _ =>
        val keep: String => Boolean = name => sheetFilter.forall(_ == name)
        val setA = namesA.toSet
        val setB = namesB.toSet
        val sheetsAdded = namesB.filter(n => keep(n) && !setA.contains(n))
        val sheetsRemoved = namesA.filter(n => keep(n) && !setB.contains(n))
        val common = namesA.filter(n => keep(n) && setB.contains(n))
        val sheetDiffs = common.flatMap { name =>
          for
            sheetA <- wbA.sheets.find(_.name.value == name)
            sheetB <- wbB.sheets.find(_.name.value == name)
            diff = diffSheet(sheetA, sheetB)
            if !diff.isEmpty
          yield diff
        }
        Right(WorkbookDiff(sheetsAdded, sheetsRemoved, sheetDiffs))

  // ==========================================================================
  // Sheet comparison
  // ==========================================================================

  private def diffSheet(sheetA: Sheet, sheetB: Sheet): SheetDiff =
    val cellsA = nonTrivialCells(sheetA)
    val cellsB = nonTrivialCells(sheetB)
    val refs = (cellsA.keySet ++ cellsB.keySet).toVector.sortBy(r => (r.row.index0, r.col.index0))

    val (added, removed, changed) = refs.foldLeft(
      (
        Vector.empty[(ARef, CellSnapshot)],
        Vector.empty[(ARef, CellSnapshot)],
        Vector.empty[CellChange]
      )
    ) { case ((add, rem, chg), ref) =>
      (cellsA.get(ref), cellsB.get(ref)) match
        case (None, Some(cellB)) => (add :+ (ref, snapshot(cellB, sheetB)), rem, chg)
        case (Some(cellA), None) => (add, rem :+ (ref, snapshot(cellA, sheetA)), chg)
        case (Some(cellA), Some(cellB)) =>
          val valueDiffers = !sameValue(cellA.value, cellB.value)
          val styleDiffers = resolvedStyleKey(cellA, sheetA) != resolvedStyleKey(cellB, sheetB)
          if valueDiffers || styleDiffers then
            (
              add,
              rem,
              chg :+ CellChange(ref, snapshot(cellA, sheetA), snapshot(cellB, sheetB), styleDiffers)
            )
          else (add, rem, chg)
        case (None, None) => (add, rem, chg)
    }

    val mergesA = sheetA.mergedRanges
    val mergesB = sheetB.mergedRanges
    def sortRanges(rs: Set[com.tjclp.xl.addressing.CellRange]): Vector[String] =
      rs.toVector
        .sortBy(r => (r.start.row.index0, r.start.col.index0, r.end.row.index0, r.end.col.index0))
        .map(_.toA1)

    val (commentsAdded, commentsRemoved, commentsChanged) =
      mapDeltas(sheetA.comments, sheetB.comments, sameComment)

    val (linksAdded, linksRemoved, linksChanged) =
      mapDeltas(hyperlinks(sheetA), hyperlinks(sheetB), (a: String, b: String) => a == b)

    SheetDiff(
      name = sheetA.name.value,
      added = added,
      removed = removed,
      changed = changed,
      mergesAdded = sortRanges(mergesB -- mergesA),
      mergesRemoved = sortRanges(mergesA -- mergesB),
      commentsAdded = commentsAdded,
      commentsRemoved = commentsRemoved,
      commentsChanged = commentsChanged,
      hyperlinksAdded = linksAdded,
      hyperlinksRemoved = linksRemoved,
      hyperlinksChanged = linksChanged
    )

  /** added/removed/changed refs (A1, row-major order) between two per-ref maps. */
  private def mapDeltas[A](
    mapA: Map[ARef, A],
    mapB: Map[ARef, A],
    same: (A, A) => Boolean
  ): (Vector[String], Vector[String], Vector[String]) =
    val refs = (mapA.keySet ++ mapB.keySet).toVector.sortBy(r => (r.row.index0, r.col.index0))
    refs.foldLeft((Vector.empty[String], Vector.empty[String], Vector.empty[String])) {
      case ((add, rem, chg), ref) =>
        (mapA.get(ref), mapB.get(ref)) match
          case (None, Some(_)) => (add :+ ref.toA1, rem, chg)
          case (Some(_), None) => (add, rem :+ ref.toA1, chg)
          case (Some(a), Some(b)) if !same(a, b) => (add, rem, chg :+ ref.toA1)
          case _ => (add, rem, chg)
    }

  private def sameComment(a: Comment, b: Comment): Boolean =
    a.text.toPlainText == b.text.toPlainText && a.author == b.author

  private def hyperlinks(sheet: Sheet): Map[ARef, String] =
    sheet.cells.flatMap((ref, cell) => cell.hyperlink.map(ref -> _))

  /** Cells that carry information: value, non-default resolved style, or hyperlink. */
  private def nonTrivialCells(sheet: Sheet): Map[ARef, Cell] =
    sheet.cells.filter { (_, cell) =>
      cell.value != CellValue.Empty ||
      resolvedStyleKey(cell, sheet) != CellStyle.default.canonicalKey ||
      cell.hyperlink.isDefined
    }

  /** Canonical key of the RESOLVED style (registry lookup; missing = default). */
  private def resolvedStyleKey(cell: Cell, sheet: Sheet): String =
    cell.styleId.flatMap(sheet.styleRegistry.get).getOrElse(CellStyle.default).canonicalKey

  /** Formulas compare by normalized formula text; everything else by typed value. */
  private def sameValue(a: CellValue, b: CellValue): Boolean =
    (a, b) match
      case (CellValue.Formula(exprA, _), CellValue.Formula(exprB, _)) =>
        normalizeFormula(exprA) == normalizeFormula(exprB)
      case (CellValue.Formula(_, _), _) | (_, CellValue.Formula(_, _)) => false
      case (va, vb) => va == vb

  private def normalizeFormula(expr: String): String =
    if expr.startsWith("=") then expr else s"=$expr"

  private def snapshot(cell: Cell, sheet: Sheet): CellSnapshot =
    cell.value match
      case CellValue.Formula(expr, cached) =>
        CellSnapshot(
          value = cached.map(ValueParser.formatCellValue).getOrElse(""),
          formula = Some(normalizeFormula(expr))
        )
      case other =>
        CellSnapshot(value = ValueParser.formatCellValue(other), formula = None)

  // ==========================================================================
  // Rendering
  // ==========================================================================

  /** Human-readable markdown report, grouped by sheet, refs in A1. */
  def renderMarkdown(diff: WorkbookDiff, fileA: String, fileB: String): String =
    val sb = new StringBuilder
    sb.append(s"Comparing $fileA vs $fileB\n")

    if diff.identical then sb.append("\nFiles are identical.\n")
    else
      if diff.sheetsAdded.nonEmpty then
        sb.append(s"\nSheets added: ${diff.sheetsAdded.mkString(", ")}\n")
      if diff.sheetsRemoved.nonEmpty then
        sb.append(s"Sheets removed: ${diff.sheetsRemoved.mkString(", ")}\n")

      diff.sheets.foreach { sd =>
        sb.append(s"\n## ${sd.name}\n")
        if sd.changed.nonEmpty then
          sb.append(s"\nChanged (${sd.changed.length}):\n")
          sd.changed.foreach { c =>
            val styleNote = if c.styleChanged then " [style]" else ""
            sb.append(s"  ${c.ref.toA1}: ${display(c.before)} -> ${display(c.after)}$styleNote\n")
          }
        if sd.added.nonEmpty then
          sb.append(s"\nAdded (${sd.added.length}):\n")
          sd.added.foreach { (ref, snap) => sb.append(s"  ${ref.toA1}: ${display(snap)}\n") }
        if sd.removed.nonEmpty then
          sb.append(s"\nRemoved (${sd.removed.length}):\n")
          sd.removed.foreach { (ref, snap) =>
            sb.append(s"  ${ref.toA1}: (was ${display(snap)})\n")
          }
        appendDelta(sb, "Merges added", sd.mergesAdded)
        appendDelta(sb, "Merges removed", sd.mergesRemoved)
        appendDelta(sb, "Comments added", sd.commentsAdded)
        appendDelta(sb, "Comments removed", sd.commentsRemoved)
        appendDelta(sb, "Comments changed", sd.commentsChanged)
        appendDelta(sb, "Hyperlinks added", sd.hyperlinksAdded)
        appendDelta(sb, "Hyperlinks removed", sd.hyperlinksRemoved)
        appendDelta(sb, "Hyperlinks changed", sd.hyperlinksChanged)
      }

      val totalChanged = diff.sheets.map(_.changed.length).sum
      val totalAdded = diff.sheets.map(_.added.length).sum
      val totalRemoved = diff.sheets.map(_.removed.length).sum
      val sheetCount =
        diff.sheets.length + diff.sheetsAdded.length + diff.sheetsRemoved.length
      sb.append(
        s"\nSummary: $totalChanged changed, $totalAdded added, $totalRemoved removed cell(s) across $sheetCount sheet(s)\n"
      )

    sb.toString

  private def appendDelta(sb: StringBuilder, label: String, refs: Vector[String]): Unit =
    if refs.nonEmpty then sb.append(s"\n$label: ${refs.mkString(", ")}\n")

  /** Formula text when present, quoted text otherwise (numbers/booleans unquoted). */
  private def display(snap: CellSnapshot): String =
    snap.formula.getOrElse {
      if snap.value.isEmpty then "(empty)"
      else if isPlainScalar(snap.value) then snap.value
      else s"\"${snap.value}\""
    }

  private def isPlainScalar(s: String): Boolean =
    s == "TRUE" || s == "FALSE" || scala.util.Try(BigDecimal(s)).isSuccess

  /**
   * Stable JSON schema:
   * {{{
   * {
   *   "identical": false,
   *   "sheetsAdded": [], "sheetsRemoved": [],
   *   "sheets": [{
   *     "name": "Sheet1",
   *     "added":   [{"ref": "D5", "value": "New", "formula": null}],
   *     "removed": [{"ref": "E5", "value": "Old", "formula": null}],
   *     "changed": [{"ref": "A5",
   *                  "before": {"value": "1", "formula": null},
   *                  "after":  {"value": "2", "formula": null},
   *                  "styleChanged": false}],
   *     "mergesAdded": [], "mergesRemoved": [],
   *     "commentsAdded": [], "commentsRemoved": [], "commentsChanged": [],
   *     "hyperlinksAdded": [], "hyperlinksRemoved": [], "hyperlinksChanged": []
   *   }]
   * }
   * }}}
   */
  def renderJson(diff: WorkbookDiff): String =
    def snapJson(snap: CellSnapshot): ujson.Obj =
      ujson.Obj(
        "value" -> ujson.Str(snap.value),
        "formula" -> snap.formula.fold[ujson.Value](ujson.Null)(ujson.Str.apply)
      )
    def cellListJson(cells: Vector[(ARef, CellSnapshot)]): ujson.Arr =
      ujson.Arr.from(cells.map { (ref, snap) =>
        val obj = snapJson(snap)
        ujson.Obj(
          "ref" -> ujson.Str(ref.toA1),
          "value" -> obj("value"),
          "formula" -> obj("formula")
        )
      })
    def strArr(values: Vector[String]): ujson.Arr =
      ujson.Arr.from(values.map(ujson.Str.apply))

    val sheetsJson = ujson.Arr.from(diff.sheets.map { sd =>
      ujson.Obj(
        "name" -> ujson.Str(sd.name),
        "added" -> cellListJson(sd.added),
        "removed" -> cellListJson(sd.removed),
        "changed" -> ujson.Arr.from(sd.changed.map { c =>
          ujson.Obj(
            "ref" -> ujson.Str(c.ref.toA1),
            "before" -> snapJson(c.before),
            "after" -> snapJson(c.after),
            "styleChanged" -> ujson.Bool(c.styleChanged)
          )
        }),
        "mergesAdded" -> strArr(sd.mergesAdded),
        "mergesRemoved" -> strArr(sd.mergesRemoved),
        "commentsAdded" -> strArr(sd.commentsAdded),
        "commentsRemoved" -> strArr(sd.commentsRemoved),
        "commentsChanged" -> strArr(sd.commentsChanged),
        "hyperlinksAdded" -> strArr(sd.hyperlinksAdded),
        "hyperlinksRemoved" -> strArr(sd.hyperlinksRemoved),
        "hyperlinksChanged" -> strArr(sd.hyperlinksChanged)
      )
    })

    val root = ujson.Obj(
      "identical" -> ujson.Bool(diff.identical),
      "sheetsAdded" -> strArr(diff.sheetsAdded),
      "sheetsRemoved" -> strArr(diff.sheetsRemoved),
      "sheets" -> sheetsJson
    )
    ujson.write(root, indent = 2)
