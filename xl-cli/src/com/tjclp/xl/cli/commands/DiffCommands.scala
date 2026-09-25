package com.tjclp.xl.cli.commands

import java.util.Locale

import scala.util.matching.Regex

import com.tjclp.xl.{Sheet, Workbook}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue, Comment}
import com.tjclp.xl.cf.{CfRule, ConditionalFormat}
import com.tjclp.xl.cli.helpers.ValueParser
import com.tjclp.xl.sheets.{DataValidation, FreezePane, SqrefShift}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.units.StyleId

/**
 * Workbook comparison for the diff command (GH-137).
 *
 * Pure: `computeDiff` and the renderers never throw and perform no IO. The caller (Main.runDiff)
 * reads both workbooks and maps `identical` to the diff-tool exit-code convention (0 identical, 1
 * differs, 2 error).
 *
 * Comparison semantics:
 *   - Formula cells are compared by formula text and record kind, then by cached value (GH-607): a
 *     formula whose text is unchanged but whose cache differs, or is present on one side only, is a
 *     change of kind `cache` — what a recalculation, a `--no-recalc` edit or a cache-stripping
 *     writer produces. `formulasOnly` restores the text-only rule. Every other cell compares by
 *     typed value
 *   - `styleChanged` compares RESOLVED styles (styleId looked up in the sheet's registry; missing
 *     style = default), so identical formatting under different style ids is not a difference
 *   - A cell with Empty value, default resolved style, and no hyperlink is equivalent to a missing
 *     cell
 *   - Merges, comments, and hyperlinks are reported as separate per-sheet deltas
 *   - Sheet structure (unless `cellsOnly`): rows and columns (effective size at 2 decimals — the
 *     explicit size, else the sheet default, else Excel's stock size — hidden, outline level with
 *     absent = 0, collapsed, resolved default style), collapsed into runs of equal changes; sheet
 *     properties (default sizes, freeze anchor, visibility); conditional formats and validations
 *     keyed by sqref with Excel's renumbered priorities, dxf ids and revision uids normalised away;
 *     the relative order of shared sheets; defined names keyed by scope and case-insensitive name,
 *     built-in `_xlnm.*` names skipped. View state, tab colour, print setup, tables, autofilter,
 *     drawings, calculation settings, theme and document properties are not compared
 */
object DiffCommands:

  /** Display snapshot of one side of a cell: formatted value + formula text when present. */
  final case class CellSnapshot(value: String, formula: Option[String]) derives CanEqual

  /**
   * What changed in a cell present on both sides (GH-607): its `value` (a constant), its `formula`
   * (text or record kind, or a formula replacing a constant and vice versa), its `cache` (same
   * formula, a different or missing cached value) or only its `style`.
   */
  enum ChangeKind derives CanEqual:
    case Value, Formula, Cache, Style

    /** The JSON `kind` field and the markdown tag. */
    def name: String = this match
      case Value => "value"
      case Formula => "formula"
      case Cache => "cache"
      case Style => "style"

  /**
   * A cell present on both sides whose value, formula, cached value or resolved style differs.
   * `styleChanged` is kept beside `kind`: a value change may also restyle the cell.
   */
  final case class CellChange(
    ref: ARef,
    before: CellSnapshot,
    after: CellSnapshot,
    styleChanged: Boolean,
    kind: ChangeKind
  )

  /** A compared structural property: its JSON `property` field and markdown label. */
  enum Property derives CanEqual:
    case Height, Width, Hidden, OutlineLevel, Collapsed, DefaultRowHeight, DefaultColumnWidth,
      FreezePanes, Visibility

    def name: String = this match
      case Height => "height"
      case Width => "width"
      case Hidden => "hidden"
      case OutlineLevel => "outlineLevel"
      case Collapsed => "collapsed"
      case DefaultRowHeight => "defaultRowHeight"
      case DefaultColumnWidth => "defaultColumnWidth"
      case FreezePanes => "freezePanes"
      case Visibility => "visibility"

  /** One side of a structural property, typed as the JSON renders it. */
  enum PropValue derives CanEqual:
    /** An effective row height (points) or column width (stored units), 2 decimals. */
    case Size(value: BigDecimal)
    case Flag(value: Boolean)
    case Level(value: Int)

    /** A freeze-pane anchor; `None` = no frozen pane. */
    case Anchor(value: Option[ARef])

    /** Sheet visibility: visible | hidden | veryHidden. */
    case State(value: String)

  final case class PropertyChange(property: Property, before: PropValue, after: PropValue)
      derives CanEqual

  /**
   * One run of rows (`"5:5"`, `"5:200"`) or columns (`"C:C"`, `"K:XFD"`) whose property changes are
   * equal; `styleChanged` compares the RESOLVED row/column default style, as for cells.
   */
  final case class AxisChange(ref: String, changes: Vector[PropertyChange], styleChanged: Boolean)
      derives CanEqual

  /** A defined name's compared content: formula text (no leading `=`) and the hidden flag. */
  final case class NameSnapshot(formula: String, hidden: Boolean) derives CanEqual

  /** A defined name added or removed; `scope` is the scope sheet's name, `None` = workbook. */
  final case class NameEntry(name: String, scope: Option[String], snapshot: NameSnapshot)
      derives CanEqual

  final case class NameChange(
    name: String,
    scope: Option[String],
    before: NameSnapshot,
    after: NameSnapshot
  ) derives CanEqual

  /** The relative order of the sheets present in both workbooks, when it differs. */
  final case class SheetOrder(before: Vector[String], after: Vector[String]) derives CanEqual

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
    hyperlinksChanged: Vector[String],
    rows: Vector[AxisChange] = Vector.empty,
    columns: Vector[AxisChange] = Vector.empty,
    properties: Vector[PropertyChange] = Vector.empty,
    conditionalFormatsAdded: Vector[String] = Vector.empty,
    conditionalFormatsRemoved: Vector[String] = Vector.empty,
    conditionalFormatsChanged: Vector[String] = Vector.empty,
    dataValidationsAdded: Vector[String] = Vector.empty,
    dataValidationsRemoved: Vector[String] = Vector.empty,
    dataValidationsChanged: Vector[String] = Vector.empty
  ):
    /** Row and column runs, sheet properties and conditional-format / validation deltas. */
    def structureChanges: Int =
      rows.length + columns.length + properties.length +
        conditionalFormatsAdded.length + conditionalFormatsRemoved.length +
        conditionalFormatsChanged.length + dataValidationsAdded.length +
        dataValidationsRemoved.length + dataValidationsChanged.length

    def isEmpty: Boolean =
      added.isEmpty && removed.isEmpty && changed.isEmpty &&
        mergesAdded.isEmpty && mergesRemoved.isEmpty &&
        commentsAdded.isEmpty && commentsRemoved.isEmpty && commentsChanged.isEmpty &&
        hyperlinksAdded.isEmpty && hyperlinksRemoved.isEmpty && hyperlinksChanged.isEmpty &&
        structureChanges == 0

  /**
   * Full workbook comparison. `sheets` contains only sheets (present in both workbooks) with at
   * least one difference, in workbook A's sheet order.
   */
  final case class WorkbookDiff(
    sheetsAdded: Vector[String],
    sheetsRemoved: Vector[String],
    sheets: Vector[SheetDiff],
    sheetOrder: Option[SheetOrder] = None,
    namesAdded: Vector[NameEntry] = Vector.empty,
    namesRemoved: Vector[NameEntry] = Vector.empty,
    namesChanged: Vector[NameChange] = Vector.empty
  ):
    /** Every structural difference: per-sheet ones, a sheet reorder and defined-name deltas. */
    def structureChanges: Int =
      sheets.map(_.structureChanges).sum + (if sheetOrder.isDefined then 1 else 0) +
        namesAdded.length + namesRemoved.length + namesChanged.length

    def identical: Boolean =
      sheetsAdded.isEmpty && sheetsRemoved.isEmpty && sheets.isEmpty && sheetOrder.isEmpty &&
        namesAdded.isEmpty && namesRemoved.isEmpty && namesChanged.isEmpty

  /**
   * Compare two workbooks, optionally restricted to one sheet name (its structure only, no sheet
   * order, and only the names scoped to it). `formulasOnly` compares formula cells by text alone,
   * ignoring cached values (the rule before 0.22.0); `cellsOnly` skips the sheet structure — cells,
   * merges, comments, hyperlinks and sheets added or removed only.
   *
   * Left when the filtered sheet exists in neither workbook.
   */
  def computeDiff(
    wbA: Workbook,
    wbB: Workbook,
    sheetFilter: Option[String],
    formulasOnly: Boolean = false,
    cellsOnly: Boolean = false
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
            diff = diffSheet(
              sheetA,
              sheetB,
              wbA.getSheetState(sheetA.name),
              wbB.getSheetState(sheetB.name),
              formulasOnly,
              cellsOnly
            )
            if !diff.isEmpty
          yield diff
        }
        if cellsOnly then Right(WorkbookDiff(sheetsAdded, sheetsRemoved, sheetDiffs))
        else
          // under -s only the filtered sheet's own names are compared, and order is not
          val order = if sheetFilter.isDefined then None else sheetOrder(namesA, namesB)
          val (namesAdded, namesRemoved, namesChanged) =
            nameDeltas(wbA, wbB, scope => sheetFilter.forall(f => scope.contains(f)))
          Right(
            WorkbookDiff(
              sheetsAdded,
              sheetsRemoved,
              sheetDiffs,
              order,
              namesAdded,
              namesRemoved,
              namesChanged
            )
          )

  // ==========================================================================
  // Sheet comparison
  // ==========================================================================

  private def diffSheet(
    sheetA: Sheet,
    sheetB: Sheet,
    stateA: Option[String],
    stateB: Option[String],
    formulasOnly: Boolean,
    cellsOnly: Boolean
  ): SheetDiff =
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
          val styleDiffers = resolvedStyleKey(cellA, sheetA) != resolvedStyleKey(cellB, sheetB)
          val kind = valueChange(cellA.value, cellB.value, formulasOnly)
            .orElse(Option.when(styleDiffers)(ChangeKind.Style))
          kind match
            case Some(k) =>
              val change =
                CellChange(ref, snapshot(cellA, sheetA), snapshot(cellB, sheetB), styleDiffers, k)
              (add, rem, chg :+ change)
            case None => (add, rem, chg)
        case (None, None) => (add, rem, chg)
    }

    val mergesA = sheetA.mergedRanges
    val mergesB = sheetB.mergedRanges
    def sortRanges(rs: Set[CellRange]): Vector[String] =
      rs.toVector
        .sortBy(r => (r.start.row.index0, r.start.col.index0, r.end.row.index0, r.end.col.index0))
        .map(_.toA1)

    val (commentsAdded, commentsRemoved, commentsChanged) =
      mapDeltas(sheetA.comments, sheetB.comments, sameComment)

    val (linksAdded, linksRemoved, linksChanged) =
      mapDeltas(hyperlinks(sheetA), hyperlinks(sheetB), (a: String, b: String) => a == b)

    val cellDiff = SheetDiff(
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
    if cellsOnly then cellDiff
    else
      val (cfAdded, cfRemoved, cfChanged) = blockDeltas(cfIndex(sheetA), cfIndex(sheetB))
      val (dvAdded, dvRemoved, dvChanged) = blockDeltas(dvIndex(sheetA), dvIndex(sheetB))
      cellDiff.copy(
        rows = axisChanges(
          rowAxis(sheetA),
          rowAxis(sheetB),
          sheetA,
          sheetB,
          Property.Height,
          sheetA.defaultRowHeight,
          sheetB.defaultRowHeight,
          StockRowHeight,
          (first, last) => s"${first + 1}:${last + 1}"
        ),
        columns = axisChanges(
          columnAxis(sheetA),
          columnAxis(sheetB),
          sheetA,
          sheetB,
          Property.Width,
          sheetA.defaultColumnWidth,
          sheetB.defaultColumnWidth,
          StockColumnWidth,
          (first, last) => s"${Column.from0(first).toLetter}:${Column.from0(last).toLetter}"
        ),
        properties = sheetProperties(sheetA, sheetB, stateA, stateB),
        conditionalFormatsAdded = cfAdded,
        conditionalFormatsRemoved = cfRemoved,
        conditionalFormatsChanged = cfChanged,
        dataValidationsAdded = dvAdded,
        dataValidationsRemoved = dvRemoved,
        dataValidationsChanged = dvChanged
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

  // ==========================================================================
  // Sheet structure
  // ==========================================================================

  /** Excel's stock row height in points: the Calibri 11 row. */
  private val StockRowHeight: Double = 15.0

  /**
   * Excel's stock column width in stored `<col width>` units: the 64-pixel Calibri 11 column the UI
   * shows as 8.43, and exactly what Excel writes for a default-width styled column.
   */
  private val StockColumnWidth: Double = 9.140625

  /** The compared row or column properties; `size` is the height or the width. */
  private final case class AxisProps(
    size: Option[Double],
    hidden: Boolean,
    styleId: Option[StyleId],
    outlineLevel: Option[Int],
    collapsed: Boolean
  )

  private object AxisProps:
    val empty: AxisProps = AxisProps(None, false, None, None, false)

  private def rowAxis(sheet: Sheet): Map[Int, AxisProps] =
    sheet.rowProperties.map { (row, p) =>
      row.index0 -> AxisProps(p.height, p.hidden, p.styleId, p.outlineLevel, p.collapsed)
    }

  private def columnAxis(sheet: Sheet): Map[Int, AxisProps] =
    sheet.columnProperties.map { (col, p) =>
      col.index0 -> AxisProps(p.width, p.hidden, p.styleId, p.outlineLevel, p.collapsed)
    }

  /** A stored size at 2 decimals; `None` for a non-finite one (`"NaN"` parses as a Double). */
  private def sizeOf(d: Double): Option[BigDecimal] =
    Option.when(d.isFinite)(BigDecimal(d).setScale(2, BigDecimal.RoundingMode.HALF_UP))

  /** The explicit size, else the sheet default, else Excel's stock size. */
  private def effectiveSize(
    explicit: Option[Double],
    sheetDefault: Option[Double],
    stock: Double
  ): BigDecimal =
    explicit
      .flatMap(sizeOf)
      .orElse(sheetDefault.flatMap(sizeOf))
      .getOrElse(BigDecimal(stock).setScale(2, BigDecimal.RoundingMode.HALF_UP))

  private def change(p: Property, a: PropValue, b: PropValue): Option[PropertyChange] =
    Option.when(a != b)(PropertyChange(p, a, b))

  /**
   * Per-index comparison over the rows (or columns) either side lists, collapsed into runs of
   * consecutive indices with equal changes. A size is compared only where a side states one: a
   * sheet-default change is a sheet property, never a line per row.
   */
  private def axisChanges(
    a: Map[Int, AxisProps],
    b: Map[Int, AxisProps],
    sheetA: Sheet,
    sheetB: Sheet,
    sizeProperty: Property,
    defaultA: Option[Double],
    defaultB: Option[Double],
    stock: Double,
    ref: (Int, Int) => String
  ): Vector[AxisChange] =
    val styleKeyA = styleKeys(a, sheetA)
    val styleKeyB = styleKeys(b, sheetB)
    val perIndex = (a.keySet ++ b.keySet).toVector.sorted.flatMap { i =>
      val pa = a.getOrElse(i, AxisProps.empty)
      val pb = b.getOrElse(i, AxisProps.empty)
      val explicitA = pa.size.filter(_.isFinite)
      val explicitB = pb.size.filter(_.isFinite)
      val size =
        if explicitA.isEmpty && explicitB.isEmpty then None
        else
          change(
            sizeProperty,
            PropValue.Size(effectiveSize(explicitA, defaultA, stock)),
            PropValue.Size(effectiveSize(explicitB, defaultB, stock))
          )
      val changes = Vector(
        size,
        change(Property.Hidden, PropValue.Flag(pa.hidden), PropValue.Flag(pb.hidden)),
        change(
          Property.OutlineLevel,
          PropValue.Level(pa.outlineLevel.getOrElse(0)),
          PropValue.Level(pb.outlineLevel.getOrElse(0))
        ),
        change(Property.Collapsed, PropValue.Flag(pa.collapsed), PropValue.Flag(pb.collapsed))
      ).flatten
      val styleChanged = styleKeyA(pa.styleId) != styleKeyB(pb.styleId)
      Option.when(changes.nonEmpty || styleChanged)((i, changes, styleChanged))
    }
    val runs = perIndex.foldLeft(Vector.empty[(Int, Int, Vector[PropertyChange], Boolean)]) {
      case (acc, (i, changes, style)) =>
        acc.lastOption match
          case Some((first, last, c, st)) if last + 1 == i && c == changes && st == style =>
            acc.updated(acc.length - 1, (first, i, c, st))
          case _ => acc :+ ((i, i, changes, style))
    }
    runs.map((first, last, changes, style) => AxisChange(ref(first, last), changes, style))

  /** Resolved style keys for the ids an axis uses, computed once per distinct id. */
  private def styleKeys(axis: Map[Int, AxisProps], sheet: Sheet): Option[StyleId] => String =
    val table =
      axis.values.flatMap(_.styleId).toSet.map(id => id -> resolvedKey(Some(id), sheet)).toMap
    val default = CellStyle.default.canonicalKey
    id => id.flatMap(table.get).getOrElse(default)

  private def sheetProperties(
    sheetA: Sheet,
    sheetB: Sheet,
    stateA: Option[String],
    stateB: Option[String]
  ): Vector[PropertyChange] =
    def anchor(sheet: Sheet): Option[ARef] =
      sheet.freezePane.collect { case FreezePane.At(a, _) => a }
    def visibility(state: Option[String]): String =
      state.filterNot(_ == "visible").getOrElse("visible")
    Vector(
      change(
        Property.DefaultRowHeight,
        PropValue.Size(effectiveSize(None, sheetA.defaultRowHeight, StockRowHeight)),
        PropValue.Size(effectiveSize(None, sheetB.defaultRowHeight, StockRowHeight))
      ),
      change(
        Property.DefaultColumnWidth,
        PropValue.Size(effectiveSize(None, sheetA.defaultColumnWidth, StockColumnWidth)),
        PropValue.Size(effectiveSize(None, sheetB.defaultColumnWidth, StockColumnWidth))
      ),
      change(
        Property.FreezePanes,
        PropValue.Anchor(anchor(sheetA)),
        PropValue.Anchor(anchor(sheetB))
      ),
      change(
        Property.Visibility,
        PropValue.State(visibility(stateA)),
        PropValue.State(visibility(stateB))
      )
    ).flatten

  /** A conditional-format / validation block's sqref key; unparseable keys sort last. */
  private final case class BlockKey(row: Int, col: Int, text: String)

  private object BlockKey:
    given Ordering[BlockKey] = Ordering.by(k => (k.row, k.col, k.text))

  private def sortedRanges(ranges: Vector[CellRange]): Vector[CellRange] =
    ranges.sortBy(r => (r.start.row.index0, r.start.col.index0, r.end.row.index0, r.end.col.index0))

  /** Ranges sorted row-major and printed as Excel's sqref tokens, so token order is irrelevant. */
  private def rangesKey(ranges: Vector[CellRange]): BlockKey =
    val sorted = sortedRanges(ranges)
    val text = sorted.map(SqrefShift.toSqrefToken).mkString(" ")
    sorted.headOption.fold(BlockKey(Int.MaxValue, Int.MaxValue, text)) { first =>
      BlockKey(first.start.row.index0, first.start.col.index0, text)
    }

  private val SqrefAttr = """(^|\s)sqref="([^"]*)"""".r

  /** A Preserved payload's key: its first `sqref`, parsed like a typed envelope when it can be. */
  private def payloadKey(xml: String): BlockKey =
    SqrefAttr.findFirstMatchIn(xml).map(_.group(2).trim) match
      case None => BlockKey(Int.MaxValue, Int.MaxValue, "(unparsed)")
      case Some(sqref) =>
        val tokens = sqref.split("\\s+").toVector.filter(_.nonEmpty)
        val parsed = tokens.map { token =>
          CellRange
            .parse(token)
            .toOption
            .orElse(ARef.parse(token).toOption.map(r => CellRange(r, r)))
        }
        if tokens.nonEmpty && parsed.forall(_.isDefined) then rangesKey(parsed.flatten)
        else BlockKey(Int.MaxValue, Int.MaxValue, sqref)

  /** A parsed key's text replaces the payload's `sqref`: token order is not content. */
  private def canonicalSqref(xml: String, key: BlockKey): String =
    if key.row == Int.MaxValue then xml
    else SqrefAttr.replaceFirstIn(xml, "$1sqref=\"" + Regex.quoteReplacement(key.text) + "\"")

  private val UidAttr = """\s[A-Za-z_][\w.-]*:uid="[^"]*"""".r
  private val PriorityAttr = """\spriority="([^"]*)"""".r
  private val DxfIdAttr = """\sdxfId="[^"]*"""".r

  /** Drop revision GUIDs (`xr:uid`) and, on request, the renumbered `priority` / `dxfId`. */
  private def stripVolatile(xml: String, dropPriority: Boolean, dropDxfId: Boolean): String =
    val noUid = UidAttr.replaceAllIn(xml, "")
    val noPriority = if dropPriority then PriorityAttr.replaceAllIn(noUid, "") else noUid
    if dropDxfId then DxfIdAttr.replaceAllIn(noPriority, "") else noPriority

  /**
   * Priority replaced by its worksheet-wide rank (Excel renumbers on save); a Preserved rule keeps
   * its dxf payload. Relative order across blocks matters when their ranges overlap.
   */
  private def normalizeRule(rule: CfRule, rank: Int): CfRule = rule match
    case CfRule.Preserved(xml, _, dxf) =>
      CfRule.Preserved(stripVolatile(xml, true, true), Some(rank), dxf)
    case typed => CfRule.withPriority(typed, rank)

  /**
   * Canonical precedence of every rule, ordered as the painter does: priority, then document order.
   * Keys are (block index, rule index), or the priority attribute's offset in an opaque block.
   * Counting opaque rules too keeps their precedence relative to typed blocks observable.
   */
  private def cfPriorityRanks(sheet: Sheet): Map[(Int, Int), Int] =
    val priorities = sheet.conditionalFormats.zipWithIndex.flatMap {
      case (ConditionalFormat.Rules(_, rules, _), block) =>
        rules.zipWithIndex.map { (rule, i) =>
          (block, i) -> CfRule.priorityOf(rule).getOrElse(Int.MaxValue)
        }
      case (ConditionalFormat.Preserved(xml), block) =>
        PriorityAttr
          .findAllMatchIn(xml)
          .map { m =>
            (block, m.start) -> m.group(1).toIntOption.getOrElse(Int.MaxValue)
          }
          .toVector
    }
    priorities
      .sortBy { case ((block, rule), priority) => (priority, block, rule) }
      .zipWithIndex
      .map { case ((key, _), i) => key -> (i + 1) }
      .toMap

  private def cfIndex(sheet: Sheet): Map[BlockKey, Vector[ConditionalFormat]] =
    val ranks = cfPriorityRanks(sheet)
    def rank(block: Int, rule: Int): Int = ranks.getOrElse((block, rule), 0)
    sheet.conditionalFormats.zipWithIndex
      .map {
        case (ConditionalFormat.Rules(ranges, rules, pivot), block) =>
          val ordered = rules.zipWithIndex.sortBy((_, i) => rank(block, i))
          val normalized = ordered.map((rule, i) => normalizeRule(rule, rank(block, i)))
          rangesKey(ranges) ->
            ConditionalFormat.Rules(sortedRanges(ranges), normalized, pivot)
        case (ConditionalFormat.Preserved(xml), block) =>
          val key = payloadKey(xml)
          val ordered =
            PriorityAttr.replaceAllIn(xml, m => s" priority=\"${rank(block, m.start)}\"")
          key -> ConditionalFormat.Preserved(
            canonicalSqref(stripVolatile(ordered, false, false), key)
          )
      }
      .groupMap(_._1)(_._2)

  private def dvIndex(sheet: Sheet): Map[BlockKey, Vector[DataValidation]] =
    sheet.dataValidations
      .map {
        case r: DataValidation.Rules =>
          rangesKey(r.ranges) -> r.copy(ranges = sortedRanges(r.ranges))
        case DataValidation.Preserved(xml) =>
          val key = payloadKey(xml)
          key -> DataValidation.Preserved(canonicalSqref(stripVolatile(xml, false, false), key))
      }
      .groupMap(_._1)(_._2)

  /** added/removed/changed sqref keys; several blocks under one key compare in document order. */
  private def blockDeltas[A](
    a: Map[BlockKey, Vector[A]],
    b: Map[BlockKey, Vector[A]]
  ): (Vector[String], Vector[String], Vector[String]) =
    val keys = (a.keySet ++ b.keySet).toVector.sorted
    (
      keys.filter(k => !a.contains(k)).map(_.text),
      keys.filter(k => !b.contains(k)).map(_.text),
      keys.filter(k => a.get(k).exists(va => b.get(k).exists(_ != va))).map(_.text)
    )

  /** The relative order of the shared sheets; an insertion or a removal is not a reorder. */
  private def sheetOrder(namesA: Vector[String], namesB: Vector[String]): Option[SheetOrder] =
    val setA = namesA.toSet
    val setB = namesB.toSet
    val orderA = namesA.filter(setB)
    val orderB = namesB.filter(setA)
    Option.when(orderA != orderB)(SheetOrder(orderA, orderB))

  /**
   * Defined names keyed by (scope sheet name, lower-cased name): Excel's case-insensitive identity,
   * following a scoped name's sheet across a reorder. Built-in `_xlnm.*` names are print and filter
   * machinery and are skipped.
   */
  private def nameIndex(wb: Workbook): Map[(Option[String], String), NameEntry] =
    wb.metadata.definedNames
      .filterNot(_.name.toLowerCase(Locale.ROOT).startsWith("_xlnm."))
      .map { dn =>
        // `[` and `]` are illegal in sheet names, so an orphan's label cannot collide
        val scope = dn.localSheetId.map { idx =>
          wb.sheets.lift(idx).fold(s"[localSheetId $idx]")(_.name.value)
        }
        val formula = if dn.formula.startsWith("=") then dn.formula.drop(1) else dn.formula
        (scope, dn.name.toLowerCase(Locale.ROOT)) ->
          NameEntry(dn.name, scope, NameSnapshot(formula, dn.hidden))
      }
      .sortBy((_, e) => (e.name, e.snapshot.formula, e.snapshot.hidden))
      .toMap

  private def nameDeltas(
    wbA: Workbook,
    wbB: Workbook,
    keepScope: Option[String] => Boolean
  ): (Vector[NameEntry], Vector[NameEntry], Vector[NameChange]) =
    val a = nameIndex(wbA).filter((key, _) => keepScope(key._1))
    val b = nameIndex(wbB).filter((key, _) => keepScope(key._1))
    val keys = (a.keySet ++ b.keySet).toVector.sortBy((scope, lower) =>
      (scope.isDefined, scope.getOrElse(""), lower)
    )
    val added = keys.flatMap(k => if a.contains(k) then None else b.get(k))
    val removed = keys.flatMap(k => if b.contains(k) then None else a.get(k))
    val changed = keys.flatMap { k =>
      for
        ea <- a.get(k)
        eb <- b.get(k)
        if ea.snapshot != eb.snapshot
      yield NameChange(ea.name, ea.scope, ea.snapshot, eb.snapshot)
    }
    (added, removed, changed)

  /** Cells that carry information: value, non-default resolved style, or hyperlink. */
  private def nonTrivialCells(sheet: Sheet): Map[ARef, Cell] =
    sheet.cells.filter { (_, cell) =>
      cell.value != CellValue.Empty ||
      resolvedStyleKey(cell, sheet) != CellStyle.default.canonicalKey ||
      cell.hyperlink.isDefined
    }

  /** Canonical key of the RESOLVED style (registry lookup; missing = default). */
  private def resolvedStyleKey(cell: Cell, sheet: Sheet): String =
    resolvedKey(cell.styleId, sheet)

  private def resolvedKey(styleId: Option[StyleId], sheet: Sheet): String =
    styleId.flatMap(sheet.styleRegistry.get).getOrElse(CellStyle.default).canonicalKey

  /**
   * The kind of value change between two cells, `None` when they agree. Formulas compare by
   * normalized text and record kind, then — unless `formulasOnly` — by cached value (GH-607);
   * everything else by typed value.
   */
  private def valueChange(a: CellValue, b: CellValue, formulasOnly: Boolean): Option[ChangeKind] =
    (a, b) match
      case (CellValue.Formula(exprA, cachedA, kindA), CellValue.Formula(exprB, cachedB, kindB)) =>
        if normalizeFormula(exprA) != normalizeFormula(exprB) || kindA != kindB then
          Some(ChangeKind.Formula)
        else if !formulasOnly && cachedA != cachedB then Some(ChangeKind.Cache)
        else None
      case (CellValue.Formula(_, _, _), _) | (_, CellValue.Formula(_, _, _)) =>
        Some(ChangeKind.Formula)
      case (va, vb) => if va == vb then None else Some(ChangeKind.Value)

  private def normalizeFormula(expr: String): String =
    if expr.startsWith("=") then expr else s"=$expr"

  private def snapshot(cell: Cell, sheet: Sheet): CellSnapshot =
    cell.value match
      case CellValue.Formula(expr, cached, _) =>
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
      diff.sheetOrder.foreach { order =>
        sb.append(
          s"\nSheet order: ${order.before.mkString(", ")} -> ${order.after.mkString(", ")}\n"
        )
      }
      appendNames(sb, "Names added", diff.namesAdded)
      appendNames(sb, "Names removed", diff.namesRemoved)
      if diff.namesChanged.nonEmpty then
        sb.append(s"\nNames changed (${diff.namesChanged.length}):\n")
        diff.namesChanged.foreach { c =>
          val formula =
            if c.before.formula == c.after.formula then c.before.formula
            else s"${c.before.formula} -> ${c.after.formula}"
          val hiddenNote =
            if c.before.hidden == c.after.hidden then ""
            else s" [hidden ${c.before.hidden} -> ${c.after.hidden}]"
          sb.append(s"  ${nameLabel(c.name, c.scope)}: $formula$hiddenNote\n")
        }

      diff.sheets.foreach { sd =>
        sb.append(s"\n## ${sd.name}\n")
        if sd.changed.nonEmpty then
          sb.append(s"\nChanged (${sd.changed.length}):\n")
          sd.changed.foreach { c =>
            val styleNote = if c.styleChanged then " [style]" else ""
            c.kind match
              case ChangeKind.Cache =>
                // the formula is the same on both sides: show it once, then the two caches
                val formula = c.before.formula.getOrElse(display(c.before))
                sb.append(
                  s"  ${c.ref.toA1}: $formula cached ${cachedDisplay(c.before)} -> " +
                    s"${cachedDisplay(c.after)} [cache]$styleNote\n"
                )
              case _ =>
                sb.append(
                  s"  ${c.ref.toA1}: ${display(c.before)} -> ${display(c.after)}$styleNote\n"
                )
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
        appendAxis(sb, "Rows", sd.rows)
        appendAxis(sb, "Columns", sd.columns)
        if sd.properties.nonEmpty then
          sb.append(s"\nSheet properties (${sd.properties.length}):\n")
          sd.properties.foreach(p => sb.append(s"  ${changeText(p)}\n"))
        appendDelta(sb, "Conditional formats added", sd.conditionalFormatsAdded)
        appendDelta(sb, "Conditional formats removed", sd.conditionalFormatsRemoved)
        appendDelta(sb, "Conditional formats changed", sd.conditionalFormatsChanged)
        appendDelta(sb, "Data validations added", sd.dataValidationsAdded)
        appendDelta(sb, "Data validations removed", sd.dataValidationsRemoved)
        appendDelta(sb, "Data validations changed", sd.dataValidationsChanged)
      }

      val totalChanged = diff.sheets.map(_.changed.length).sum
      val totalAdded = diff.sheets.map(_.added.length).sum
      val totalRemoved = diff.sheets.map(_.removed.length).sum
      val sheetCount =
        diff.sheets.length + diff.sheetsAdded.length + diff.sheetsRemoved.length
      val structure = diff.structureChanges
      val structureNote = if structure > 0 then s"; $structure structure change(s)" else ""
      sb.append(
        s"\nSummary: $totalChanged changed, $totalAdded added, $totalRemoved removed cell(s) across $sheetCount sheet(s)$structureNote\n"
      )

    sb.toString

  private def appendDelta(sb: StringBuilder, label: String, refs: Vector[String]): Unit =
    if refs.nonEmpty then sb.append(s"\n$label: ${refs.mkString(", ")}\n")

  /** `  5:7: height 15 -> 30, hidden false -> true [style]` lines under `Rows (n):`. */
  private def appendAxis(sb: StringBuilder, label: String, runs: Vector[AxisChange]): Unit =
    if runs.nonEmpty then
      sb.append(s"\n$label (${runs.length}):\n")
      runs.foreach { run =>
        val changes = run.changes.map(changeText).mkString(", ")
        val style =
          if !run.styleChanged then "" else if changes.isEmpty then "[style]" else " [style]"
        sb.append(s"  ${run.ref}: $changes$style\n")
      }

  private def appendNames(sb: StringBuilder, label: String, names: Vector[NameEntry]): Unit =
    if names.nonEmpty then
      sb.append(s"\n$label (${names.length}):\n")
      names.foreach { n =>
        val hidden = if n.snapshot.hidden then " [hidden]" else ""
        sb.append(s"  ${nameLabel(n.name, n.scope)} = ${n.snapshot.formula}$hidden\n")
      }

  /** `Name` for a workbook-scoped name, `Sheet!Name` (quoted when needed) for a scoped one. */
  private def nameLabel(name: String, scope: Option[String]): String =
    scope.fold(name)(sheet => s"${SheetName.quoteForFormula(sheet)}!$name")

  private def changeText(c: PropertyChange): String =
    s"${c.property.name} ${valueText(c.before)} -> ${valueText(c.after)}"

  private def valueText(v: PropValue): String = v match
    case PropValue.Size(d) => d.bigDecimal.stripTrailingZeros.toPlainString
    case PropValue.Flag(b) => b.toString
    case PropValue.Level(l) => l.toString
    case PropValue.Anchor(a) => a.fold("(none)")(_.toA1)
    case PropValue.State(s) => s

  /** Formula text when present, quoted text otherwise (numbers/booleans unquoted). */
  private def display(snap: CellSnapshot): String =
    snap.formula.getOrElse {
      if snap.value.isEmpty then "(empty)"
      else if isPlainScalar(snap.value) then snap.value
      else s"\"${snap.value}\""
    }

  private def isPlainScalar(s: String): Boolean =
    s == "TRUE" || s == "FALSE" || scala.util.Try(BigDecimal(s)).isSuccess

  /** A formula's cached value as text, `(none)` when the side has no cache. */
  private def cachedDisplay(snap: CellSnapshot): String =
    if snap.value.isEmpty then "(none)"
    else if isPlainScalar(snap.value) then snap.value
    else s"\"${snap.value}\""

  /**
   * Stable JSON schema (every key always present):
   * {{{
   * {
   *   "identical": false,
   *   "sheetsAdded": [], "sheetsRemoved": [],
   *   "sheetOrder": null,          // or {"before": ["A", "B"], "after": ["B", "A"]}
   *   "namesAdded": [{"name": "Rate", "scope": null, "formula": "Model!$B$1", "hidden": false}],
   *   "namesRemoved": [],
   *   "namesChanged": [{"name": "Rate", "scope": "Model",
   *                     "before": {"formula": "1", "hidden": false},
   *                     "after":  {"formula": "2", "hidden": false}}],
   *   "sheets": [{
   *     "name": "Sheet1",
   *     "added":   [{"ref": "D5", "value": "New", "formula": null}],
   *     "removed": [{"ref": "E5", "value": "Old", "formula": null}],
   *     "changed": [{"ref": "A5",
   *                  "before": {"value": "1", "formula": null},
   *                  "after":  {"value": "2", "formula": null},
   *                  "styleChanged": false,
   *                  "kind": "value"}],
   *     "mergesAdded": [], "mergesRemoved": [],
   *     "commentsAdded": [], "commentsRemoved": [], "commentsChanged": [],
   *     "hyperlinksAdded": [], "hyperlinksRemoved": [], "hyperlinksChanged": [],
   *     "rows": [{"ref": "5:5",
   *               "changes": [{"property": "height", "before": 15, "after": 30}],
   *               "styleChanged": false}],
   *     "columns": [],
   *     "properties": [{"property": "freezePanes", "before": null, "after": "B2"}],
   *     "conditionalFormatsAdded": [], "conditionalFormatsRemoved": [],
   *     "conditionalFormatsChanged": [],
   *     "dataValidationsAdded": [], "dataValidationsRemoved": [], "dataValidationsChanged": []
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
    def valueJson(v: PropValue): ujson.Value = v match
      case PropValue.Size(d) => ujson.Num(d.toDouble)
      case PropValue.Flag(b) => ujson.Bool(b)
      case PropValue.Level(l) => ujson.Num(l.toDouble)
      case PropValue.Anchor(a) => a.fold[ujson.Value](ujson.Null)(r => ujson.Str(r.toA1))
      case PropValue.State(s) => ujson.Str(s)
    def changeJson(c: PropertyChange): ujson.Obj =
      ujson.Obj(
        "property" -> ujson.Str(c.property.name),
        "before" -> valueJson(c.before),
        "after" -> valueJson(c.after)
      )
    def axisJson(runs: Vector[AxisChange]): ujson.Arr =
      ujson.Arr.from(runs.map { run =>
        ujson.Obj(
          "ref" -> ujson.Str(run.ref),
          "changes" -> ujson.Arr.from(run.changes.map(changeJson)),
          "styleChanged" -> ujson.Bool(run.styleChanged)
        )
      })
    def scopeJson(scope: Option[String]): ujson.Value =
      scope.fold[ujson.Value](ujson.Null)(ujson.Str.apply)
    def nameSnapshotJson(snap: NameSnapshot): ujson.Obj =
      ujson.Obj("formula" -> ujson.Str(snap.formula), "hidden" -> ujson.Bool(snap.hidden))
    def nameEntryJson(n: NameEntry): ujson.Obj =
      ujson.Obj(
        "name" -> ujson.Str(n.name),
        "scope" -> scopeJson(n.scope),
        "formula" -> ujson.Str(n.snapshot.formula),
        "hidden" -> ujson.Bool(n.snapshot.hidden)
      )

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
            "styleChanged" -> ujson.Bool(c.styleChanged),
            "kind" -> ujson.Str(c.kind.name)
          )
        }),
        "mergesAdded" -> strArr(sd.mergesAdded),
        "mergesRemoved" -> strArr(sd.mergesRemoved),
        "commentsAdded" -> strArr(sd.commentsAdded),
        "commentsRemoved" -> strArr(sd.commentsRemoved),
        "commentsChanged" -> strArr(sd.commentsChanged),
        "hyperlinksAdded" -> strArr(sd.hyperlinksAdded),
        "hyperlinksRemoved" -> strArr(sd.hyperlinksRemoved),
        "hyperlinksChanged" -> strArr(sd.hyperlinksChanged),
        "rows" -> axisJson(sd.rows),
        "columns" -> axisJson(sd.columns),
        "properties" -> ujson.Arr.from(sd.properties.map(changeJson)),
        "conditionalFormatsAdded" -> strArr(sd.conditionalFormatsAdded),
        "conditionalFormatsRemoved" -> strArr(sd.conditionalFormatsRemoved),
        "conditionalFormatsChanged" -> strArr(sd.conditionalFormatsChanged),
        "dataValidationsAdded" -> strArr(sd.dataValidationsAdded),
        "dataValidationsRemoved" -> strArr(sd.dataValidationsRemoved),
        "dataValidationsChanged" -> strArr(sd.dataValidationsChanged)
      )
    })

    val root = ujson.Obj(
      "identical" -> ujson.Bool(diff.identical),
      "sheetsAdded" -> strArr(diff.sheetsAdded),
      "sheetsRemoved" -> strArr(diff.sheetsRemoved),
      "sheetOrder" -> diff.sheetOrder.fold[ujson.Value](ujson.Null) { order =>
        ujson.Obj("before" -> strArr(order.before), "after" -> strArr(order.after))
      },
      "namesAdded" -> ujson.Arr.from(diff.namesAdded.map(nameEntryJson)),
      "namesRemoved" -> ujson.Arr.from(diff.namesRemoved.map(nameEntryJson)),
      "namesChanged" -> ujson.Arr.from(diff.namesChanged.map { c =>
        ujson.Obj(
          "name" -> ujson.Str(c.name),
          "scope" -> scopeJson(c.scope),
          "before" -> nameSnapshotJson(c.before),
          "after" -> nameSnapshotJson(c.after)
        )
      }),
      "sheets" -> sheetsJson
    )
    ujson.write(root, indent = 2)
