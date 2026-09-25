package com.tjclp.xl.cli.commands

import cats.effect.IO
import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.{Depth, Direction}
import com.tjclp.xl.cli.contract.{CliError, CliException, CliSignal, ErrorCode, OutputMode, Payload}
import com.tjclp.xl.cli.helpers.{Resolve, SheetResolver}
import com.tjclp.xl.cli.output.RendererCommon
import com.tjclp.xl.ooxml.metadata.LightMetadata
import com.tjclp.xl.sheets.FreezePane
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.workbooks.{CalcMode, CalcPr, DefinedName}

/**
 * The inspection verbs of ADR-017 §2.10 — `describe`, `audit`, `deps` — each a typed
 * [[contract.Payload.Json]] under `--json` and readable text otherwise. The analyses live in
 * xl-evaluator ([[WorkbookSummary]], [[WorkbookAudit]], [[QualifiedGraph]]) so scripts get the same
 * answers; this object only resolves arguments and renders.
 *
 * Every listing is deterministic: cells in workbook order (sheet position, row, column) for the
 * audit, [[QualifiedGraph.nodeOrder]] for dependency layers, sheets in workbook order.
 */
object InspectCommands:

  // ===========================================================================================
  // describe
  // ===========================================================================================

  /**
   * `describe` without `--full`: the orientation card from [[LightMetadata]] alone — sheets with
   * state and dimension, defined names with their `hidden` flag, the date system — so it is instant
   * for any file size and identical under `--stream`.
   */
  def describeLight(meta: LightMetadata, mode: OutputMode): Payload =
    val scope: Int => Option[String] = idx => meta.sheets.lift(idx).map(_.name.value)
    mode match
      case OutputMode.Json =>
        Payload.Json(
          ujson.Obj(
            "sheets" -> ujson.Arr.from(meta.sheets.zipWithIndex.map { (info, i) =>
              sheetHeader(info.name.value, i + 1, info.state, info.dimension)
            }),
            "definedNames" -> namesJson(meta.definedNames, scope),
            "date1904" -> ujson.Bool(meta.date1904)
          )
        )
      case OutputMode.Text =>
        val lines = meta.sheets.zipWithIndex.map { (info, i) =>
          SheetLine(i + 1, info.name.value, info.state, info.dimension, None)
        }
        Payload.text(describeText(lines, meta.definedNames, scope, meta.date1904, None))

  /**
   * `describe --full`: the loaded book's [[WorkbookSummary]] — the light card plus every count. Its
   * defined names are the book's effective table (`Workbook.effectiveDefinedNames`, GH-674): the
   * read lifted each sheet's modelable `_xlnm.Print_Area` / `_xlnm.Print_Titles` out of
   * `metadata.definedNames` into its PageSetup (GH-259), and the summary re-derives them, so
   * `--full` lists the sheet-scoped names the light card and `names` read verbatim from
   * workbook.xml (GH-667).
   */
  def describe(wb: Workbook, mode: OutputMode): Payload =
    val summary = WorkbookSummary.of(wb)
    val definedNames = summary.definedNames
    val scope: Int => Option[String] = idx => wb.sheets.lift(idx).map(_.name.value)
    mode match
      case OutputMode.Json =>
        Payload.Json(
          ujson.Obj(
            "sheets" -> ujson.Arr.from(summary.sheets.map(sheetJson)),
            "definedNames" -> namesJson(definedNames, scope),
            "date1904" -> ujson.Bool(summary.date1904),
            "calcPr" -> calcPrJson(summary.calcPr)
          )
        )
      case OutputMode.Text =>
        val lines = summary.sheets.map { s =>
          SheetLine(s.index, s.name.value, s.state, s.dimension, Some(facets(s)))
        }
        Payload.text(
          describeText(lines, definedNames, scope, summary.date1904, Some(summary.calcPr))
        )

  /** One sheet as text mode prints it; `facets` is the `--full` count line. */
  private final case class SheetLine(
    index: Int,
    name: String,
    state: Option[String],
    dimension: Option[CellRange],
    facets: Option[String]
  )

  private def describeText(
    sheets: Vector[SheetLine],
    names: Vector[DefinedName],
    scope: Int => Option[String],
    date1904: Boolean,
    calcPr: Option[Option[CalcPr]]
  ): String =
    val sb = new StringBuilder
    sb.append(s"Sheets (${sheets.size}):")
    if sheets.nonEmpty then
      val nameWidth = sheets.foldLeft(0)((w, l) => math.max(w, l.name.length))
      val dimWidth =
        sheets.foldLeft(0)((w, l) => math.max(w, dimensionText(l.dimension).length))
      sheets.foreach { line =>
        val state = line.state match
          case Some("hidden") => "  hidden"
          case Some("veryHidden") => "  very hidden"
          case _ => ""
        sb.append("\n  ")
          .append(line.index)
          .append("  ")
          .append(line.name.padTo(nameWidth, ' '))
          .append("  ")
          .append(dimensionText(line.dimension).padTo(dimWidth, ' '))
          .append(state)
        line.facets.foreach(text => sb.append("\n       ").append(text))
      }
    sb.append('\n')
    if names.isEmpty then sb.append("Defined names: none")
    else
      sb.append(s"Defined names (${names.size}):")
      val nameWidth = names.foldLeft(0)((w, dn) => math.max(w, dn.name.length))
      names.foreach { dn =>
        val scoped = dn.localSheetId.flatMap(scope).fold("")(s => s"  ($s)")
        val hidden = if dn.hidden then "  (hidden)" else ""
        sb.append("\n  ")
          .append(dn.name.padTo(nameWidth, ' '))
          .append("  ")
          .append(dn.formula)
          .append(scoped)
          .append(hidden)
      }
    sb.append('\n')
    sb.append(s"Date system: ${if date1904 then "1904" else "1900"}")
    calcPr.foreach(cp => sb.append('\n').append(s"Calculation: ${calcPrText(cp)}"))
    sb.toString

  private def dimensionText(dimension: Option[CellRange]): String =
    dimension.fold("(empty)")(_.toA1)

  /** The `--full` count line: the always-present counts, then only the facets the sheet has. */
  private def facets(s: SheetSummary): String =
    val counted = Vector(
      Some(s"cells ${s.cellCount}"),
      Some(
        if s.uncachedFormulas > 0 then
          s"formulas ${s.formulaCount} (${s.uncachedFormulas} uncached)"
        else s"formulas ${s.formulaCount}"
      ),
      Option.when(s.mergedRanges > 0)(s"merged ${s.mergedRanges}"),
      Option.when(s.comments > 0)(s"comments ${s.comments}"),
      Option.when(s.hyperlinks > 0)(s"hyperlinks ${s.hyperlinks}"),
      s.freeze.collect { case FreezePane.At(at, _) => s"freeze ${at.toA1}" },
      s.tabColor.map(c => s"tab ${colorText(c)}"),
      s.autoFilter.map(r => s"autofilter ${r.toA1}"),
      Option.when(s.tables.nonEmpty)(s"tables ${s.tables.mkString(",")}"),
      Option.when(s.charts > 0)(s"charts ${s.charts}"),
      Option.when(s.pictures > 0)(s"pictures ${s.pictures}"),
      Option.when(s.conditionalFormats > 0)(s"cf ${s.conditionalFormats}"),
      Option.when(s.dataValidations > 0)(s"dv ${s.dataValidations}"),
      Option.when(s.hiddenRows > 0)(s"hidden rows ${s.hiddenRows}"),
      Option.when(s.hiddenCols > 0)(s"hidden cols ${s.hiddenCols}")
    )
    counted.flatten.mkString(", ")

  /**
   * `{name, index, state, dimension}` — one element of `describe --json`'s `data.sheets` and of
   * `sheets --json`'s (GH-618: [[WorkbookCommands]] builds its listing through this function, so
   * the two verbs cannot drift apart).
   */
  private[commands] def sheetHeader(
    name: String,
    index: Int,
    state: Option[String],
    dimension: Option[CellRange]
  ): ujson.Obj =
    ujson.Obj(
      "name" -> ujson.Str(name),
      "index" -> ujson.Num(index),
      "state" -> ujson.Str(state.getOrElse("visible")),
      "dimension" -> dimension.fold[ujson.Value](ujson.Null)(r => ujson.Str(r.toA1))
    )

  private def sheetJson(s: SheetSummary): ujson.Obj =
    val obj = sheetHeader(s.name.value, s.index, s.state, s.dimension)
    obj("cells") = ujson.Num(s.cellCount)
    obj("formulas") = ujson.Num(s.formulaCount)
    obj("uncachedFormulas") = ujson.Num(s.uncachedFormulas)
    obj("mergedRanges") = ujson.Num(s.mergedRanges)
    obj("comments") = ujson.Num(s.comments)
    obj("hyperlinks") = ujson.Num(s.hyperlinks)
    obj("freeze") = s.freeze
      .collect { case FreezePane.At(at, _) => at }
      .fold[ujson.Value](ujson.Null)(at => ujson.Str(at.toA1))
    obj("tabColor") = s.tabColor.fold[ujson.Value](ujson.Null)(c => ujson.Str(colorText(c)))
    obj("autoFilter") = s.autoFilter.fold[ujson.Value](ujson.Null)(r => ujson.Str(r.toA1))
    obj("tables") = ujson.Arr.from(s.tables.map(ujson.Str.apply))
    obj("charts") = ujson.Num(s.charts)
    obj("pictures") = ujson.Num(s.pictures)
    obj("conditionalFormats") = ujson.Num(s.conditionalFormats)
    obj("dataValidations") = ujson.Num(s.dataValidations)
    obj("hiddenRows") = ujson.Num(s.hiddenRows)
    obj("hiddenCols") = ujson.Num(s.hiddenCols)
    obj

  /**
   * `[{name, refersTo, scope, hidden}]`, hidden names included — `describe --json`'s
   * `data.definedNames` and the array `names --json` keys as `data.names` (GH-618).
   */
  private[commands] def namesJson(
    names: Vector[DefinedName],
    scope: Int => Option[String]
  ): ujson.Arr =
    ujson.Arr.from(names.map { dn =>
      ujson.Obj(
        "name" -> ujson.Str(dn.name),
        "refersTo" -> ujson.Str(dn.formula),
        "scope" -> dn.localSheetId
          .fold[ujson.Value](ujson.Null)(idx => ujson.Str(scope(idx).getOrElse(s"sheet $idx"))),
        "hidden" -> ujson.Bool(dn.hidden)
      )
    })

  private def colorText(color: Color): String = color match
    case rgb: Color.Rgb => rgb.toHex
    case Color.Theme(slot, tint) =>
      s"theme:${slot.toString.toLowerCase}${if tint == 0.0 then "" else s"/$tint"}"

  private def calcModeText(mode: CalcMode): String = mode match
    case CalcMode.Manual => "manual"
    case CalcMode.Auto => "auto"
    case CalcMode.AutoNoTable => "autoNoTable"

  private def calcPrJson(calcPr: Option[CalcPr]): ujson.Value =
    calcPr.fold[ujson.Value](ujson.Null) { cp =>
      ujson.Obj(
        "iterativeCalculation" -> ujson.Bool(cp.iterativeCalculation),
        "maxIterations" -> cp.maxIterations.fold[ujson.Value](ujson.Null)(n => ujson.Num(n)),
        "maxChange" -> cp.maxChange.fold[ujson.Value](ujson.Null)(d => ujson.Num(d.toDouble)),
        "calcMode" -> cp.calcMode.fold[ujson.Value](ujson.Null)(m => ujson.Str(calcModeText(m))),
        "fullCalcOnLoad" -> cp.fullCalcOnLoad.fold[ujson.Value](ujson.Null)(ujson.Bool.apply),
        "calcId" -> cp.calcId.fold[ujson.Value](ujson.Null)(n => ujson.Num(n))
      )
    }

  /** `default` for a book without modeled calcPr settings, else the settings it carries. */
  private def calcPrText(calcPr: Option[CalcPr]): String =
    calcPr.fold("default") { cp =>
      val iterative = Option.when(cp.iterativeCalculation) {
        val detail =
          cp.maxIterations.map(n => s"$n iterations").toList ++
            cp.maxChange.map(d => s"max change $d").toList
        if detail.isEmpty then "iterative" else detail.mkString("iterative (", ", ", ")")
      }
      val mode = cp.calcMode.map(m => s"calcMode ${calcModeText(m)}")
      val fullCalc = cp.fullCalcOnLoad.filter(identity).map(_ => "full calc on load")
      val calcId = cp.calcId.map(id => s"calcId $id")
      val parts = iterative.toList ++ mode.toList ++ fullCalc.toList ++ calcId.toList
      if parts.isEmpty then "default" else parts.mkString(", ")
    }

  /**
   * `describe` is a workbook verb: `-s` selects nothing. An explicit `-s` that names no sheet is
   * still the caller's mistake, refused as `SHEET_NOT_FOUND` with candidates (exit 3) the way every
   * sheet-taking verb refuses it; an existing sheet is accepted and ignored. `readMeta` runs only
   * when the flag is present (a metadata read: instant for any file size).
   */
  def checkSheetFlag(sheetFlag: Option[String], verb: String)(
    readMeta: IO[LightMetadata]
  ): IO[Unit] =
    sheetFlag.fold(IO.unit) { _ =>
      readMeta.flatMap { meta =>
        IO.fromEither(Resolve.sheetName(meta, sheetFlag, None, verb).left.map(CliException(_))).void
      }
    }

  // ===========================================================================================
  // audit
  // ===========================================================================================

  /**
   * `audit [-s S] [--fail-on-findings]`: the [[WorkbookAudit]] buckets, restricted to `-s` when
   * given. With `--fail-on-findings` a dirty book is the `AUDIT_FINDINGS` signal (exit 1) whose
   * report is the payload — text mode prints the report and no `Error:` line, `--json` keeps it as
   * `data`; without the flag the same report exits 0.
   */
  def audit(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    failOnFindings: Boolean,
    mode: OutputMode
  ): IO[Payload] =
    val whole = WorkbookAudit.of(wb)
    val audit = sheetOpt.fold(whole)(sheet => whole.restrictTo(sheet.name))
    val payload = mode match
      case OutputMode.Json => Payload.Json(auditJson(audit))
      case OutputMode.Text => Payload.text(auditText(audit))
    if failOnFindings && !audit.isClean then
      IO.raiseError(
        CliSignal(
          payload,
          CliError(
            ErrorCode.AUDIT_FINDINGS,
            findingsHeadline(audit) + sheetOpt.fold("")(s => s" on ${s.name.value}"),
            hint = Some(
              "recalculate (`xl recalc`) or fix the listed cells; omit --fail-on-findings to report without failing"
            )
          )
        )
      )
    else IO.pure(payload)

  private def findingsHeadline(audit: WorkbookAudit): String =
    if audit.isClean then "Audit: clean"
    else if audit.findings == 1 then "Audit: 1 finding"
    else s"Audit: ${audit.findings} findings"

  private def refs(qs: Iterable[QualifiedRef]): ujson.Arr =
    ujson.Arr.from(qs.iterator.map(q => ujson.Str(q.toString)))

  private def auditJson(audit: WorkbookAudit): ujson.Obj =
    ujson.Obj(
      "clean" -> ujson.Bool(audit.isClean),
      "findings" -> ujson.Num(audit.findings),
      "errorCells" -> ujson.Arr.from(audit.errorCells.map { (q, e) =>
        ujson.Obj("ref" -> ujson.Str(q.toString), "error" -> ujson.Str(e.toExcel))
      }),
      "uncachedFormulas" -> refs(audit.uncachedFormulas),
      "unparseable" -> ujson.Arr.from(audit.unparseable.map { (q, message) =>
        ujson.Obj("ref" -> ujson.Str(q.toString), "message" -> ujson.Str(message))
      }),
      "volatile" -> refs(audit.volatile),
      "dynamic" -> refs(audit.dynamic),
      "cycles" -> ujson.Arr.from(audit.cycles.map(scc => refs(scc.members))),
      "externalRefs" -> refs(audit.externalRefs),
      "unresolvedReaders" -> refs(audit.unresolvedReaders),
      "calcPr" -> calcPrJson(audit.calcPr),
      // a note, not a finding: the book declares iterative calculation, so its cycles are intended
      "iterativeCycles" -> ujson.Arr.from(audit.iterativeCycles.map(scc => refs(scc.members)))
    )

  /** The headline, then one section per non-empty bucket: findings first, notes after. */
  private def auditText(audit: WorkbookAudit): String =
    def section(title: String, entries: Vector[Vector[String]]): Vector[String] =
      if entries.isEmpty then Vector.empty
      else s"$title (${entries.size}):" +: entries.flatten.map(line => s"  $line")
    def one(text: String): Vector[String] = Vector(text)
    // #676: one line per cell — the formula (capped) and the parser's reason, as lint prints it
    val unparseable = audit.unparseable.map((q, message) => one(s"$q  $message"))
    val sections =
      section("Error values", audit.errorCells.map((q, e) => one(s"$q  ${e.toExcel}"))) ++
        section("Uncached formulas", audit.uncachedFormulas.map(q => one(q.toString))) ++
        section("Unparseable formulas", unparseable) ++
        section("Cycles", audit.cycles.map(scc => one(scc.members.mkString(", ")))) ++
        section("Unresolved names", audit.unresolvedReaders.map(q => one(q.toString))) ++
        section(
          "Cycles (iterative calculation on, not a finding)",
          audit.iterativeCycles.map(scc => one(scc.members.mkString(", ")))
        ) ++
        section("Volatile", audit.volatile.map(q => one(q.toString))) ++
        section("Dynamic", audit.dynamic.map(q => one(q.toString))) ++
        section("External references", audit.externalRefs.map(q => one(q.toString))) ++
        audit.calcPr.toList.map(cp => s"Calculation: ${calcPrText(Some(cp))}").toVector
    (findingsHeadline(audit) +: sections).mkString("\n")

  // ===========================================================================================
  // deps
  // ===========================================================================================

  /**
   * `deps <ref> [--direction precedents|dependents|both] [--depth N|all] [--expand]`: the cell,
   * then its precedent and/or dependent layers from [[QualifiedGraph]] — each node with its depth,
   * formula and value. A precedent range is ONE node carrying its occupied-cell and formula counts
   * ([[QualifiedGraph.declaredPrecedents]]), as Excel's Trace Precedents draws it; `--expand` lists
   * its occupied cells one by one instead ([[QualifiedGraph.precedents]]). Dependents are formula
   * cells either way. The ref resolves through the one sheet rule ([[SheetResolver.resolveRef]]);
   * the parser has already turned the flags into [[Direction]] and [[Depth]] (`Depth.Hops(1)` when
   * `--depth` is absent).
   */
  def deps(
    wb: Workbook,
    sheetOpt: Option[Sheet],
    refStr: String,
    direction: Direction,
    depth: Depth,
    expand: Boolean,
    mode: OutputMode
  ): IO[Payload] =
    for
      resolved <- SheetResolver.resolveRef(wb, sheetOpt, refStr, "deps")
      (sheet, refOrRange) = resolved
      ref <- refOrRange match
        case Left(r) => IO.pure(r)
        case Right(range) =>
          IO.raiseError(
            CliException(
              CliError.fromXLError(
                XLError.InvalidReference(
                  s"deps requires a single cell, not the range ${range.toA1}"
                ),
                None
              )
            )
          )
    yield
      val graph = QualifiedGraph.of(wb)
      val start = QualifiedRef(sheet.name, ref)
      // QualifiedGraph's walk takes a hop budget where `<= 0` means unbounded
      val bound = depth match
        case Depth.All => 0
        case Depth.Hops(n) => n
      def cells(layers: Vector[Vector[QualifiedRef]]): Vector[Vector[QualifiedGraph.Node]] =
        layers.map(_.map(QualifiedGraph.Node.Cell(_)))
      val precedents = Option.when(direction != Direction.Dependents) {
        if expand then cells(graph.precedents(start, bound))
        else graph.declaredPrecedents(start, bound)
      }
      val dependents =
        Option.when(direction != Direction.Precedents)(cells(graph.dependents(start, bound)))
      mode match
        case OutputMode.Json =>
          Payload.Json(
            depsJson(wb, graph, start, direction, depth, expand, precedents, dependents)
          )
        case OutputMode.Text =>
          Payload.text(depsText(wb, graph, start, depth, precedents, dependents))

  private def valueAt(wb: Workbook, q: QualifiedRef): CellValue =
    wb.sheets.find(_.name == q.sheet).flatMap(_.cells.get(q.ref)).fold(CellValue.Empty)(_.value)

  private def formulaOf(value: CellValue): Option[String] = value match
    case CellValue.Formula(expr, _, kind) => Some(RendererCommon.formulaDisplay(expr, kind))
    case _ => None

  /** The raw value as JSON: a formula's cached value, `null` when it has none. */
  private def valueJson(value: CellValue): ujson.Value = value match
    case CellValue.Text(s) => ujson.Str(s)
    case CellValue.Number(n) => ujson.Num(n.toDouble)
    case CellValue.Bool(b) => ujson.Bool(b)
    case CellValue.DateTime(dt) => ujson.Str(dt.toString)
    case CellValue.Error(err) => ujson.Str(err.toExcel)
    case CellValue.RichText(rt) => ujson.Str(rt.toPlainText)
    case CellValue.Empty => ujson.Null
    case CellValue.Formula(_, cached, _) => cached.fold[ujson.Value](ujson.Null)(valueJson)

  /** The raw value as text: a formula's cached value, `(uncached)` when it has none. */
  private def valueText(value: CellValue): String = value match
    case CellValue.Text(s) => s"\"$s\""
    case CellValue.Number(n) =>
      if n.isWhole then n.toBigInt.toString else n.underlying.stripTrailingZeros.toPlainString
    case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
    case CellValue.DateTime(dt) => dt.toString
    case CellValue.Error(err) => err.toExcel
    case CellValue.RichText(rt) => s"\"${rt.toPlainText}\""
    case CellValue.Empty => "(empty)"
    case CellValue.Formula(_, cached, _) => cached.fold("(uncached)")(valueText)

  /** A range node's counts: its occupied cells, and how many of those are formulas. */
  private def rangeCounts(graph: QualifiedGraph, range: QualifiedGraph.DeclaredRange): (Int, Int) =
    val occupied = graph.occupiedIn(range)
    (occupied.size, occupied.count(graph.formulas.contains))

  // `kind` second, so a reader that switches on it sees it before the fields it gates; a range
  // keeps every cell-node key (`formula`/`value` null) so `.ref`/`.value` readers never miss one
  private def nodeJson(
    wb: Workbook,
    graph: QualifiedGraph,
    node: QualifiedGraph.Node,
    depth: Int
  ): ujson.Obj =
    node match
      case QualifiedGraph.Node.Cell(q) =>
        val value = valueAt(wb, q)
        ujson.Obj(
          "ref" -> ujson.Str(q.toString),
          "kind" -> ujson.Str("cell"),
          "depth" -> ujson.Num(depth),
          "formula" -> formulaOf(value).fold[ujson.Value](ujson.Null)(ujson.Str.apply),
          "value" -> valueJson(value)
        )
      case QualifiedGraph.Node.Range(range) =>
        val (occupied, formulas) = rangeCounts(graph, range)
        ujson.Obj(
          "ref" -> ujson.Str(range.toString),
          "kind" -> ujson.Str("range"),
          "depth" -> ujson.Num(depth),
          "formula" -> ujson.Null,
          "value" -> ujson.Null,
          "occupied" -> ujson.Num(occupied),
          "formulas" -> ujson.Num(formulas)
        )

  private def layersJson(
    wb: Workbook,
    graph: QualifiedGraph,
    layers: Vector[Vector[QualifiedGraph.Node]]
  ): ujson.Arr =
    ujson.Arr.from(layers.zipWithIndex.flatMap { (layer, i) =>
      layer.map(node => nodeJson(wb, graph, node, i + 1))
    })

  private def depthJson(depth: Depth): ujson.Value = depth match
    case Depth.All => ujson.Str("all")
    case Depth.Hops(n) => ujson.Num(n)

  private def depsJson(
    wb: Workbook,
    graph: QualifiedGraph,
    start: QualifiedRef,
    direction: Direction,
    depth: Depth,
    expand: Boolean,
    precedents: Option[Vector[Vector[QualifiedGraph.Node]]],
    dependents: Option[Vector[Vector[QualifiedGraph.Node]]]
  ): ujson.Obj =
    val value = valueAt(wb, start)
    ujson.Obj(
      "ref" -> ujson.Str(start.toString),
      "formula" -> formulaOf(value).fold[ujson.Value](ujson.Null)(ujson.Str.apply),
      "value" -> valueJson(value),
      "direction" -> ujson.Str(direction.flag),
      "depth" -> depthJson(depth),
      "expand" -> ujson.Bool(expand),
      "precedents" -> precedents.fold[ujson.Value](ujson.Null)(layersJson(wb, graph, _)),
      "dependents" -> dependents.fold[ujson.Value](ujson.Null)(layersJson(wb, graph, _))
    )

  private def depsText(
    wb: Workbook,
    graph: QualifiedGraph,
    start: QualifiedRef,
    depth: Depth,
    precedents: Option[Vector[Vector[QualifiedGraph.Node]]],
    dependents: Option[Vector[Vector[QualifiedGraph.Node]]]
  ): String =
    val depthLabel = depth match
      case Depth.All => "depth all"
      case Depth.Hops(n) => s"depth $n"
    // Plain digits, never locale grouping: the line stays deterministic and greppable
    def count(n: Int, noun: String): String = if n == 1 then s"1 $noun" else s"$n ${noun}s"
    def describe(node: QualifiedGraph.Node): String = node match
      case QualifiedGraph.Node.Cell(q) =>
        val value = valueAt(wb, q)
        formulaOf(value).fold(valueText(value))(f => s"$f -> ${valueText(value)}")
      case QualifiedGraph.Node.Range(range) =>
        val (occupied, formulas) = rangeCounts(graph, range)
        val formulaNote = if formulas == 0 then "" else s" (${count(formulas, "formula")})"
        s"range, ${count(occupied, "occupied cell")}$formulaNote"
    def block(title: String, layers: Vector[Vector[QualifiedGraph.Node]]): Vector[String] =
      val nodes = layers.zipWithIndex.flatMap { (layer, i) =>
        layer.map(node => s"  ${i + 1}  ${node.label}  ${describe(node)}")
      }
      s"$title ($depthLabel): ${nodes.size}" +: (if nodes.isEmpty then Vector("  (none)")
                                                 else nodes)
    val value = valueAt(wb, start)
    val head =
      Vector(s"Cell: $start") ++ formulaOf(value).toList.map(f => s"Formula: $f") :+
        s"Value: ${valueText(value)}"
    val body =
      precedents.fold(Vector.empty[String])(block("Precedents", _)) ++
        dependents.fold(Vector.empty[String])(block("Dependents", _))
    (head ++ body).mkString("\n")
