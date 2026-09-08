// Regression probe (issue #252): the public API surface must work on receivers exactly as
// scripts see them — package-level aliases, factory-result chains, and the prelude — from
// OUTSIDE com.tjclp.xl. Compile success is the test.
package xlprelude

object BaseImportProbe:
  import com.tjclp.xl.{*, given}

  // Alias-typed receivers
  val cell: ARef = ref"A1"
  val a1: String = cell.toA1
  val shifted: ARef = cell.shift(1, 1)
  val shiftedA1: String = shifted.toA1
  val column: Column = cell.col
  val letter: String = column.toLetter
  val colIdx: Int = column.index0
  val colNext: Column = column + 1
  val theRow: Row = cell.row
  val rowIdx: Int = theRow.index1
  val rowNext: Row = theRow + 1

  // Companion factories through the package-level singleton vals
  val constructed: ARef = ARef(column, theRow)
  val fromIdx: ARef = ARef.from0(2, 2)
  val parsed: Either[String, ARef] = ARef.parse("C3")
  val colParsed: Either[String, Column] = Column.parse("D")
  val colFromRef: Either[String, Column] = Column.parse("D1") // trailing row tolerated
  val refTypeCol: Either[String, Column] = RefType.parse("Sales!C2:E9").map(_.col)

  // Extension methods chained directly on factory results (demo.sc regression)
  val chainedLetter: String = Column.from0(0).toLetter
  val chainedA1: String = ARef.from0(2, 2).toA1
  val chainedShift: String = ARef.from0(0, 0).shift(1, 1).toA1
  val chainedIdx: Int = Row.from0(4).index1
  val inferredFactory = Column.from0(0)
  val inferredLetter: String = inferredFactory.toLetter

  // SheetName + style units: aliases, factories, and chains
  val sn: Either[String, SheetName] = SheetName("Data")
  val snValue: Option[String] = sn.toOption.map(_.value)
  val pt: Pt = Pt(12.0)
  val px: Px = pt.toPx
  val back: Double = px.value
  val chainedUnit: Double = Pt(12.0).toPx.value

  // Conditional formatting (GH-136): enums, Dxf builders, and the Sheet authoring surface all
  // resolve through the base import; constructors chain directly off companion factories.
  val cfDxf: Dxf = Dxf.fillAndFont(Color.Rgb(0xffffc7ce), DxfFont(bold = Some(true)))
  val cfRule: CfRule = CfRule.cellIs(CfOperator.GreaterThan, "100", cfDxf)
  val cfPoint: CfPoint = CfPoint(Cfvo.Percentile(BigDecimal(50)), Color.Rgb(0xffffeb84))
  val cfScale: CfRule = CfRule.colorScale2(CfPoint(Cfvo.Min, Color.Rgb(0xffff0000)), cfPoint)
  val cfSheet: Sheet = Sheet(SheetName.unsafe("CF")).conditionalFormat(ref"A1:A9", cfRule, cfScale)
  val cfBlocks: Vector[ConditionalFormat.Rules] = cfSheet.typedConditionalFormats
  val cfText: CfTextOp = CfTextOp.Contains

  // GH-589: the recalculating writes reach the base import (package-level ExcelRecalc export), so
  // a model built here has the cache-completing write and not only Excel.write. Never invoked.
  val checkedWrite: Workbook => RecalcResult = wb => Excel.writeChecked(wb, "/tmp/never-run.xlsx")
  val recalculatedWrite: (Workbook, RecalcOptions) => RecalcResult =
    (wb, opts) => Excel.writeRecalculated(wb, "/tmp/never-run.xlsx", opts)

  // GH-465 / GH-589: outline collapse composes hidden members + the collapsed summary marker
  // (whole-row/column spans are runtime strings — the ref macro takes A1 / A1:B2 shapes only —
  // and ColSpan/RowSpan carry the axis; the CellRange overload is XLResult and refuses the other)
  val colSpan: XLResult[ColSpan] = ColSpan.parse("E:H")
  val colRange: Either[String, CellRange] = CellRange.parse("E:H")
  val collapsed: Sheet =
    cfSheet.collapseRows(Row.from1(2), Row.from1(4)).collapseCols(Column.from0(4), Column.from0(7))
  val reopened: Sheet =
    colSpan
      .fold(_ => collapsed, span => collapsed.expandCols(span))
      .expandRows(Row.from1(2), Row.from1(4))
  val viaRange: XLResult[Sheet] =
    colRange.left.map(XLError.InvalidReference(_)).flatMap(r => collapsed.expandCols(r))
