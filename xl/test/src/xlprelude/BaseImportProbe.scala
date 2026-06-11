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
