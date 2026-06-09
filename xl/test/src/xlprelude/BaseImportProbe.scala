// Regression probe (issue #252): extension methods on opaque types must resolve and apply on
// receivers typed via the package-level exported ALIASES (com.tjclp.xl.ARef etc.), from OUTSIDE
// com.tjclp.xl — exactly how scripts see them. Compile success is the test.
package xlprelude

object BaseImportProbe:
  import com.tjclp.xl.{*, given}

  // Alias-typed receivers (the case that breaks with inline extensions)
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

  // Companion factories through the package-level forwarders
  val constructed: ARef = ARef(column, theRow)
  val fromIdx: ARef = ARef.from0(2, 2)
  val parsed: Either[String, ARef] = ARef.parse("C3")

  // SheetName + style units through aliases
  val sn: Either[String, SheetName] = SheetName("Data")
  val snValue: Option[String] = sn.toOption.map(_.value)
  val pt: Pt = Pt(12.0)
  val px: Px = pt.toPx
  val back: Double = px.value
