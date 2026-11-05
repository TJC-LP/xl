package xlprelude

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.ArrayResult
import munit.FunSuite

/** #735: existing Vector construction remains valid outside the library package. */
class ArrayFillPreludeProbe extends FunSuite:
  test("axis fill constructors and extractors retain Vector payloads") {
    val values = Vector(CellValue.Number(BigDecimal(3)), CellValue.Number(BigDecimal(4)))
    val byRow = ArrayResult.Fill.ByRow(values)
    val byCol = ArrayResult.Fill.ByCol(values)
    assertEquals(byRow.at(1, 0), values(1))
    assertEquals(byCol.at(0, 1), values(1))
    byRow match
      case ArrayResult.Fill.ByRow(entries) =>
        val vector: Vector[CellValue] = entries
        assertEquals(vector, values)
      case other => fail(s"expected row fill: $other")
  }
