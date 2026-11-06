package com.tjclp.xl.formula

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.eval.{ArrayArithmetic, ArrayResult}
import munit.FunSuite

class LazyReplayBudgetSpec extends FunSuite:
  test("GH-735: repeated operands have linear total leaf reads") {
    for views <- List(false, true) do
      var reads = 0
      val rows = 4
      val levels = 8
      val leaf = ArrayResult.lazyDense(rows, 1) { (r, _) =>
        reads += 1
        CellValue.Number(BigDecimal(r + 1))
      }
      val result = (1 to levels).foldLeft(leaf) { (previous, _) =>
        val operand =
          if views then previous.transpose.transpose.slice(0, rows, 0, 1).toOption.get
          else previous
        ArrayArithmetic
          .broadcastN(Vector(operand, operand), pureFill = true) { values =>
            values match
              case Vector(CellValue.Number(a), CellValue.Number(b)) =>
                Right(CellValue.Number(a + b))
              case _ => Right(CellValue.Empty)
          }
          .toOption
          .get
      }
      assertEquals(result(0, 0), CellValue.Number(BigDecimal(256)))
      assertEquals(result(rows - 1, 0), CellValue.Number(BigDecimal(1024)))
      assert(reads <= rows * levels * 8, s"views=$views, leaf reads=$reads")
  }
