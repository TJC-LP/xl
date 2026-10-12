package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError}

/** #750: the ArrayResult API every consumer reads through (storage-independent). */
class ArrayResultSpec extends FunSuite:

  private def n(i: Int): CellValue = CellValue.Number(BigDecimal(i))

  private val grid = ArrayResult(Vector(Vector(n(1), n(2), n(3)), Vector(n(4), n(5), n(6))))

  test("shape, cellCount and out-of-bounds reads") {
    assertEquals((grid.rows, grid.cols, grid.cellCount), (2, 3, 6L))
    assertEquals(grid(1, 2), n(6))
    assertEquals(grid(2, 0), CellValue.Empty)
    assertEquals(grid(-1, 0), CellValue.Empty)
    assertEquals(grid.fill, ArrayResult.Fill.Const(CellValue.Empty))
    assertEquals((ArrayResult.empty.rows, ArrayResult.empty.cols), (0, 0))
    assert(ArrayResult.empty.isEmpty)
  }

  test("runs and cells walk row-major") {
    assertEquals(grid.cells.toVector, Vector(n(1), n(2), n(3), n(4), n(5), n(6)))
    assertEquals(grid.runs.map(_._2).sum, grid.cellCount)
    assertEquals(
      grid.runs.flatMap((v, k) => Iterator.fill(k.toInt)(v)).toVector,
      grid.cells.toVector
    )
  }

  test("map is row-major and stops at the first error") {
    val seen = Vector.newBuilder[CellValue]
    val out = grid.map { cv =>
      seen += cv
      if cv == n(2) || cv == n(4) then Left(EvalError.EvalFailed(s"bad $cv", None))
      else Right(cv)
    }
    assertEquals(out, Left(EvalError.EvalFailed(s"bad ${n(2)}", None)))
    assertEquals(seen.result(), Vector(n(1), n(2)))
    assertEquals(grid.map(Right(_)), Right(grid))
    assertEquals(grid.mapPure(_ => CellValue.Bool(true)).cells.toSet, Set(CellValue.Bool(true)))
  }

  test("slice reads a window, Empty past the edge") {
    assertEquals(grid.slice(0, 2, 1, 1), ArrayResult(Vector(Vector(n(2)), Vector(n(5)))))
    assertEquals(
      grid.slice(1, 2, 2, 2),
      ArrayResult(
        Vector(Vector(n(6), CellValue.Empty), Vector(CellValue.Empty, CellValue.Empty))
      )
    )
    assertEquals(grid.slice(0, 0, 0, 3), ArrayResult.empty)
  }

  test("transpose and toDense") {
    assertEquals(grid.transpose, ArrayResult.fromFlat(Vector(1, 4, 2, 5, 3, 6).map(n), 3, 2))
    assertEquals(
      grid.toDense("T"),
      Right(Vector(Vector(n(1), n(2), n(3)), Vector(n(4), n(5), n(6))))
    )
  }

  test("equality is cell by cell; hashCode agrees") {
    val same = ArrayResult.fromFlat(Vector(1, 2, 3, 4, 5, 6).map(n), 2, 3)
    assertEquals(grid, same)
    assertEquals(grid.hashCode, same.hashCode)
    assertNotEquals(grid, grid.transpose)
    assertNotEquals(grid, ArrayResult.single(CellValue.Error(CellError.NA)))
    assertNotEquals(ArrayResult.empty, ArrayResult(Vector(Vector.empty)))
  }
