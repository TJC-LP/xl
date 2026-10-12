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
    assertEquals(grid.slice(0, 2, 1, 1), Right(ArrayResult(Vector(Vector(n(2)), Vector(n(5))))))
    assertEquals(
      grid.slice(1, 2, 2, 2),
      Right(
        ArrayResult(
          Vector(Vector(n(6), CellValue.Empty), Vector(CellValue.Empty, CellValue.Empty))
        )
      )
    )
    assertEquals(grid.slice(0, 0, 0, 3), Right(ArrayResult.empty))
  }

  test(
    "#750: a window past the edge of a non-Empty fill is capped; inside the array it stays sparse"
  ) {
    val whole = ArrayResult.constant(1048576, 16384, n(7))
    // inside: sparse, whatever its size
    whole.slice(0, 1048576, 0, 16384) match
      case Right(ar) => assertEquals((ar.rows, ar.cols, ar(5, 5)), (1048576, 16384, n(7)))
      case Left(err) => fail(s"an inside window is never refused: $err")
    // past the edge: two outside values varying along both axes, built cell by cell up to the cap
    assertEquals(
      whole.slice(1048575, 2, 16383, 2),
      Right(ArrayResult(Vector(Vector(n(7), CellValue.Empty), Vector.fill(2)(CellValue.Empty))))
    )
    whole.slice(1, 1048576, 0, 16384) match
      case Left(EvalError.EvalFailed(msg, _)) => assert(msg.contains("1,048,576-cell limit"), msg)
      case other => fail(s"expected the cell-limit refusal, got $other")
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

  test("#750: == on whole-sheet arrays filled by column compares boxes and fills, not cells") {
    // 1,048,576 × 16,384 with a fill per column: a run walk is a piece per cell (1.7e10)
    def wide(entry: Int => Int, corner: Int) = ArrayResult.sparse(
      1048576,
      16384,
      2,
      3,
      Vector(Vector(n(corner), n(0)), Vector(n(1), n(2))),
      ArrayResult.Fill.ByCol(Vector.tabulate(16384)(c => n(entry(c))))
    )
    val start = System.nanoTime()
    assertEquals(wide(_ % 7, 9), wide(_ % 7, 9))
    assertNotEquals(wide(_ % 7, 9), wide(c => if c == 16000 then 99 else c % 7, 9))
    assertNotEquals(wide(_ % 7, 9), wide(_ % 7, 8))
    val elapsedMs = (System.nanoTime() - start) / 1000000L
    assert(elapsedMs < 5000L, s"== took $elapsedMs ms")
    // other geometry, same cells: the run walk decides (a box cell equal to the fill it replaces)
    val boxed = ArrayResult.sparse(3, 2, 0, 0, Vector(Vector(n(5))), ArrayResult.Fill.Const(n(5)))
    assertEquals(boxed, ArrayResult.constant(3, 2, n(5)))
    assertEquals(
      ArrayResult.constant(3, 2, n(5)),
      ArrayResult(Vector.fill(3)(Vector.fill(2)(n(5))))
    )
  }
