package com.tjclp.xl.cli.read

import cats.effect.IO
import fs2.Stream
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.RowData
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.style.WorkbookStyles

/**
 * The streaming source's dense projection of a window: the rows the reader never emitted are filled
 * with empty records, and the window's last row ends the pull — a row past it is never a reason to
 * read on.
 */
class StreamingSourceSpec extends CatsEffectSuite:

  private val sheet = SheetName.unsafe("S")

  test("dense: gaps are filled, a row past the window ends the pull, the tail is never read") {
    val window = CellRange(ref"A1", ref"B3")
    val streamed: Stream[IO, RowData] =
      Stream.emits(
        Vector(
          RowData(1, Map(0 -> CellValue.Number(1))),
          RowData(3, Map(1 -> CellValue.Text("x"))),
          RowData(7, Map(0 -> CellValue.Number(7)))
        )
      ) ++ Stream.raiseError[IO](new IllegalStateException("pulled past the window"))
    StreamingSource.dense(streamed, sheet, window, WorkbookStyles.default).compile.toVector.map {
      rows =>
        assertEquals(
          rows.map(_.map(_.ref.toA1)),
          Vector(Vector("A1", "B1"), Vector("A2", "B2"), Vector("A3", "B3"))
        )
        assertEquals(
          rows.flatten.map(_.value),
          Vector(
            CellValue.Number(1),
            CellValue.Empty,
            CellValue.Empty,
            CellValue.Empty,
            CellValue.Empty,
            CellValue.Text("x")
          )
        )
    }
  }

  test("dense: a window the reader leaves entirely empty is every row of it, empty") {
    val window = CellRange(ref"C2", ref"D3")
    StreamingSource
      .dense(Stream.empty, sheet, window, WorkbookStyles.default)
      .compile
      .toVector
      .map { rows =>
        assertEquals(rows.map(_.map(_.ref.toA1)), Vector(Vector("C2", "D2"), Vector("C3", "D3")))
        assertEquals(rows.flatten.map(_.value), Vector.fill(4)(CellValue.Empty: CellValue))
      }
  }
