package com.tjclp.xl.cli.contract

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * GH-596: `sort` places blank keys LAST in both directions — Excel excludes blanks from the
 * ordering and appends them, so `--desc` reverses the values, not the blanks. `gaps.xlsx` has
 * `Data!A1 = 1`, `A2` blank, `A3 = 3`.
 */
class SortContractSpec extends CatsEffectSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "sort-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString

  private def columnA(path: String): IO[Vector[CellValue]] =
    ExcelIO.instance[IO].read(java.nio.file.Paths.get(path)).map { wb =>
      val data = wb.sheets.find(_.name.value == "Data").getOrElse(fail("no Data sheet"))
      Vector(ref"A1", ref"A2", ref"A3").map(r => data(r).value)
    }

  test("sort --desc keeps the blank key last (3, 1, blank), ascending too (1, 3, blank)") {
    val desc = file("sort-desc.xlsx")
    val asc = file("sort-asc.xlsx")
    for
      d <- CliHarness.run(
        "-f",
        file("gaps.xlsx"),
        "-s",
        "Data",
        "-o",
        desc,
        "sort",
        "A1:A3",
        "--by",
        "A",
        "--desc"
      )
      a <- CliHarness.run(
        "-f",
        file("gaps.xlsx"),
        "-s",
        "Data",
        "-o",
        asc,
        "sort",
        "A1:A3",
        "--by",
        "A"
      )
      descA <- columnA(desc)
      ascA <- columnA(asc)
    yield
      assertEquals(d.exit, 0, d.stderr)
      assertEquals(a.exit, 0, a.stderr)
      assertEquals(
        descA,
        Vector(CellValue.Number(BigDecimal(3)), CellValue.Number(BigDecimal(1)), CellValue.Empty)
      )
      assertEquals(
        ascA,
        Vector(CellValue.Number(BigDecimal(1)), CellValue.Number(BigDecimal(3)), CellValue.Empty)
      )
  }
