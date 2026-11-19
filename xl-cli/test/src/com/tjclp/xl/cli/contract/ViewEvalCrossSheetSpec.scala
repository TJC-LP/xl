package com.tjclp.xl.cli.contract

import java.nio.file.Files

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * #891: `view --eval` computes the window's precedents on other sheets. After `put --no-recalc`
 * every cache of the book is stale; a window whose formula reads another sheet used to compute off
 * that sheet's stale cache and show a value no recalculation gives.
 */
class ViewEvalCrossSheetSpec extends CatsEffectSuite:

  private def cached(text: String, value: Int): CellValue =
    CellValue.Formula(text, Some(CellValue.Number(BigDecimal(value))))

  /** Settled at `S!A1 = 1`: `S!F1 = A1*2`, `T!B1 = S!A1+S!F1`, `S!H1 = T!B1*10`. */
  private def book(): Workbook =
    Workbook(
      Vector(
        Sheet("S")
          .put(ref"A1", 1)
          .put(ref"F1", cached("A1*2", 2))
          .put(ref"H1", cached("T!B1*10", 30)),
        Sheet("T").put(ref"B1", cached("S!A1+S!F1", 3))
      )
    )

  private val fixtures = ResourceSuiteLocalFixture(
    "view-eval-cross-sheet",
    Resource.make(
      IO.blocking(Files.createTempDirectory("xl-view-eval-cross-")).flatTap { dir =>
        ExcelIO.instance[IO].write(book(), dir.resolve("chain.xlsx"))
      }
    )(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def h1(json: String): ujson.Value =
    ujson
      .read(json)("data")("rows")
      .arr
      .flatMap(_("cells").arr)
      .find(_("ref").str == "H1")
      .map(_("value"))
      .getOrElse(fail(json))

  test("after `put --no-recalc`, `view --eval` shows what `recalc` then writes") {
    val source = fixtures().resolve("chain.xlsx").toString
    val edited = fixtures().resolve("edited.xlsx").toString
    val settled = fixtures().resolve("settled.xlsx").toString
    for
      put <- CliHarness.run("-f", source, "-s", "S", "-o", edited, "put", "A1", "5", "--no-recalc")
      stale <- CliHarness.run("-f", edited, "-s", "S", "--json", "view", "A1:H1")
      live <- CliHarness.run("-f", edited, "-s", "S", "--json", "view", "A1:H1", "--eval")
      recalc <- CliHarness.run("-f", edited, "-o", settled, "recalc")
      after <- CliHarness.run("-f", settled, "-s", "S", "--json", "view", "A1:H1")
    yield
      assertEquals(put.exit, 0, put.stderr)
      assertEquals(live.exit, 0, live.stderr)
      assertEquals(recalc.exit, 0, recalc.stderr)
      // the file still holds the cache written when A1 was 1
      assertEquals(h1(stale.stdout), ujson.Num(30))
      // T!B1 = 5 + 10 is computed, not read off its cache of 3
      assertEquals(h1(live.stdout), ujson.Num(150))
      assertEquals(h1(after.stdout), h1(live.stdout))
      assert(!live.stdout.contains("EVAL_FAILED"), live.stdout)
  }
