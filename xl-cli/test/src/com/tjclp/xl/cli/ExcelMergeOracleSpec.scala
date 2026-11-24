package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import cats.effect.unsafe.implicits.global

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.contract.CliHarness
import com.tjclp.xl.formula.eval.CellMover
import com.tjclp.xl.io.ExcelIO

import munit.FunSuite

/**
 * GH-935: the Excel oracle. `excel-oracle/merge-references.json` (xl-evaluator test resources)
 * holds 93 merge scenarios run in Excel for Mac 16.115 — 477 cell expectations, every defined name
 * and merged area after the merge, and for each of the 69 empty-anchor cases a cut-and-paste of the
 * winner onto the anchor that produced the same formulas character for character. Each scenario's
 * workbook is built here, merged by the workbook-level merge (`CellMover.merge`, what
 * `wb.edit(Edit.Merge(...))`, the `merge` verb and the batch op run), recalculated, and compared:
 * every formula's text exactly (Excel's `=` prefix is xl's stored text with `=` added; no other
 * normalisation), every value, the merged areas and the names.
 *
 * Inputs: a scenario's `sheets` is the workbook as Excel stored it before the merge (ranges
 * normalised, `=SUM(C1:B1)` read back as `=SUM(B1:C1)`), used as is; `submitted` adds the cells
 * Excel does not read back (a value written into a merge's hidden cell). Excel's internal
 * `_xlfn.SINGLE` name (its bookkeeping for `@`) is not part of a scenario and is ignored.
 */
class ExcelMergeOracleSpec extends FunSuite:

  override val munitTimeout = scala.concurrent.duration.Duration(180, "s")

  private val fixture: ujson.Value =
    val stream = Option(getClass.getResourceAsStream("/excel-oracle/merge-references.json"))
      .getOrElse(fail("excel-oracle/merge-references.json is not on the test classpath"))
    try ujson.read(new String(stream.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8))
    finally stream.close()

  private val scenarios: Vector[ujson.Value] = fixture.arr.toVector

  /**
   * Values xl does not yet compute as Excel does, by scenario and cell, with the reason. These are
   * evaluator limits, not merge ones: every formula text, name and merged area of these scenarios
   * still matches and is asserted; only the listed values are checked separately, by a test that is
   * expected to fail until the limit is lifted (so a fix shows up as a failing `.fail` test).
   */
  private val selfReference =
    "xl's dependency graph counts the cell arguments of ROW/COLUMN/ROWS/COLUMNS as value reads, so " +
      "a formula naming its own cell there is reported circular and left uncached (Excel reads only " +
      "the address)"
  private val pendingValues: Map[String, (Set[String], String)] = Map(
    "f4-13" -> (Set("Sheet1!A1"), selfReference),
    "f4-14" -> (Set("Sheet1!A1"), selfReference),
    "f4-15" -> (Set("Sheet1!A1"), selfReference),
    "f4-16" -> (Set("Sheet1!A1"), selfReference)
  )

  private def ref(a1: String): ARef = ARef.parse(a1).fold(e => fail(e), identity)
  private def range(a1: String): CellRange = CellRange.parse(a1).fold(e => fail(e), identity)

  private def input(v: ujson.Value): CellValue = v match
    case ujson.Num(n) => CellValue.Number(BigDecimal(n))
    case ujson.Str(s) if s.startsWith("=") => CellValue.Formula(s.drop(1), None)
    case ujson.Str(s) => CellValue.Text(s)
    case ujson.Bool(b) => CellValue.Bool(b)
    case other => fail(s"unexpected input $other")

  private def sheetNames(sc: ujson.Value): Vector[String] =
    val fromSheets = sc("sheets").obj.keys.toVector
    val fromSubmitted = sc.obj.get("submitted").toList.flatMap(_.obj.keys).toVector
    val fromExpected = sc("expectedAfter").obj.keys.toVector.map(_.takeWhile(_ != '!'))
    val fromMerge = Vector(sc("merge")("sheet").str)
    (Vector("Sheet1") ++ fromSheets ++ fromSubmitted ++ fromExpected ++ fromMerge).distinct

  private def build(sc: ujson.Value): Workbook =
    val names = sheetNames(sc)
    val sheets = names.map { name =>
      val stored = sc("sheets").obj.get(name).map(_.obj.toVector).getOrElse(Vector.empty)
      val submitted = sc.obj
        .get("submitted")
        .flatMap(_.obj.get(name))
        .map(_.obj.toVector)
        .getOrElse(Vector.empty)
        .filterNot((a1, _) => stored.exists(_._1 == a1))
      val merges = sc("mergedAreasBefore").arr.map(_.str).collect {
        case q if q.takeWhile(_ != '!') == name => range(q.dropWhile(_ != '!').drop(1))
      }
      (stored ++ submitted)
        .foldLeft(Sheet(SheetName.unsafe(name)))((s, kv) => s.put(ref(kv._1), input(kv._2)))
        .copy(mergedRanges = merges.toSet)
    }
    sc("names").obj
      .filterNot((n, _) => n.startsWith("_xlfn."))
      .foldLeft(Workbook(sheets))((wb, kv) => wb.withDefinedName(kv._1, kv._2.str.stripPrefix("=")))

  private def merged(sc: ujson.Value): Workbook =
    val wb = build(sc)
    val sheet = SheetName.unsafe(sc("merge")("sheet").str)
    CellMover.merge(wb, sheet, range(sc("merge")("range").str)).fold(e => fail(e.message), identity)

  /** The formula text and the value Excel reports for `cell`. */
  private def actual(wb: Workbook, qualified: String): (String, ujson.Value) =
    val (sheet, a1) = (qualified.takeWhile(_ != '!'), qualified.dropWhile(_ != '!').drop(1))
    val value = wb(SheetName.unsafe(sheet)).toOption.flatMap(_.cells.get(ref(a1))).map(_.value)
    def json(v: CellValue): ujson.Value = v match
      case CellValue.Empty => ujson.Null
      case CellValue.Number(n) => ujson.Num(n.toDouble)
      case CellValue.Text(s) => ujson.Str(s)
      case CellValue.Bool(b) => ujson.Bool(b)
      case CellValue.Error(e) => ujson.Str(e.toExcel)
      case CellValue.Formula(_, cached, _) => cached.fold(ujson.Null)(json)
      case other => ujson.Str(other.toString)
    value match
      case None | Some(CellValue.Empty) => ("", ujson.Null)
      case Some(f @ CellValue.Formula(text, _, _)) => (s"=$text", json(f))
      case Some(CellValue.Number(n)) =>
        (n.bigDecimal.stripTrailingZeros.toPlainString, ujson.Num(n.toDouble))
      case Some(CellValue.Text(s)) => (s, ujson.Str(s))
      case Some(CellValue.Bool(b)) => (if b then "TRUE" else "FALSE", ujson.Bool(b))
      case Some(other) => (other.toString, json(other))

  private def sameValue(expected: ujson.Value, got: ujson.Value): Boolean =
    (expected, got) match
      case (ujson.Num(a), ujson.Num(b)) => math.abs(a - b) < 1e-9
      case (a, b) => a == b

  for sc <- scenarios do
    val id = sc("id").str
    val title = s"oracle $id: ${sc("description").str}"
    val skipValues = pendingValues.get(id).fold(Set.empty[String])(_._1)
    test(title) {
      val after = merged(sc)
      val recalculated = after.recalculate().workbook
      val problems = Vector.newBuilder[String]
      sc("expectedAfter").obj.foreach { (cell, exp) =>
        val (formula, _) = actual(after, cell)
        val (_, value) = actual(recalculated, cell)
        if formula != exp("formula").str then
          problems += s"$cell formula: expected ${exp("formula").str}, got $formula"
        if !skipValues.contains(cell) && !sameValue(exp("value"), value) then
          problems += s"$cell value: expected ${exp("value")}, got $value"
      }
      val names = after.metadata.definedNames
        .filterNot(_.name.startsWith("_xlfn."))
        .map(n => n.name -> s"=${n.formula}")
        .toMap
      val expectedNames =
        sc("namesAfter").obj
          .filterNot((n, _) => n.startsWith("_xlfn."))
          .map((n, v) => n -> v.str)
          .toMap
      if names != expectedNames then problems += s"names: expected $expectedNames, got $names"
      val areas =
        after.sheets.flatMap(s => s.mergedRanges.map(r => s"${s.name.value}!${r.toA1}")).toSet
      val expectedAreas = sc("mergedAreasAfter").arr.map(_.str).toSet
      if areas != expectedAreas then
        problems += s"merged areas: expected $expectedAreas, got $areas"
      val found = problems.result()
      assert(found.isEmpty, found.mkString(s"$id:\n  ", "\n  ", ""))
    }
    pendingValues.get(id).foreach { (cells, reason) =>
      test(s"oracle $id PENDING values ${cells.mkString(", ")}: $reason".fail) {
        val recalculated = merged(sc).recalculate().workbook
        cells.foreach { cell =>
          val (_, value) = actual(recalculated, cell)
          val exp = sc("expectedAfter")(cell)("value")
          assert(sameValue(exp, value), s"$cell value: expected $exp, got $value")
        }
      }
    }

  test("oracle: the fixture holds the 93 scenarios Excel ran") {
    assertEquals(scenarios.size, 93)
    assertEquals(scenarios.map(_("expectedAfter").obj.size).sum, 477)
  }

  // ===== the CLI merge verb over a sample of the scenarios =====

  private def withDir[A](f: Path => A): A =
    val dir = Files.createTempDirectory("excel-merge-oracle")
    try f(dir)
    finally
      Files.list(dir).forEach(p => Files.deleteIfExists(p))
      Files.deleteIfExists(dir)

  for id <- List("f1-08", "f1-09", "f1-23", "f2-04", "f3-02", "f4-02", "f5-08", "f6-09", "f7-01") do
    test(s"oracle via the CLI merge verb: $id") {
      val sc = scenarios.find(_("id").str == id).getOrElse(fail(s"no scenario $id"))
      withDir { dir =>
        val in = dir.resolve("in.xlsx")
        val out = dir.resolve("out.xlsx")
        ExcelIO.instance[IO].write(build(sc), in).unsafeRunSync()
        val run = CliHarness
          .run(
            "-f",
            in.toString,
            "-o",
            out.toString,
            "-s",
            sc("merge")("sheet").str,
            "merge",
            sc("merge")("range").str
          )
          .unsafeRunSync()
        assertEquals(run.exit, 0, run.stderr)
        val wb = ExcelIO.instance[IO].read(out).unsafeRunSync()
        sc("expectedAfter").obj.foreach { (cell, exp) =>
          val (formula, value) = actual(wb, cell)
          assertEquals(formula, exp("formula").str, s"$id $cell")
          assert(
            sameValue(exp("value"), value),
            s"$id $cell value: expected ${exp("value")}, got $value"
          )
        }
      }
    }
