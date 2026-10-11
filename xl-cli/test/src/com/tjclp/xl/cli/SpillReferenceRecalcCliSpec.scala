package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipOutputStream}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.DynamicArrayFixtures

/**
 * GH-695, end to end through the real command tree on the book Excel writes: A1:A5 hold numbers, B1
 * is a dynamic-array anchor (`<f t="array" ref="B1:B5">_xlfn._xlws.SORT(A1:A5)</f>`, `cm="1"` and
 * metadata.xml) with its spill cached in B2:B5, and C1 = `SUM(_xlfn.ANCHORARRAY(B1))` cached as 15.
 * `eval`, `recalc` and a `put` elsewhere all evaluated B1 before C1 and threaded B1's value back as
 * a constant, so C1 read `B1#` off a non-anchor: `#REF!` overwrote the correct 15.
 */
class SpillReferenceRecalcCliSpec extends CatsEffectSuite:

  // GH-714: the Excel-shaped package lives in the shared fixture builder
  private val parts: Vector[(String, String)] = DynamicArrayFixtures.parts()

  private def excelShapedSource(): IO[Path] = IO.blocking {
    val out = Files.createTempFile("xl-gh695-src-", ".xlsx")
    val zos = new ZipOutputStream(Files.newOutputStream(out))
    try
      parts.foreach { (name, text) =>
        zos.putNextEntry(new ZipEntry(name))
        zos.write(text.getBytes(StandardCharsets.UTF_8))
        zos.closeEntry()
      }
    finally zos.close()
    out
  }

  private def outPath(): IO[Path] = IO.blocking(Files.createTempFile("xl-gh695-out-", ".xlsx"))

  private def c1(path: Path): IO[CellValue] =
    ExcelIO.instance[IO].read(path).map { wb =>
      wb.sheets.headOption.getOrElse(fail("no sheet"))(ref"C1").value
    }

  private def assertC1Is15(run: CliRun, out: Path): IO[Unit] =
    assertEquals(run.exit, 0, run.toString)
    c1(out).map {
      case CellValue.Formula(_, Some(CellValue.Number(v)), _) => assertEquals(v, BigDecimal(15))
      case other => fail(s"C1 should keep its cached 15, got $other")
    }

  test("eval: the unqualified and the sheet-qualified spill reference agree") {
    for
      src <- excelShapedSource()
      local <- CliHarness.run("-f", src.toString, "-s", "Sheet1", "eval", "=SUM(B1#)")
      qualified <- CliHarness.run("-f", src.toString, "-s", "Sheet1", "eval", "=SUM(Sheet1!B1#)")
    yield
      assertEquals(local.exit, 0, local.toString)
      assert(local.stdout.contains("Result: 15 (number)"), local.stdout)
      assert(qualified.stdout.contains("Result: 15 (number)"), qualified.stdout)
  }

  test("recalc keeps C1's correct value") {
    for
      src <- excelShapedSource()
      out <- outPath()
      run <- CliHarness.run("-f", src.toString, "-s", "Sheet1", "-o", out.toString, "recalc")
      _ <- assertC1Is15(run, out)
    yield assert(!run.stdout.contains("error value"), run.stdout)
  }

  test("a put that changes the anchor's input re-spills it for C1") {
    for
      src <- excelShapedSource()
      out <- outPath()
      // A1 5 → 0: SORT(A1:A5) is {0;1;2;3;4}, so Excel's C1 is 10
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-s",
        "Sheet1",
        "-o",
        out.toString,
        "put",
        "A1",
        "0"
      )
      _ = assertEquals(run.exit, 0, run.toString)
      value <- c1(out)
    yield value match
      case CellValue.Formula(_, Some(CellValue.Number(v)), _) => assertEquals(v, BigDecimal(10))
      case other => fail(s"C1 should be 10, got $other")
  }

  test("eval --with re-spills the anchor over the override") {
    for
      src <- excelShapedSource()
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-s",
        "Sheet1",
        "eval",
        "=SUM(B1#)",
        "--with",
        "A1=0"
      )
    yield assert(run.stdout.contains("Result: 10 (number)"), run.toString)
  }

  test("evala spills the anchor's recomputed array") {
    for
      src <- excelShapedSource()
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-s",
        "Sheet1",
        "evala",
        "=B1#*10",
        "--with",
        "A1=0"
      )
    yield
      assertEquals(run.exit, 0, run.toString)
      // SORT over {0;3;1;4;2} is {0;1;2;3;4}: the grid's column, not the stale {10;…;50}
      val column = run.stdout.linesIterator.collect {
        case s"| $row | $value |" if row.trim.nonEmpty && row.trim.forall(_.isDigit) => value.trim
      }.toVector
      assert(run.stdout.contains("Spill Range: Z1000:Z1004 (5×1)"), run.stdout)
      assertEquals(column, Vector("0", "10", "20", "30", "40"), run.stdout)
  }

  test("recalc over unchanged inputs leaves the worksheet XML byte-identical") {
    def sheetXml(path: Path): String =
      val zip = new java.util.zip.ZipFile(path.toFile)
      try
        new String(
          zip.getInputStream(zip.getEntry("xl/worksheets/sheet1.xml")).readAllBytes(),
          StandardCharsets.UTF_8
        )
      finally zip.close()
    for
      src <- excelShapedSource()
      out <- outPath()
      run <- CliHarness.run("-f", src.toString, "-s", "Sheet1", "-o", out.toString, "recalc")
    yield
      assertEquals(run.exit, 0, run.toString)
      // every recomputed cache equals the file's, so B1's record (t="array", ref, cm) and the
      // spill cells are carried verbatim
      assertEquals(sheetXml(out), sheetXml(src))
  }

  test("a function returning the anchor writes the anchor's value as its cache") {
    for
      src <- excelShapedSource()
      out <- outPath()
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-s",
        "Sheet1",
        "-o",
        out.toString,
        "putf",
        "D1",
        "=IFERROR(B1,0)"
      )
      wb <- ExcelIO.instance[IO].read(out)
    yield
      assertEquals(run.exit, 0, run.toString)
      wb.sheets.headOption.getOrElse(fail("no sheet"))(ref"D1").value match
        case CellValue.Formula(_, Some(CellValue.Number(v)), _) => assertEquals(v, BigDecimal(1))
        case other => fail(s"D1 should cache the anchor's value 1, got $other")
  }
