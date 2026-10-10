package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.DynamicArrayFixtures

/**
 * GH-714 end to end: a dynamic-array anchor (`cm` + `xl/metadata.xml`) survives every in-memory
 * write, shows as Excel 365 shows it (no braces), and reports `formulaKind: "dynamicArray"`.
 */
class DynamicArrayCliSpec extends CatsEffectSuite:

  protected def source(parts: Vector[(String, String)] = DynamicArrayFixtures.parts()): IO[Path] =
    IO.blocking {
      val p = Files.createTempFile("xl-gh714-src-", ".xlsx")
      Files.write(p, DynamicArrayFixtures.xlsx(parts))
      p
    }

  protected def outPath(): IO[Path] = IO.blocking {
    val p = Files.createTempFile("xl-gh714-out-", ".xlsx")
    Files.delete(p)
    p
  }

  protected def bytes(p: Path): IO[Array[Byte]] = IO.blocking(Files.readAllBytes(p))

  protected def sheetXml(p: Path): IO[String] =
    bytes(p).map(b =>
      DynamicArrayFixtures.entryText(b, "xl/worksheets/sheet1.xml").getOrElse(fail("no sheet1"))
    )

  protected def kindAt(p: Path, at: ARef): IO[Option[FormulaKind]] =
    ExcelIO.instance[IO].read(p).map { wb =>
      wb.sheets.headOption.flatMap(_.cells.get(at)).map(_.value).collect {
        case CellValue.Formula(_, _, kind) => kind
      }
    }

  protected def ok(run: CliRun): Unit = assertEquals(run.exit, 0, run.toString)

  test("issue repro: put D1 keeps B1's cm, so Excel still reads a dynamic array") {
    for
      src <- source()
      out <- outPath()
      run <- CliHarness.run("-f", src.toString, "-o", out.toString, "put", "D1", "5")
      xml <- sheetXml(out)
      b <- bytes(out)
      kind <- kindAt(out, ref"B1")
    yield
      ok(run)
      assert(
        xml.contains("""cm="1"><f t="array" ref="B1:B5">_xlfn._xlws.SORT(A1:A5)</f>"""),
        xml
      )
      assertEquals(
        DynamicArrayFixtures.entryText(b, "xl/metadata.xml"),
        Some(DynamicArrayFixtures.metadataXml)
      )
      assertEquals(kind, Some(FormulaKind.dynamicArray(CellRange(ref"B1", ref"B5"))))
  }

  test("--stream put elsewhere passes the anchor's cm through untouched") {
    for
      src <- source()
      out <- outPath()
      run <- CliHarness.run("-f", src.toString, "-o", out.toString, "--stream", "put", "D1", "5")
      xml <- sheetXml(out)
    yield
      ok(run)
      assert(xml.contains("""<c r="B1" cm="1">"""), xml)
  }

  test("--stream putf over the anchor drops its cm (no phantom dynamic array)") {
    for
      src <- source()
      out <- outPath()
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-o",
        out.toString,
        "--stream",
        "putf",
        "B1",
        "=A1*2"
      )
      xml <- sheetXml(out)
      kind <- kindAt(out, ref"B1")
    yield
      ok(run)
      assert(!xml.contains("""<c r="B1" cm="1""""), xml)
      assertEquals(kind, Some(FormulaKind.Normal()))
  }

  test("cell and view --formulas show a dynamic array without braces") {
    for
      src <- source()
      cell <- CliHarness.run("-f", src.toString, "cell", "B1")
      view <- CliHarness.run("-f", src.toString, "view", "B1", "--formulas")
    yield
      ok(cell)
      ok(view)
      assert(cell.stdout.contains("=SORT(A1:A5)"), cell.stdout)
      assert(!cell.stdout.contains("{=SORT"), cell.stdout)
      assert(view.stdout.contains("=SORT(A1:A5)"), view.stdout)
      assert(!view.stdout.contains("{=SORT"), view.stdout)
  }

  test("JSON output names the kind dynamicArray") {
    for
      src <- source()
      run <- CliHarness.run("-f", src.toString, "view", "B1", "--format", "json")
    yield
      ok(run)
      assert(run.stdout.contains(""""formulaKind": "dynamicArray""""), run.stdout)
  }
