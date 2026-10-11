package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import cats.syntax.all.*
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

  // ---------------------------------------------------------------------------------------------
  // putf --array / batch "array": dynamic-array authoring

  /** A1:A10 = 1..10, B1:B10 = 1..10, D1:D3 = 3, 1, 2 — written by xl itself (no metadata part). */
  protected def numbers(): IO[Path] =
    for
      p <- IO.blocking(Files.createTempFile("xl-gh714-nums-", ".xlsx"))
      sheet = (1 to 10)
        .foldLeft(Sheet("Sheet1")) { (s, i) =>
          s.put(ARef.from0(0, i - 1), CellValue.Number(BigDecimal(i)))
            .put(ARef.from0(1, i - 1), CellValue.Number(BigDecimal(i)))
        }
        .put(ref"D1", CellValue.Number(3))
        .put(ref"D2", CellValue.Number(1))
        .put(ref"D3", CellValue.Number(2))
      _ <- ExcelIO.instance[IO].write(Workbook(sheet), p)
    yield p

  protected def valueAt(p: Path, at: ARef): IO[CellValue] =
    ExcelIO.instance[IO].read(p).map(_.sheets.headOption.fold(CellValue.Empty)(_(at).value))

  private def putfArray(src: Path, out: Path, at: String, formula: String, extra: String*) =
    CliHarness.run(
      (List("-f", src.toString, "-o", out.toString, "putf", "--array", at, formula) ++ extra)*
    )

  test("putf --array writes the anchor, the spill values, cm and the metadata part") {
    for
      src <- numbers()
      out <- outPath()
      run <- putfArray(src, out, "F1", "=SORT(D1:D3)")
      xml <- sheetXml(out)
      b <- bytes(out)
      kind <- kindAt(out, ref"F1")
      f2 <- valueAt(out, ref"F2")
      f3 <- valueAt(out, ref"F3")
    yield
      ok(run)
      assert(run.stdout.contains("Put dynamic array =SORT(D1:D3) at F1 (spills F1:F3)"), run.stdout)
      assert(
        xml.contains(
          """<c r="F1" t="n" cm="1"><f t="array" ref="F1:F3">_xlfn._xlws.SORT(D1:D3)</f><v>1</v></c>"""
        ),
        xml
      )
      assertEquals(
        DynamicArrayFixtures.entryText(b, "xl/metadata.xml"),
        Some(DynamicArrayFixtures.metadataXml)
      )
      assertEquals(kind, Some(FormulaKind.dynamicArray(CellRange(ref"F1", ref"F3"))))
      assertEquals(f2, CellValue.Number(2))
      assertEquals(f3, CellValue.Number(3))
  }

  test("putf --array of a scalar array formula: a bare 1x1 record with the array value") {
    for
      src <- numbers()
      out <- outPath()
      run <- putfArray(src, out, "C5", "=SUM(A1:A10*B1:B10)")
      xml <- sheetXml(out)
      c5 <- valueAt(out, ref"C5")
    yield
      ok(run)
      // 1²+…+10² = 385, never row 5's implicit intersection (25)
      assert(
        xml.contains(
          """<c r="C5" t="n" cm="1"><f t="array" ref="C5">SUM(A1:A10*B1:B10)</f><v>385</v></c>"""
        ),
        xml
      )
      c5 match
        case CellValue.Formula(_, Some(CellValue.Number(v)), _) => assertEquals(v, BigDecimal(385))
        case other => fail(s"C5 should cache 385, got $other")
  }

  test("a blocked spill is FORMULA_ERROR naming the blocker, and writes nothing") {
    for
      src <- numbers()
      out <- outPath()
      // B9:B11 would spill over B10 = 10 (B9 is the anchor itself)
      run <- putfArray(src, out, "B9", "=SORT(A1:A3)", "--json")
      exists <- IO.blocking(Files.exists(out))
    yield
      assertEquals(run.exit, 3, run.toString)
      assert(run.stdout.contains("FORMULA_ERROR"), run.stdout)
      assert(run.stdout.contains("spill range B9:B11 is blocked by B10"), run.stdout)
      assert(run.stdout.contains("clear B10 or choose another anchor"), run.stdout)
      assert(run.stdout.contains("#SPILL!"), run.stdout)
      assert(!exists, "a refused spill must not write the output")
  }

  test("a dependent of a spill cell is refreshed in the same command") {
    for
      src <- numbers()
      mid <- outPath()
      _ <- CliHarness.run("-f", src.toString, "-o", mid.toString, "putf", "H1", "=F2*10")
      out <- outPath()
      run <- putfArray(mid, out, "F1", "=SORT(D1:D3)")
      h1 <- valueAt(out, ref"H1")
    yield
      ok(run)
      h1 match
        case CellValue.Formula(_, Some(CellValue.Number(v)), _) => assertEquals(v, BigDecimal(20))
        case other => fail(s"H1 should read the spill's 2, got $other")
  }

  test("--no-recalc still sizes the extent from evaluation") {
    for
      src <- numbers()
      out <- outPath()
      run <- putfArray(src, out, "F1", "=SORT(D1:D3)", "--no-recalc")
      kind <- kindAt(out, ref"F1")
    yield
      ok(run)
      assertEquals(kind, Some(FormulaKind.dynamicArray(CellRange(ref"F1", ref"F3"))))
  }

  test("--array with a range or two formulas is a USAGE error") {
    for
      src <- numbers()
      out <- outPath()
      range <- putfArray(src, out, "F1:F3", "=SORT(D1:D3)", "--json")
      two <- CliHarness.run(
        "-f",
        src.toString,
        "-o",
        out.toString,
        "putf",
        "--array",
        "F1:F2",
        "=1",
        "=2",
        "--json"
      )
    yield
      assert(range.stdout.contains("USAGE"), range.toString)
      assert(range.stdout.contains("--array anchors at one cell"), range.toString)
      assert(two.stdout.contains("USAGE"), two.toString)
  }

  test("--stream --array is UNSUPPORTED_IN_STREAM, and nothing is written") {
    for
      src <- numbers()
      out <- outPath()
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-o",
        out.toString,
        "--stream",
        "putf",
        "--array",
        "F1",
        "=SORT(D1:D3)",
        "--json"
      )
      exists <- IO.blocking(Files.exists(out))
    yield
      assertEquals(run.exit, 2, run.toString)
      assert(run.stdout.contains("UNSUPPORTED_IN_STREAM"), run.stdout)
      assert(!exists, "nothing may be written")
  }

  test("batch array:true authors a dynamic array; refresh and display follow") {
    val ops = """[{"op":"putf","ref":"F1","value":"=SORT(D1:D3)","array":true}]"""
    for
      src <- numbers()
      out <- outPath()
      run <- CliHarness.run(List("-f", src.toString, "-o", out.toString, "batch", "-"), ops)
      kind <- kindAt(out, ref"F1")
      f3 <- valueAt(out, ref"F3")
      cell <- CliHarness.run("-f", out.toString, "cell", "F1")
    yield
      ok(run)
      assertEquals(kind, Some(FormulaKind.dynamicArray(CellRange(ref"F1", ref"F3"))))
      assertEquals(f3, CellValue.Number(3))
      assert(cell.stdout.contains("=SORT(D1:D3)") && !cell.stdout.contains("{="), cell.stdout)
  }

  test("batch array:true with from, values or a range is BATCH_OP_INVALID, --dry-run included") {
    val bad = Vector(
      """[{"op":"putf","ref":"F1:F3","value":"=SORT(D1:D3)","array":true}]""",
      """[{"op":"putf","ref":"F1:F3","value":"=D1","from":"F1","array":true}]""",
      """[{"op":"putf","ref":"F1:F2","values":["=1","=2"],"array":true}]""",
      """[{"op":"putf","ref":"F1","value":"=1","array":"yes"}]"""
    )
    for
      src <- numbers()
      out <- outPath()
      runs <- bad.traverse(ops =>
        CliHarness.run(List("-f", src.toString, "-o", out.toString, "batch", "-", "--json"), ops)
      )
      dry <- bad.traverse(ops => CliHarness.run(List("batch", "--dry-run", "-", "--json"), ops))
    yield (runs ++ dry).foreach { run =>
      assert(run.stdout.contains("BATCH_OP_INVALID"), run.toString)
    }
  }

  test("a streamed batch with array:true is refused before any write") {
    val ops =
      """[{"op":"put","ref":"Z1","value":1},{"op":"putf","ref":"F1","value":"=SORT(D1:D3)","array":true}]"""
    for
      src <- numbers()
      out <- outPath()
      run <- CliHarness.run(
        List("-f", src.toString, "-o", out.toString, "--stream", "batch", "-", "--json"),
        ops
      )
      exists <- IO.blocking(Files.exists(out))
    yield
      assert(run.stdout.contains("UNSUPPORTED_IN_STREAM"), run.toString)
      // the second op (1-based index 2) is the one refused
      assert(run.stdout.contains("\"opIndex\": 2"), run.stdout)
      assert(!exists, "nothing may be written")
  }

  test("xl batch --schema publishes the array field and x-streamRefusedFields") {
    for run <- CliHarness.run("batch", "--schema")
    yield
      ok(run)
      assert(run.stdout.contains("x-streamRefusedFields"), run.stdout.take(400))
      assert(run.stdout.contains("dynamic-array anchor"), run.stdout.take(400))
  }
