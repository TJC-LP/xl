package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.cli.contract.{CliError, CliHarness, CliRun}
import com.tjclp.xl.formula.eval.DynamicArrayAuthoring
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

  // ---------------------------------------------------------------------------------------------
  // The spill refusal's hint comes from the blocker kind (PR #721 review)

  private def errorOf(run: CliRun): ujson.Value = ujson.read(run.stdout)("error")

  test("every blocker kind maps to FORMULA_ERROR with that kind's own hint") {
    import DynamicArrayAuthoring.{SpillBlocked, SpillBlocker}
    val extent = Some(CellRange(ref"A1", ref"A3"))
    def hint(blocker: SpillBlocker, at: Option[CellRange] = extent): CliError =
      CliError.fromSpillBlocked(SpillBlocked("=SORT(B1:B3)", ref"A1", at, blocker))
    val cases = Vector(
      hint(SpillBlocker.Occupied(Vector(ref"A2", ref"A3"))) ->
        "clear A2, A3 or choose another anchor",
      hint(SpillBlocker.Merged(Vector(CellRange(ref"A2", ref"B3")))) ->
        "unmerge A2:B3 or choose another anchor",
      hint(SpillBlocker.ArrayRecord(Vector(ref"A2" -> CellRange(ref"A2", ref"C2")))) ->
        "clear or move the array formula at A2, or choose another anchor",
      hint(SpillBlocker.Table(Vector("Sales" -> CellRange(ref"A2", ref"B9")))) ->
        "choose an anchor whose spill misses table Sales, or convert the table to a range",
      hint(SpillBlocker.OffGrid(3, 1), None) ->
        "choose an anchor with 3 row(s) and 1 column(s) of room before XFD1048576",
      hint(SpillBlocker.Overwritten(Vector(ref"A2"))) ->
        "overwrite or clear A1 first (its spill goes with it), or write outside A1:A3"
    )
    // xl's own size guard: no #SPILL! in Excel, so its message says so plainly
    val tooLarge = hint(SpillBlocker.TooLarge(2000000, 1, 1048576L), None)
    assertEquals(tooLarge.code, "FORMULA_ERROR")
    assert(tooLarge.message.contains("at most 1,048,576 cells"), tooLarge.message)
    assertEquals(
      tooLarge.hint,
      Some("bound the formula's ranges to the data (A1:A1000, not A:A), or write a plain formula")
    )
    cases.foreach { (error, expected) =>
      assertEquals(error.code, "FORMULA_ERROR")
      assertEquals(error.hint, Some(expected))
      assert(error.message.contains("(Excel shows #SPILL!)"), error.message)
      assert(error.cause.nonEmpty)
    }
  }

  test("a merge-blocked spill says to unmerge, never to clear the merge") {
    for
      base <- numbers()
      merged <- ExcelIO.instance[IO].read(base).map { wb =>
        wb.put(wb.sheets.headOption.getOrElse(fail("no sheet")).merge(CellRange(ref"F2", ref"G3")))
      }
      src <- IO.blocking(Files.createTempFile("xl-gh714-merged-", ".xlsx"))
      _ <- ExcelIO.instance[IO].write(merged, src)
      out <- outPath()
      run <- putfArray(src, out, "F1", "=SORT(D1:D3)", "--json")
    yield
      assertEquals(run.exit, 3, run.toString)
      val error = errorOf(run)
      assertEquals(error("code").str, "FORMULA_ERROR")
      assert(error("message").str.contains("blocked by merged cells F2:G3"), run.stdout)
      assertEquals(error("hint").str, "unmerge F2:G3 or choose another anchor")
  }

  test("putf --array of a formula that cannot be evaluated says uncached, not spills") {
    for
      src <- numbers()
      out <- outPath()
      run <- putfArray(src, out, "F1", "=NoSuchName*2")
    yield
      ok(run)
      assert(run.stdout.contains("Put dynamic array =NoSuchName*2 at F1 (uncached"), run.stdout)
      assert(!run.stdout.contains("spills F1"), run.stdout)
  }

  // ---------------------------------------------------------------------------------------------
  // A plain write over a dynamic array (PR #721 review): Excel's rule, CLI and batch

  /** `numbers()` with F1 = SORT(D1:D3) spilling F1:F3 and H1 = F2*10 reading the spill. */
  protected def withSpill(): IO[Path] =
    for
      src <- numbers()
      mid <- outPath()
      _ <- CliHarness.run("-f", src.toString, "-o", mid.toString, "putf", "H1", "=F2*10")
      out <- outPath()
      run <- putfArray(mid, out, "F1", "=SORT(D1:D3)")
      _ = ok(run)
    yield out

  private def cachedNumber(p: Path, at: ARef): IO[BigDecimal] =
    valueAt(p, at).map {
      case CellValue.Formula(_, Some(CellValue.Number(v)), _) => v
      case other => fail(s"${at.toA1} should cache a number, got $other")
    }

  test("a plain putf over a dynamic anchor clears its spill, and dependents refresh") {
    for
      src <- withSpill()
      out <- outPath()
      run <- CliHarness.run("-f", src.toString, "-o", out.toString, "putf", "F1", "=1")
      kind <- kindAt(out, ref"F1")
      f2 <- valueAt(out, ref"F2")
      f3 <- valueAt(out, ref"F3")
      h1 <- cachedNumber(out, ref"H1")
      xml <- sheetXml(out)
    yield
      ok(run)
      assertEquals(kind, Some(FormulaKind.Normal()))
      assertEquals(f2, CellValue.Empty)
      assertEquals(f3, CellValue.Empty)
      // H1 = F2*10 read the spill's 2; F2 is now empty
      assertEquals(h1, BigDecimal(0))
      assert(!xml.contains("cm=\"1\""), xml)
  }

  test("a plain put over a dynamic anchor, alone or with its spill, releases it") {
    for
      src <- withSpill()
      single <- outPath()
      one <- CliHarness.run("-f", src.toString, "-o", single.toString, "put", "F1", "7")
      f3 <- valueAt(single, ref"F3")
      whole <- outPath()
      all <- CliHarness.run("-f", src.toString, "-o", whole.toString, "put", "F1:F3", "4", "5", "6")
      kind <- kindAt(whole, ref"F1")
      f2 <- valueAt(whole, ref"F2")
    yield
      ok(one)
      ok(all)
      assertEquals(f3, CellValue.Empty)
      assertEquals(kind, None)
      assertEquals(f2, CellValue.Number(5))
  }

  test("a put into a spill cell is refused as the #SPILL! it would cause; nothing written") {
    for
      src <- withSpill()
      out <- outPath()
      put <- CliHarness.run("-f", src.toString, "-o", out.toString, "put", "F2", "9", "--json")
      drag <- CliHarness.run(
        "-f",
        src.toString,
        "-o",
        out.toString,
        "putf",
        "F3:F4",
        "=D3",
        "--json"
      )
      exists <- IO.blocking(Files.exists(out))
    yield
      assertEquals(put.exit, 3, put.toString)
      val error = errorOf(put)
      assertEquals(error("code").str, "FORMULA_ERROR")
      assert(
        error("message").str.contains(
          "F2: it is in the spill range F1:F3 of the dynamic array =SORT(D1:D3) at F1"
        ),
        put.stdout
      )
      assertEquals(
        error("hint").str,
        "overwrite or clear F1 first (its spill goes with it), or write outside F1:F3"
      )
      assertEquals(drag.exit, 3, drag.toString)
      assert(errorOf(drag)("message").str.startsWith("F3: it is in the spill range"), drag.stdout)
      assert(!exists, "a refused write must not write the output")
  }

  test("batch: a later op into a spill an earlier op authored fails like Excel's #SPILL!") {
    val ops =
      """[{"op":"putf","ref":"F1","value":"=SORT(D1:D3)","array":true},
        |{"op":"put","ref":"F2","value":9}]""".stripMargin
    for
      src <- numbers()
      out <- outPath()
      run <- CliHarness.run(
        List("-f", src.toString, "-o", out.toString, "batch", "-", "--json"),
        ops
      )
      exists <- IO.blocking(Files.exists(out))
    yield
      assertEquals(run.exit, 3, run.toString)
      val error = errorOf(run)
      assertEquals(error("code").str, "BATCH_OP_FAILED")
      assertEquals(error("location")("opIndex").num.toInt, 2)
      assert(error("message").str.contains("F2: it is in the spill range F1:F3"), run.stdout)
      assertEquals(
        error("hint").str,
        "overwrite or clear F1 first (its spill goes with it), or write outside F1:F3"
      )
      assert(!exists, "a failed batch must not write the output")
  }

  test("batch: a plain putf over an anchor clears its spill, an earlier op's or the file's") {
    val authored =
      """[{"op":"putf","ref":"F1","value":"=SORT(D1:D3)","array":true},
        |{"op":"putf","ref":"F1","value":"=1"}]""".stripMargin
    val overFile = """[{"op":"put","ref":"F1","value":"x"}]"""
    for
      src <- numbers()
      out <- outPath()
      run <- CliHarness.run(List("-f", src.toString, "-o", out.toString, "batch", "-"), authored)
      f2 <- valueAt(out, ref"F2")
      f3 <- valueAt(out, ref"F3")
      spilled <- withSpill()
      out2 <- outPath()
      run2 <- CliHarness.run(
        List("-f", spilled.toString, "-o", out2.toString, "batch", "-"),
        overFile
      )
      g2 <- valueAt(out2, ref"F2")
      h1 <- cachedNumber(out2, ref"H1")
    yield
      ok(run)
      ok(run2)
      assertEquals(f2, CellValue.Empty)
      assertEquals(f3, CellValue.Empty)
      assertEquals(g2, CellValue.Empty)
      assertEquals(h1, BigDecimal(0))
  }

  test("an Excel-written spill: put over its anchor clears it, put into it is refused") {
    for
      src <- source()
      out <- outPath()
      over <- CliHarness.run("-f", src.toString, "-o", out.toString, "put", "B1", "0")
      b5 <- valueAt(out, ref"B5")
      into <- CliHarness.run("-f", src.toString, "-o", out.toString, "put", "B4", "0", "--json")
    yield
      ok(over)
      assertEquals(b5, CellValue.Empty)
      assertEquals(into.exit, 3, into.toString)
      assert(errorOf(into)("message").str.contains("in the spill range B1:B5"), into.stdout)
  }

  // ---------------------------------------------------------------------------------------------
  // Spill children after a later recalculation (PR #721 review): pinned limitation

  test("a later input change refreshes the anchor's cache only: no re-spill, no resize") {
    for
      src <- withSpill()
      // D1:D3 = 3, 1, 2 → F1:F3 = 1, 2, 3; D1 = 0 makes SORT {0; 1; 2}
      changed <- outPath()
      run <- CliHarness.run("-f", src.toString, "-o", changed.toString, "put", "D1", "0")
      f1 <- cachedNumber(changed, ref"F1")
      f2 <- valueAt(changed, ref"F2")
      f3 <- valueAt(changed, ref"F3")
      kind <- kindAt(changed, ref"F1")
      h1 <- cachedNumber(changed, ref"H1")
      // a growing result: SEQUENCE(D2) at J1 is 1x1 while D2 = 1, and stays 1x1 after D2 = 3
      seq <- outPath()
      _ <- putfArray(src, seq, "J1", "=SEQUENCE(D2)")
      grown <- outPath()
      grow <- CliHarness.run("-f", seq.toString, "-o", grown.toString, "put", "D2", "3")
      seqKind <- kindAt(grown, ref"J1")
      j2 <- valueAt(grown, ref"J2")
    yield
      ok(run)
      ok(grow)
      // the anchor's cache is the new array's top-left value ...
      assertEquals(f1, BigDecimal(0))
      // ... while the spill cells keep their authored values and the extent its authored size
      // (Excel re-spills on open: it recalculates a book xl wrote, which carries no calcId)
      assertEquals(f2, CellValue.Number(2))
      assertEquals(f3, CellValue.Number(3))
      assertEquals(kind, Some(FormulaKind.dynamicArray(CellRange(ref"F1", ref"F3"))))
      assertEquals(h1, BigDecimal(20))
      assertEquals(seqKind, Some(FormulaKind.dynamicArray(CellRange(ref"J1", ref"J1"))))
      assertEquals(j2, CellValue.Empty)
  }

  // ---------------------------------------------------------------------------------------------
  // copy, fill and clear follow the put/putf rule (third #721 review)

  private def run(src: Path, out: Path, args: String*): IO[CliRun] =
    CliHarness.run((List("-f", src.toString, "-o", out.toString) ++ args)*)

  private def assertRefused(run: CliRun, code: String, mention: String): Unit =
    assertEquals(run.exit, 3, run.toString)
    val error = errorOf(run)
    assertEquals(error("code").str, code)
    assert(error("message").str.contains(mention), run.stdout)

  test("copy or fill into a live spill is refused; nothing is written") {
    for
      src <- withSpill()
      out <- outPath()
      copy <- run(src, out, "copy", "D2", "F2", "--json")
      // fill right from D2:D3 writes E2:F3, and F2:F3 belong to F1's spill
      fill <- run(src, out, "fill", "D2:D3", "D2:F3", "--right", "--json")
      exists <- IO.blocking(Files.exists(out))
    yield
      assertRefused(copy, "FORMULA_ERROR", "F2: it is in the spill range F1:F3")
      assertEquals(
        errorOf(copy)("hint").str,
        "overwrite or clear F1 first (its spill goes with it), or write outside F1:F3"
      )
      assertRefused(fill, "FORMULA_ERROR", "F2, F3: they are in the spill range F1:F3")
      assert(!exists, "a refused write must not write the output")
  }

  test("copy or fill over a dynamic anchor releases its spill; dependents refresh") {
    for
      src <- withSpill()
      copied <- outPath()
      copy <- run(src, copied, "copy", "D1", "F1")
      c1 <- valueAt(copied, ref"F1")
      c2 <- valueAt(copied, ref"F2")
      c3 <- valueAt(copied, ref"F3")
      h1 <- cachedNumber(copied, ref"H1")
      filled <- outPath()
      // fill right from D1 writes E1 and F1: the anchor goes, and so do F2:F3
      fill <- run(src, filled, "fill", "D1", "D1:F1", "--right")
      f1 <- valueAt(filled, ref"F1")
      f2 <- valueAt(filled, ref"F2")
      f3 <- valueAt(filled, ref"F3")
    yield
      ok(copy)
      ok(fill)
      assertEquals(c1, CellValue.Number(3))
      assertEquals((c2, c3), (CellValue.Empty, CellValue.Empty))
      assertEquals(h1, BigDecimal(0))
      assertEquals(f1, CellValue.Number(3))
      assertEquals((f2, f3), (CellValue.Empty, CellValue.Empty))
  }

  test("a whole spill copied elsewhere and copied onto itself keeps working") {
    for
      src <- withSpill()
      out <- outPath()
      // the source spill F1:F3 lands at K1:K3 as a complete spill (665f2b9)
      away <- run(src, out, "copy", "F1:F3", "K1")
      kind <- kindAt(out, ref"K1")
      onto <- outPath()
      // copying the spill over itself rewrites the anchor and its own cells: no refusal
      self <- run(src, onto, "copy", "F1:F3", "F1")
      selfKind <- kindAt(onto, ref"F1")
    yield
      ok(away)
      ok(self)
      assertEquals(kind, Some(FormulaKind.dynamicArray(CellRange(ref"K1", ref"K3"))))
      assertEquals(selfKind, Some(FormulaKind.dynamicArray(CellRange(ref"F1", ref"F3"))))
  }

  private def warningsOf(run: CliRun): Vector[ujson.Value] =
    ujson.read(run.stdout).obj.get("warnings").fold(Vector.empty)(_.arr.toVector)

  private def skipped(run: CliRun): Vector[String] =
    warningsOf(run).filter(_("code").str == "SPILL_CELLS_SKIPPED").map(_("message").str)

  test("clear: the anchor takes its spill; cells of a spill whose anchor stays are kept") {
    for
      src <- withSpill()
      anchor <- outPath()
      whole <- run(src, anchor, "clear", "F1:F2", "--json")
      f3 <- valueAt(anchor, ref"F3")
      // D3:F3 crosses F1's spill at F3, its anchor outside the range
      out <- outPath()
      partial <- run(src, out, "clear", "D3:F3", "--json")
      d3 <- valueAt(out, ref"D3")
      keptF3 <- valueAt(out, ref"F3")
      anchorKind <- kindAt(out, ref"F1")
      styled <- outPath()
      styles <- run(src, styled, "clear", "F2", "--styles")
      kept <- valueAt(styled, ref"F2")
      everything <- outPath()
      all <- run(src, everything, "clear", "F2", "--all", "--json")
      keptAll <- valueAt(everything, ref"F2")
    yield
      ok(whole)
      assertEquals(f3, CellValue.Empty)
      assertEquals(skipped(whole), Vector.empty)
      // the rest of the range is cleared, the spill cell is left, and the warning says so
      ok(partial)
      assertEquals(d3, CellValue.Empty)
      assertEquals(keptF3, CellValue.Number(3))
      assertEquals(anchorKind, Some(FormulaKind.dynamicArray(CellRange(ref"F1", ref"F3"))))
      assertEquals(skipped(partial).size, 1, partial.stdout)
      assert(
        skipped(partial).forall(_.contains("Sheet1!F3 (the spill of =SORT(D1:D3) at Sheet1!F1)")),
        partial.stdout
      )
      ok(styles)
      assertEquals(kept, CellValue.Number(2))
      // --all: the contents part skips the spill cell (styles and comments clear as anywhere)
      ok(all)
      assertEquals(keptAll, CellValue.Number(2))
      assertEquals(skipped(all).size, 1, all.stdout)
  }

  test("batch copy and clear follow the same rule") {
    val copyInto = """[{"op":"copy","source":"D2","target":"F3"}]"""
    val clearInto = """[{"op":"clear","range":"D2:F3"},{"op":"put","ref":"A1","value":1}]"""
    val clearAnchor =
      """[{"op":"clear","range":"F1"},{"op":"copy","source":"D1:D2","target":"F2"}]"""
    for
      src <- withSpill()
      out <- outPath()
      copy <- CliHarness.run(
        List("-f", src.toString, "-o", out.toString, "batch", "-", "--json"),
        copyInto
      )
      cleared <- outPath()
      clear <- CliHarness.run(
        List("-f", src.toString, "-o", cleared.toString, "batch", "-", "--json"),
        clearInto
      )
      d2 <- valueAt(cleared, ref"D2")
      keptF2 <- valueAt(cleared, ref"F2")
      released <- outPath()
      both <- CliHarness.run(
        List("-f", src.toString, "-o", released.toString, "batch", "-"),
        clearAnchor
      )
      f2 <- valueAt(released, ref"F2")
      f3 <- valueAt(released, ref"F3")
    yield
      assertRefused(copy, "BATCH_OP_FAILED", "F3: it is in the spill range F1:F3")
      // the clear skips F1's spill cells, clears the rest and the batch goes on
      ok(clear)
      assertEquals(d2, CellValue.Empty)
      assertEquals(keptF2, CellValue.Number(2))
      assertEquals(skipped(clear).size, 1, clear.stdout)
      assert(skipped(clear).forall(_.contains("Sheet1!F2:F3")), clear.stdout)
      ok(both)
      // the clear took the spill, so the later copy writes plain constants where it was
      assertEquals((f2, f3), (CellValue.Number(3), CellValue.Number(1)))
  }

  test("copying a whole spill onto a merge or a table is refused, as Excel's #SPILL!") {
    import com.tjclp.xl.tables.{TableColumn, TableSpec}
    val table = TableSpec(
      name = "T1",
      displayName = "T1",
      range = CellRange(ref"L2", ref"M4"),
      columns = Vector(TableColumn(1, "a"), TableColumn(2, "b"))
    )
    def rewritten(src: Path)(f: Sheet => Sheet): IO[Path] =
      for
        wb <- ExcelIO.instance[IO].read(src)
        p <- IO.blocking(Files.createTempFile("xl-gh714-target-", ".xlsx"))
        _ <- ExcelIO
          .instance[IO]
          .write(wb.put(f(wb.sheets.headOption.getOrElse(fail("no sheet")))), p)
      yield p
    for
      src <- withSpill()
      merged <- rewritten(src)(_.merge(CellRange(ref"K2", ref"K3")))
      tabled <- rewritten(src)(s => s.copy(tables = Map("T1" -> table)))
      out <- outPath()
      ontoMerge <- run(merged, out, "copy", "F1:F3", "K1", "--json")
      ontoTable <- run(tabled, out, "copy", "F1:F3", "L1", "--json")
      batch <- CliHarness.run(
        List("-f", merged.toString, "-o", out.toString, "batch", "-", "--json"),
        """[{"op":"copy","source":"F1:F3","target":"K1"}]"""
      )
      exists <- IO.blocking(Files.exists(out))
      // the same spill copied to open cells keeps its extent there (665f2b9)
      clear <- outPath()
      open <- run(merged, clear, "copy", "F1:F3", "N1")
      kind <- kindAt(clear, ref"N1")
    yield
      assertRefused(
        ontoMerge,
        "FORMULA_ERROR",
        "spill range K1:K3 is blocked by merged cells K2:K3"
      )
      assertEquals(errorOf(ontoMerge)("hint").str, "unmerge K2:K3 or choose another anchor")
      assertRefused(ontoTable, "FORMULA_ERROR", "overlaps table T1 (L2:M4)")
      assertRefused(batch, "BATCH_OP_FAILED", "blocked by merged cells K2:K3")
      assert(!exists, "a refused copy must not write the output")
      ok(open)
      assertEquals(kind, Some(FormulaKind.dynamicArray(CellRange(ref"N1", ref"N3"))))
  }

  test("batch: the batch's clock is the one an array spill is authored with") {
    import com.tjclp.xl.cli.helpers.BatchParser
    import com.tjclp.xl.formula.Clock
    val at = java.time.LocalDateTime.of(2020, 1, 2, 3, 4, 5)
    val fixed = Clock.fixed(at.toLocalDate, at)
    val ops = """[{"op":"putf","ref":"F1","value":"=NOW()+SEQUENCE(3)","array":true}]"""
    for
      src <- numbers()
      wb <- ExcelIO.instance[IO].read(src)
      parsed <- BatchParser.parseBatchOperations(ops)
      result <- BatchParser.applyScopedReporting(
        wb,
        wb.sheets.headOption,
        parsed.scoped,
        recalcDependents = false,
        clock = fixed
      )
    yield
      val sheet = result._1.sheets.headOption.getOrElse(fail("no sheet"))
      val expected =
        sheet.evaluateFormula("=NOW()+3", clock = fixed).fold(e => fail(e.message), identity)
      // F3 = NOW()+3 at the batch's instant, not at whatever time the op happened to run
      assertEquals(sheet(ref"F3").value, expected)
  }

  test("batch: one clock for an array spill of NOW(), its refresh and the check") {
    val ops = """[{"op":"putf","ref":"F1","value":"=NOW()+SEQUENCE(3)","array":true}]"""
    for
      src <- numbers()
      out <- outPath()
      run <- CliHarness.run(List("-f", src.toString, "-o", out.toString, "batch", "-"), ops)
      f1 <- cachedNumber(out, ref"F1")
      f2 <- valueAt(out, ref"F2")
      f3 <- valueAt(out, ref"F3")
    yield
      ok(run)
      // the refresh recomputed the anchor; its cache and the authored spill agree exactly
      (f2, f3) match
        case (CellValue.Number(two), CellValue.Number(three)) =>
          assertEquals(two - f1, BigDecimal(1))
          assertEquals(three - f1, BigDecimal(2))
        case other => fail(s"F2:F3 should hold numbers, got $other")
  }

  test("--stream put elsewhere keeps an untouched cell's vm and cm") {
    // E1 is an image in a cell as Excel stores it: an error value whose vm names rich-value data
    val withImage = DynamicArrayFixtures.sortSheetXml.replace(
      """<v>15</v></c></row>""",
      """<v>15</v></c><c r="E1" t="e" vm="1"><v>#VALUE!</v></c></row>"""
    )
    assert(withImage.contains("vm=\"1\""), "fixture edit missed")
    for
      src <- source(DynamicArrayFixtures.parts(sheetXml = withImage))
      out <- outPath()
      run <- CliHarness.run("-f", src.toString, "-o", out.toString, "--stream", "put", "D2", "5")
      xml <- sheetXml(out)
    yield
      ok(run)
      assert(xml.contains("""<c r="E1" t="e" vm="1">"""), xml)
      assert(xml.contains("""<c r="B1" cm="1">"""), xml)
  }

  // ---------------------------------------------------------------------------------------------
  // Fourth #721 review

  test("--stream reads the metadata part under the run's --max-size limits") {
    // a highly compressible metadata part trips the default compression-ratio limit
    val padded = DynamicArrayFixtures.parts().map { (name, xml) =>
      if name == "xl/metadata.xml" then name -> (xml + (" " * 200000)) else name -> xml
    }
    for
      src <- source(padded)
      refused <- CliHarness.run("-f", src.toString, "--stream", "cell", "B1", "--json")
      lifted <- CliHarness.run(
        "-f",
        src.toString,
        "--max-size",
        "0",
        "--stream",
        "cell",
        "B1",
        "--json"
      )
    yield
      assertEquals(refused.exit, 3, refused.toString)
      assert(refused.stdout.contains("Compression ratio"), refused.stdout)
      ok(lifted)
      assert(lifted.stdout.contains("SORT(A1:A5)"), lifted.stdout)
  }

  test("a putf drag over an anchor and its own spill cells releases rather than refuses") {
    for
      src <- withSpill()
      out <- outPath()
      drag <- run(src, out, "putf", "F1:F3", "=D1*2")
      kind <- kindAt(out, ref"F1")
      f2 <- valueAt(out, ref"F2")
      h1 <- cachedNumber(out, ref"H1")
    yield
      ok(drag)
      assertEquals(kind, Some(FormulaKind.Normal()))
      f2 match
        case CellValue.Formula("D2*2", _, FormulaKind.Normal(_, _)) => ()
        case other => fail(s"F2 should hold the dragged plain formula, got $other")
      // H1 = F2*10 reads the dragged D2*2 = 2
      assertEquals(h1, BigDecimal(20))
  }

  test("putf with formulas: a malformed formula reports before a spill refusal") {
    for
      src <- withSpill()
      out <- outPath()
      res <- run(src, out, "putf", "F2:F3", "=1", "=SUM(", "--json")
    yield
      assertEquals(res.exit, 3, res.toString)
      val error = errorOf(res)
      assertEquals(error("code").str, "FORMULA_ERROR")
      assert(error("message").str.contains("does not parse"), res.stdout)
      assert(!error("message").str.contains("spill"), res.stdout)
  }

  test("a batch rescans a sheet for spills only when an op could have changed them") {
    import java.util.concurrent.atomic.AtomicInteger
    import com.tjclp.xl.cli.helpers.BatchParser
    // a 20,000-cell sheet whose F1 spills F1:F3
    val big = (0 until 4000).foldLeft(Sheet("Sheet1")) { (s, row) =>
      (0 until 5).foldLeft(s)((acc, col) =>
        acc.put(ARef.from0(col + 10, row), CellValue.Number(BigDecimal(row)))
      )
    }
    val seeded = big
      .put(ref"D1", CellValue.Number(3))
      .put(ref"D2", CellValue.Number(1))
      .put(ref"D3", CellValue.Number(2))
    val wb0 = Workbook(seeded)
    val authored = DynamicArrayAuthoring
      .author(wb0, seeded, ref"F1", "=SORT(D1:D3)", com.tjclp.xl.formula.Clock.system)
      .fold(r => fail(r.message), identity)
    val wb = wb0.put(authored.sheet)
    val puts = (1 to 500).map(i => s"""{"op":"put","ref":"H$i","value":$i}""")
    val styles = (1 to 50).map(i => s"""{"op":"style","range":"H$i","bold":true}""")
    val array = """{"op":"putf","ref":"Z1","value":"=SEQUENCE(2)","array":true}"""
    val ops = (puts.take(250) ++ styles ++ Vector(array) ++ puts.drop(250)).mkString("[", ",", "]")
    val scans = new AtomicInteger(0)
    val counting: Sheet => Vector[DynamicArrayAuthoring.Spill] = s =>
      scans.incrementAndGet()
      DynamicArrayAuthoring.spillsOf(s)
    for
      parsed <- BatchParser.parseBatchOperations(ops)
      result <- BatchParser.applyScopedReporting(
        wb,
        wb.sheets.headOption,
        parsed.scoped,
        recalcDependents = false,
        scanSpills = counting
      )
    yield
      // one scan for the first put, one after the array op dropped the index: not one per op
      assertEquals(scans.get(), 2)
      val sheet = result._1.sheets.headOption.getOrElse(fail("no sheet"))
      assertEquals(
        sheet(ref"F1").value match
          case CellValue.Formula(_, _, k) => Some(k)
          case _ => None
        ,
        Some(FormulaKind.dynamicArray(CellRange(ref"F1", ref"F3")))
      )
      assertEquals(sheet(ref"H500").value, CellValue.Number(500))
  }

  // ---------------------------------------------------------------------------------------------
  // Sixth #721 review: one refusal for an anchor inside a spill; a metadata part xl cannot extend

  test("putf --array inside another spill gets the Overwritten refusal a plain putf gets") {
    for
      src <- withSpill()
      out <- outPath()
      array <- putfArray(src, out, "F2", "=SORT(D1:D3)", "--json")
      plain <- run(src, out, "putf", "F2", "=SORT(D1:D3)", "--json")
      exists <- IO.blocking(Files.exists(out))
    yield
      assertRefused(array, "FORMULA_ERROR", "F2: it is in the spill range F1:F3")
      assertEquals(
        errorOf(array)("hint").str,
        "overwrite or clear F1 first (its spill goes with it), or write outside F1:F3"
      )
      assertEquals(errorOf(array), errorOf(plain))
      assert(!exists, "a refused write must not write the output")
  }

  test("an unreadable metadata part: the write succeeds and warns DYNAMIC_ARRAY_DOWNGRADED") {
    for
      src <- source(DynamicArrayFixtures.parts(metadata = Some("<metadata")))
      out <- outPath()
      array <- putfArray(src, out, "E1", "=SORT(A1:A5)", "--json")
      b <- bytes(out)
      kind <- kindAt(out, ref"E1")
      plainOut <- outPath()
      plain <- run(src, plainOut, "put", "D1", "9", "--json")
    yield
      ok(array)
      assert(array.stdout.contains("DYNAMIC_ARRAY_DOWNGRADED"), array.stdout)
      assert(array.stdout.contains("'xl/metadata.xml' is not readable XML"), array.stdout)
      // the part is kept as it was, and the anchor reads back as the legacy array it was saved as
      assertEquals(DynamicArrayFixtures.entryText(b, "xl/metadata.xml"), Some("<metadata"))
      assertEquals(kind, Some(FormulaKind.ArrayFormula(CellRange(ref"E1", ref"E5"))))
      // a write that authors no dynamic array loses nothing, so it does not warn
      ok(plain)
      assert(!plain.stdout.contains("DYNAMIC_ARRAY_DOWNGRADED"), plain.stdout)
  }

  test("a sheetMetadata relationship naming a missing part: the part is generated, no warning") {
    for
      src <- source(DynamicArrayFixtures.parts().filterNot(_._1 == "xl/metadata.xml"))
      out <- outPath()
      array <- putfArray(src, out, "E1", "=SORT(A1:A5)", "--json")
      b <- bytes(out)
      kind <- kindAt(out, ref"E1")
    yield
      ok(array)
      assert(!array.stdout.contains("DYNAMIC_ARRAY_DOWNGRADED"), array.stdout)
      assertEquals(
        DynamicArrayFixtures.entryText(b, "xl/metadata.xml"),
        Some(DynamicArrayFixtures.metadataXml)
      )
      assertEquals(kind, Some(FormulaKind.dynamicArray(CellRange(ref"E1", ref"E5"))))
  }

  test("batch: an unreadable metadata part warns DYNAMIC_ARRAY_DOWNGRADED too") {
    val ops = """[{"op":"putf","ref":"E1","value":"=SORT(A1:A5)","array":true}]"""
    for
      src <- source(DynamicArrayFixtures.parts(metadata = Some("<metadata")))
      out <- outPath()
      batch <- CliHarness.run(
        List("-f", src.toString, "-o", out.toString, "batch", "-", "--json"),
        ops
      )
    yield
      ok(batch)
      assert(batch.stdout.contains("DYNAMIC_ARRAY_DOWNGRADED"), batch.stdout)
  }

  // ---------------------------------------------------------------------------------------------
  // Seventh #721 review: a circular spill; the batch spill index per sheet

  test("putf --array A1 \"=A1:A3\" reads its own anchor: written uncached, not spilled") {
    for
      src <- numbers()
      out <- outPath()
      self <- putfArray(src, out, "A1", "=A1:A3")
      kind <- kindAt(out, ref"A1")
      a2 <- valueAt(out, ref"A2")
    yield
      // as Excel, which reports a circular reference rather than spilling
      ok(self)
      assert(self.stdout.contains("uncached: it could not be evaluated ("), self.stdout)
      assert(self.stdout.contains("circular reference"), self.stdout)
      assertEquals(kind, Some(FormulaKind.dynamicArray(CellRange(ref"A1", ref"A1"))))
      // A2 keeps its own value: nothing spilled over it
      assertEquals(a2, CellValue.Number(2))
  }

  test("a batch drops only the sheet an op changed from its spill index") {
    import java.util.concurrent.atomic.AtomicInteger
    import com.tjclp.xl.cli.helpers.BatchParser
    def spilled(sheet: Sheet): Sheet =
      sheet
        .put(ref"D1", CellValue.Number(3))
        .put(ref"D2", CellValue.Number(1))
        .put(ref"D3", CellValue.Number(2))
    val wb = Workbook(spilled(Sheet("One")), spilled(Sheet("Two")))
    // a copy on One between writes on Two: Two is never rescanned; One is, after the copy
    val ops = Vector(
      """{"op":"put","ref":"One!A1","value":1}""",
      """{"op":"put","ref":"Two!A1","value":1}""",
      """{"op":"copy","source":"One!D1","target":"One!B1"}""",
      """{"op":"put","ref":"Two!A2","value":2}""",
      """{"op":"put","ref":"One!A2","value":2}""",
      """{"op":"copy","source":"One!D2","target":"One!B2"}""",
      """{"op":"put","ref":"Two!A3","value":3}"""
    ).mkString("[", ",", "]")
    val scans = scala.collection.mutable.Map.empty[String, AtomicInteger]
    val counting: Sheet => Vector[DynamicArrayAuthoring.Spill] = s =>
      scans.getOrElseUpdate(s.name.value, new AtomicInteger(0)).incrementAndGet()
      DynamicArrayAuthoring.spillsOf(s)
    for
      parsed <- BatchParser.parseBatchOperations(ops)
      result <- BatchParser.applyScopedReporting(
        wb,
        wb.sheets.headOption,
        parsed.scoped,
        recalcDependents = false,
        scanSpills = counting
      )
    yield
      assertEquals(scans.get("Two").map(_.get), Some(1))
      // One: its first put, then the put after the first copy dropped it (each copy reads the
      // index it was given); dropping the whole index would rescan Two after each copy as well
      assertEquals(scans.get("One").map(_.get), Some(2), scans.toString)
      assertEquals(
        result._1.sheets.find(_.name.value == "Two").map(_(ref"A3").value),
        Some(CellValue.Number(3))
      )
  }

  // ---------------------------------------------------------------------------------------------
  // Ninth #721 review: a fill from a multi-cell dynamic anchor

  private def overlapping(run: CliRun): Vector[String] =
    warningsOf(run).filter(_("code").str == "FILLED_SPILLS_OVERLAP").map(_("message").str)

  test("fill from a multi-cell dynamic anchor warns that the copies' spills overlap") {
    for
      src <- withSpill()
      down <- outPath()
      // F1 = SORT(D1:D3) last spilled F1:F3; filling it down F5:F8 makes four anchors a row apart
      filled <- run(src, down, "fill", "F1", "F5:F8", "--json")
      f6 <- kindAt(down, ref"F6")
      // down from the anchor itself, the first copy lands in F1's own live spill: refused
      own <- outPath()
      intoOwn <- run(src, own, "fill", "F1", "F1:F4", "--json")
      across <- outPath()
      // across, the copies sit side by side: a one-column spill cannot reach the next copy
      right <- run(src, across, "fill", "F1", "F1:G1", "--right", "--json")
      plain <- outPath()
      // a plain formula fills as before, silently
      noArray <- run(src, plain, "fill", "D1", "D1:D4", "--json")
    yield
      ok(filled)
      assertEquals(overlapping(filled).size, 1, filled.stdout)
      assert(
        overlapping(filled).forall(m =>
          m.contains("Sheet1!F1: the dynamic array =SORT(D1:D3) last spilled F1:F3 (3 rows)") &&
            m.contains("every copy but the last shows #SPILL!") && m.contains("#737")
        ),
        filled.stdout
      )
      // pinned (#737): each copy is written as a 1x1 dynamic record with its own top-left value
      assertEquals(f6, Some(FormulaKind.dynamicArray(CellRange(ref"F6", ref"F6"))))
      assertRefused(intoOwn, "FORMULA_ERROR", "F2, F3: they are in the spill range F1:F3")
      ok(right)
      assertEquals(overlapping(right), Vector.empty, right.stdout)
      ok(noArray)
      assertEquals(overlapping(noArray), Vector.empty, noArray.stdout)
  }
