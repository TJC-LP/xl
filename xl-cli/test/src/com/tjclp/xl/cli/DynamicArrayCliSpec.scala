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
