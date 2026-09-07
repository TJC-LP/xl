package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite
import munit.ScalaCheckSuite
import org.scalacheck.Prop.forAll

import com.tjclp.xl.{*, given}
import com.tjclp.xl.Generators.genSheetName
import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.cli.helpers.Resolve
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.metadata.{LightMetadata, SheetInfo}

/**
 * THE sheet rule (ADR-017 §2.5), as one implementation and as the CLI's behaviour: a qualified ref
 * names the sheet; else `-s` (for a batch op: its `sheet` key, then `-s`); else the only sheet of a
 * single-sheet book — announced as `SHEET_AUTOSELECTED` under `--json` only; else `SHEET_REQUIRED`
 * (exit 3) with the sheet names as candidates. Every verb, batch op and streaming path.
 */
class ResolveSpec extends CatsEffectSuite with ScalaCheckSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "resolve-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString

  private val data = SheetName.unsafe("Data")
  private val summary = SheetName.unsafe("Summary")
  private val two: Workbook = TestFixtures.simpleBook()
  private val one: Workbook = TestFixtures.singleSheetBook()

  private def meta(names: SheetName*): LightMetadata =
    LightMetadata(
      names.zipWithIndex.map((n, i) => SheetInfo(n, i + 1, None, None)).toVector,
      Vector.empty
    )

  private val metaTwo = meta(data, summary)
  private val metaOne = meta(SheetName.unsafe("Sheet1"))

  private def sheetOf(result: Either[CliError, Resolve.Resolved]): String =
    result.fold(e => fail(s"expected a sheet, got $e"), _.sheet.name.value)

  private def error[A](result: Either[CliError, A]): CliError =
    result.fold(identity, a => fail(s"expected a failure, got $a"))

  // ---------------------------------------------------------------------------------------------
  // target: the four steps
  // ---------------------------------------------------------------------------------------------

  test("step 1: a qualified ref names the sheet, whatever -s says") {
    val resolved = Resolve.target(two, Some("Summary"), "Data!A1", "view")
    assertEquals(sheetOf(resolved), "Data")
    assertEquals(resolved.map(_.target), Right(Resolve.Target.Cell(ref"A1")))
    assertEquals(resolved.map(_.viaQualifiedRef), Right(true))
    assertEquals(resolved.map(_.autoSelected), Right(false))
    assertEquals(
      Resolve.target(two, None, "'Summary'!A1:B2", "view").map(_.target),
      Right(Resolve.Target.Range(CellRange(ref"A1", ref"B2")))
    )
  }

  property("step 1 holds for every -s value, existing or not") {
    forAll(genSheetName) { (flag: SheetName) =>
      sheetOf(Resolve.target(two, Some(flag.value), "Data!A1:B2", "view")) == "Data"
    }
  }

  test("step 2: -s selects the sheet of an unqualified ref") {
    val resolved = Resolve.target(two, Some("Summary"), "A1", "view")
    assertEquals(sheetOf(resolved), "Summary")
    assertEquals(resolved.map(_.viaQualifiedRef), Right(false))
    assertEquals(resolved.map(_.autoSelected), Right(false))
    assertEquals(Resolve.sheet(two, Some("Summary"), "bounds").map(_.name), Right(summary))
  }

  test("step 3: the only sheet of a single-sheet book is selected, and says so") {
    val resolved = Resolve.target(one, None, "A1", "cell")
    assertEquals(sheetOf(resolved), "Sheet1")
    assertEquals(resolved.map(_.autoSelected), Right(true))
    assertEquals(Resolve.sheet(one, None, "bounds").map(_.name.value), Right("Sheet1"))
  }

  test("step 4: a multi-sheet book without a sheet is SHEET_REQUIRED, exit 3, with candidates") {
    val cell = error(Resolve.target(two, None, "A1", "cell"))
    assertEquals(cell.code, "SHEET_REQUIRED")
    assertEquals(cell.exitCode.code, 3)
    assertEquals(
      cell.message,
      "cell with unqualified ref 'A1' requires --sheet or qualified ref (e.g., Sheet1!A1). Available sheets: Data, Summary"
    )
    assertEquals(cell.candidates, Vector("Data", "Summary"))
    assertEquals(cell.hint, Some("use -s <name> or a qualified ref like 'Name'!A1"))

    val range = error(Resolve.target(two, None, "A1:B2", "view"))
    assertEquals(
      range.message,
      "view with unqualified range 'A1:B2' requires --sheet or qualified ref (e.g., Sheet1!A1). Available sheets: Data, Summary"
    )
    val bare = error(Resolve.sheet(two, None, "bounds"))
    assertEquals(
      bare.message,
      "bounds requires --sheet or qualified ref (e.g., Sheet1!A1). Available sheets: Data, Summary"
    )
    assertEquals(bare.code, "SHEET_REQUIRED")
  }

  test("an unknown -s or qualifier is SHEET_NOT_FOUND with the legacy text and a did-you-mean") {
    val flag = error(Resolve.sheet(two, Some("Dat"), "view"))
    assertEquals(flag.code, "SHEET_NOT_FOUND")
    assertEquals(flag.message, "Sheet not found: Dat. Available: Data, Summary")
    assertEquals(flag.candidates, Vector("Data"))
    val qualified = error(Resolve.target(two, None, "Dat!A1", "view"))
    assertEquals(qualified.code, "SHEET_NOT_FOUND")
    assertEquals(qualified.message, "Sheet not found: Dat. Available: Data, Summary")
  }

  test("an invalid -s is INVALID_SHEET_NAME; an unparseable ref is INVALID_REFERENCE") {
    assertEquals(error(Resolve.sheet(two, Some("Bad:Name"), "view")).code, "INVALID_SHEET_NAME")
    assertEquals(
      error(Resolve.target(two, Some("Data"), "not a ref", "view")).code,
      "INVALID_REFERENCE"
    )
  }

  test("default: -s by name; the only sheet only for a verb that takes a sheet; else None") {
    assertEquals(
      Resolve.default(two, Some("Data"), takesSheet = false).map(_.map(_.name)),
      Right(Some(data))
    )
    assertEquals(Resolve.default(two, None, takesSheet = true), Right(None))
    assertEquals(
      Resolve.default(one, None, takesSheet = true).map(_.map(_.name.value)),
      Right(Some("Sheet1"))
    )
    assertEquals(Resolve.default(one, None, takesSheet = false), Right(None))
    assertEquals(
      error(Resolve.default(two, Some("Nope"), takesSheet = true)).code,
      "SHEET_NOT_FOUND"
    )
  }

  // ---------------------------------------------------------------------------------------------
  // sheetName: the same rule over metadata (the streaming paths)
  // ---------------------------------------------------------------------------------------------

  test("sheetName follows the four steps over LightMetadata; a qualifier beats -s") {
    assertEquals(Resolve.sheetName(metaTwo, Some("Summary"), Some(data), "view"), Right(data))
    assertEquals(Resolve.sheetName(metaTwo, None, Some(data), "view"), Right(data))
    assertEquals(Resolve.sheetName(metaTwo, Some("Summary"), None, "view"), Right(summary))
    assertEquals(Resolve.sheetName(metaOne, None, None, "view").map(_.value), Right("Sheet1"))
    val required = error(Resolve.sheetName(metaTwo, None, None, "view"))
    assertEquals(required.code, "SHEET_REQUIRED")
    assertEquals(required.candidates, Vector("Data", "Summary"))
    assertEquals(
      required.message,
      "view requires --sheet or qualified ref (e.g., Sheet1!A1). Available sheets: Data, Summary"
    )
    assertEquals(
      error(Resolve.sheetName(metaTwo, Some("Dat"), None, "view")).code,
      "SHEET_NOT_FOUND"
    )
    assertEquals(
      error(Resolve.sheetName(metaTwo, None, Some(SheetName.unsafe("Dat")), "view")).code,
      "SHEET_NOT_FOUND"
    )
  }

  // ---------------------------------------------------------------------------------------------
  // forOp: one batch op
  // ---------------------------------------------------------------------------------------------

  test("forOp: qualified ref > op sheet > default > the only sheet > SHEET_REQUIRED") {
    assertEquals(Resolve.forOp(two, Some(summary), None, Some(data), 1, "put"), Right(data))
    assertEquals(Resolve.forOp(two, Some(summary), Some("Data"), None, 1, "put"), Right(data))
    assertEquals(Resolve.forOp(two, Some(summary), None, None, 1, "put"), Right(summary))
    assertEquals(Resolve.forOp(one, None, None, None, 1, "put").map(_.value), Right("Sheet1"))
    val required = error(Resolve.forOp(two, None, None, None, 2, "rowheight"))
    assertEquals(required.code, "SHEET_REQUIRED")
    assertEquals(
      required.message,
      "batch rowheight requires --sheet or qualified ref (e.g., Sheet1!A1). Available sheets: Data, Summary"
    )
    assertEquals(required.candidates, Vector("Data", "Summary"))
    assertEquals(required.location.flatMap(_.opIndex), Some(2))
  }

  property("forOp: a qualified ref wins over any default") {
    forAll(genSheetName) { (default: SheetName) =>
      Resolve.forOp(two, Some(default), None, Some(data), 1, "put") == Right(data)
    }
  }

  test(
    "forOp: a qualified ref that disagrees with the op's sheet is BATCH_OP_INVALID at the index"
  ) {
    val conflict = error(Resolve.forOp(two, None, Some("Summary"), Some(data), 3, "put"))
    assertEquals(conflict.code, "BATCH_OP_INVALID")
    assertEquals(conflict.exitCode.code, 2)
    assert(conflict.message.startsWith("Object 3 (put): "), conflict.message)
    assert(conflict.message.contains("names sheet 'Data'"), conflict.message)
    assert(conflict.message.contains("\"sheet\" is 'Summary'"), conflict.message)
    assertEquals(conflict.location.flatMap(_.opIndex), Some(3))
    // agreement is fine
    assertEquals(Resolve.forOp(two, None, Some("Data"), Some(data), 3, "put"), Right(data))
    // an invalid sheet key is a malformed op too
    val invalid = error(Resolve.forOp(two, None, Some("Bad:Name"), None, 4, "merge"))
    assertEquals(invalid.code, "BATCH_OP_INVALID")
    assertEquals(invalid.location.flatMap(_.opIndex), Some(4))
  }

  test("the auto-select warning carries the code and the sheet") {
    val warning = Resolve.autoSelected(data)
    assertEquals(warning.code, WarningCode.SHEET_AUTOSELECTED)
    assert(warning.message.contains("Data"), warning.message)
    assertEquals(warning.location.flatMap(_.sheet), Some("Data"))
  }

  // ---------------------------------------------------------------------------------------------
  // Through the harness: every verb, in memory and under --stream
  // ---------------------------------------------------------------------------------------------

  private def assertSheetRequired(run: CliRun, label: String): Unit =
    assertEquals(run.exit, 3, s"$label: ${run.stderr}")
    assertEquals(run.stdout, "", label)
    assert(run.stderr.contains("  code: SHEET_REQUIRED"), s"$label: ${run.stderr}")
    assert(run.stderr.contains("  did you mean: Data, Summary"), s"$label: ${run.stderr}")

  test("single-sheet book: view, cell, stats, bounds, put and --stream view work without -s") {
    val single = file("single.xlsx")
    val out = file("single-put.xlsx")
    for
      view <- CliHarness.run("-f", single, "view", "A1:B1")
      cell <- CliHarness.run("-f", single, "cell", "A1")
      stats <- CliHarness.run("-f", single, "stats", "A1:B1")
      bounds <- CliHarness.run("-f", single, "bounds")
      put <- CliHarness.run("-f", single, "-o", out, "put", "A1", "5")
      streamView <- CliHarness.run("-f", single, "--stream", "view", "A1:B1")
      streamCell <- CliHarness.run("-f", single, "--stream", "cell", "B1")
      streamBounds <- CliHarness.run("-f", single, "--stream", "bounds")
    yield
      List(view, cell, stats, bounds, put, streamView, streamCell, streamBounds).foreach { run =>
        assertEquals(run.exit, 0, run.stderr)
        assertEquals(run.stderr, "", "text mode: no auto-select notice on stderr")
      }
      assert(view.stdout.contains("solo"), view.stdout)
      assert(streamView.stdout.contains("| solo | 7 |"), streamView.stdout)
      assert(cell.stdout.contains("solo"), cell.stdout)
      assert(streamCell.stdout.contains("7"), streamCell.stdout)
      assert(stats.stdout.startsWith("count: 1"), stats.stdout)
      assert(bounds.stdout.contains("Sheet1"), bounds.stdout)
      assert(put.stdout.contains("Put: A1 = 5"), put.stdout)
      assert(Files.exists(Path.of(out)))
  }

  test(
    "two-sheet book: the same verbs are SHEET_REQUIRED with candidates, in memory and streaming"
  ) {
    val simple = file("simple.xlsx")
    val out = file("two-sheet-put.xlsx")
    for
      view <- CliHarness.run("-f", simple, "view", "A1:B2")
      cell <- CliHarness.run("-f", simple, "cell", "A1")
      stats <- CliHarness.run("-f", simple, "stats", "B1:B3")
      bounds <- CliHarness.run("-f", simple, "bounds")
      put <- CliHarness.run("-f", simple, "-o", out, "put", "A1", "5")
      streamView <- CliHarness.run("-f", simple, "--stream", "view", "A1:B2")
      streamCell <- CliHarness.run("-f", simple, "--stream", "cell", "A1")
      streamStats <- CliHarness.run("-f", simple, "--stream", "stats", "B1:B3")
      streamBounds <- CliHarness.run("-f", simple, "--stream", "bounds")
      streamPut <- CliHarness.run("-f", simple, "-o", out, "--stream", "put", "A1", "5")
    yield
      assertSheetRequired(view, "view")
      assertSheetRequired(cell, "cell")
      assertSheetRequired(stats, "stats")
      assertSheetRequired(bounds, "bounds")
      assertSheetRequired(put, "put")
      assertSheetRequired(streamView, "--stream view")
      assertSheetRequired(streamCell, "--stream cell")
      assertSheetRequired(streamStats, "--stream stats")
      assertSheetRequired(streamBounds, "--stream bounds")
      assertSheetRequired(streamPut, "--stream put")
      assert(!Files.exists(Path.of(out)), "nothing written")
  }

  test("--stream writes honour a qualified ref: `--stream put Data!A1 5` selects Data") {
    val out = file("stream-qualified-put.xlsx")
    val outf = file("stream-qualified-putf.xlsx")
    for
      put <- CliHarness.run("-f", file("simple.xlsx"), "-o", out, "--stream", "put", "Data!A5", "5")
      cell <- CliHarness.run("-f", out, "cell", "Data!A5")
      putf <- CliHarness
        .run("-f", file("simple.xlsx"), "-o", outf, "--stream", "putf", "'Summary'!A1", "=1+1")
      formula <- CliHarness.run("-f", outf, "cell", "Summary!A1")
    yield
      assertEquals(put.exit, 0, put.stderr)
      assertEquals(cell.exit, 0, cell.stderr)
      assert(cell.stdout.contains("5"), cell.stdout)
      assertEquals(putf.exit, 0, putf.stderr)
      assert(formula.stdout.contains("=1+1"), formula.stdout)
  }

  test("--stream view with -s and a different qualified sheet uses the qualified one") {
    for
      qualified <- CliHarness
        .run("-f", file("simple.xlsx"), "-s", "Summary", "--stream", "view", "Data!A1:B2")
      plain <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "--stream", "view", "A1:B2")
      inMemory <- CliHarness
        .run("-f", file("simple.xlsx"), "-s", "Summary", "view", "Data!A1:B2")
      inMemoryPlain <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "view", "A1:B2")
    yield
      assertEquals(qualified.exit, 0, qualified.stderr)
      assertEquals(qualified.stdout, plain.stdout)
      assertEquals(inMemory.stdout, inMemoryPlain.stdout)
      assert(qualified.stdout.contains("Hello"), qualified.stdout)
  }

  test("--stream with an unknown -s is SHEET_NOT_FOUND on reads and writes") {
    val out = file("stream-nope.xlsx")
    for
      view <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Nope", "--stream", "view", "A1:B2")
      put <- CliHarness
        .run("-f", file("simple.xlsx"), "-s", "Nope", "-o", out, "--stream", "put", "A1", "1")
    yield
      assertEquals(view.exit, 3, view.stderr)
      assert(view.stderr.contains("  code: SHEET_NOT_FOUND"), view.stderr)
      assertEquals(put.exit, 3, put.stderr)
      assert(put.stderr.contains("  code: SHEET_NOT_FOUND"), put.stderr)
      assert(!Files.exists(Path.of(out)))
  }

  test("SHEET_AUTOSELECTED appears in warnings[] under --json only, and only when step 3 applied") {
    val single = file("single.xlsx")
    def codes(run: CliRun): Vector[String] =
      ujson.read(run.stdout)("warnings").arr.map(_("code").str).toVector
    for
      auto <- CliHarness.run("-f", single, "--json", "view", "A1:B1")
      streamAuto <- CliHarness.run("-f", single, "--json", "--stream", "view", "A1:B1")
      cellAuto <- CliHarness.run("-f", single, "--json", "cell", "A1")
      boundsAuto <- CliHarness.run("-f", single, "--json", "bounds")
      flagged <- CliHarness.run("-f", single, "-s", "Sheet1", "--json", "view", "A1:B1")
      search <- CliHarness.run("-f", single, "--json", "search", "solo")
      sheets <- CliHarness.run("-f", single, "--json", "sheets")
      text <- CliHarness.run("-f", single, "view", "A1:B1")
      // step 1 decided these: the default sheet never mattered, so nothing was auto-selected
      qualifiedView <- CliHarness.run("-f", single, "--json", "view", "Sheet1!A1:B1")
      qualifiedCell <- CliHarness.run("-f", single, "--json", "cell", "Sheet1!A1")
      qualifiedStreamPut <- CliHarness.run(
        "-f",
        single,
        "-o",
        file("single-qualified-stream-put.xlsx"),
        "--json",
        "--stream",
        "put",
        "Sheet1!A1",
        "9"
      )
      qualifiedEval <- CliHarness.run("-f", single, "--json", "eval", "=Sheet1!B1*2")
      unqualifiedEval <- CliHarness.run("-f", single, "--json", "eval", "=B1*2")
    yield
      List(auto, streamAuto, cellAuto, boundsAuto, unqualifiedEval).foreach { run =>
        assertEquals(run.exit, 0, run.stderr)
        assertEquals(codes(run), Vector(WarningCode.SHEET_AUTOSELECTED), run.stdout)
        assertEquals(run.stderr, "", "the notice rides in the envelope, not on stderr")
      }
      assertEquals(codes(flagged), Vector.empty, flagged.stdout)
      assertEquals(codes(search), Vector.empty, "search reads every sheet: no auto-select")
      assertEquals(codes(sheets), Vector.empty, "a workbook verb: no auto-select")
      List(qualifiedView, qualifiedCell, qualifiedStreamPut, qualifiedEval).foreach { run =>
        assertEquals(run.exit, 0, run.stderr)
        assertEquals(codes(run), Vector.empty, s"qualified: no auto-select in ${run.stdout}")
      }
      assertEquals(text.exit, 0)
      assertEquals(text.stderr, "")
  }

  test("import into an existing sheet follows the rule; --new-sheet needs no sheet") {
    val csv = fixtures().resolve("import-rule.csv")
    val outSingle = file("import-single.xlsx")
    val outTwo = file("import-two.xlsx")
    val outNew = file("import-new.xlsx")
    for
      _ <- IO.blocking(Files.writeString(csv, "name,qty\nwidget,3\n"))
      single <- CliHarness.run("-f", file("single.xlsx"), "-o", outSingle, "import", csv.toString)
      two <- CliHarness.run("-f", file("simple.xlsx"), "-o", outTwo, "import", csv.toString)
      newSheet <- CliHarness.run(
        "-f",
        file("simple.xlsx"),
        "-o",
        outNew,
        "--json",
        "import",
        csv.toString,
        "--new-sheet",
        "Imported"
      )
      cell <- CliHarness.run("-f", outSingle, "cell", "A2")
    yield
      assertEquals(single.exit, 0, single.stderr)
      assert(cell.stdout.contains("widget"), cell.stdout)
      assertSheetRequired(two, "import")
      assert(two.stderr.contains("Error: import requires --sheet"), two.stderr)
      assert(!Files.exists(Path.of(outTwo)))
      assertEquals(newSheet.exit, 0, newSheet.stderr)
      assertEquals(
        ujson.read(newSheet.stdout)("warnings").arr.size,
        0,
        "--new-sheet needs no default: nothing auto-selected"
      )
  }

  test("a batch op qualified with the renamed default resolves by the new identity") {
    val out = file("batch-renamed.xlsx")
    val ops =
      """[{"op":"rename-sheet","from":"Data","to":"Sheet2"},{"op":"put","ref":"Sheet2!A9","value":"renamed"}]"""
    for
      batch <- CliHarness.run(
        List("-f", file("simple.xlsx"), "-s", "Data", "-o", out, "batch", "-"),
        ops
      )
      cell <- CliHarness.run("-f", out, "cell", "Sheet2!A9")
    yield
      assertEquals(batch.exit, 0, batch.stderr)
      assertEquals(cell.exit, 0, cell.stderr)
      assert(cell.stdout.contains("renamed"), cell.stdout)
  }

  test(
    "batch ops on a single-sheet book need neither -s nor sheet; on two sheets they are SHEET_REQUIRED"
  ) {
    val single = file("single.xlsx")
    val outSingle = file("batch-single.xlsx")
    val outTwo = file("batch-two.xlsx")
    val ops = """[{"op":"put","ref":"A5","value":1},{"op":"rowheight","row":2,"height":30}]"""
    for
      ok <- CliHarness.run(List("-f", single, "-o", outSingle, "batch", "-"), ops)
      required <- CliHarness.run(List("-f", file("simple.xlsx"), "-o", outTwo, "batch", "-"), ops)
    yield
      assertEquals(ok.exit, 0, ok.stderr)
      assert(Files.exists(Path.of(outSingle)))
      assertEquals(required.exit, 3, required.stderr)
      assert(required.stderr.contains("  code: BATCH_OP_FAILED"), required.stderr)
      assert(
        required.stderr.contains("Object 1 (put): batch put requires --sheet"),
        required.stderr
      )
      assert(required.stderr.contains("  did you mean: Data, Summary"), required.stderr)
      assert(!Files.exists(Path.of(outTwo)))
  }
