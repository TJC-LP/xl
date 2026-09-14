package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cli.commands.SheetCommands
import com.tjclp.xl.cli.contract.CliHarness
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

import munit.FunSuite

/**
 * PR #659 review: `xl name add|rm` must leave the same caches as their batch twins `define-name` /
 * `remove-name` ([[BatchNameRecalcSpec]]) — a file written by the verb must not carry a cached
 * `<v>` that contradicts its own defined-name table. Same helper, same `--no-recalc` posture, same
 * "Recalculated N formula(s)" summary shape.
 */
class NameVerbRecalcSpec extends FunSuite:
  given unsafe.IORuntime = unsafe.IORuntime.global

  private def formula(text: String, cache: Int): CellValue =
    CellValue.Formula(text, Some(CellValue.Number(cache)))

  private def withTemp[A](f: Path => A): A =
    val out = Files.createTempFile("name-verb-recalc", ".xlsx")
    try f(out)
    finally Files.deleteIfExists(out)

  private def add(
    wb: Workbook,
    name: String,
    refersTo: String,
    scope: Option[SheetName] = None,
    policy: WritePolicy = WritePolicy.default
  ): (Workbook, String) =
    withTemp { out =>
      val summary = SheetCommands
        .nameAdd(wb, scope, name, refersTo, out, WriterConfig.default, false, policy)
        .unsafeRunSync()
      (ExcelIO.instance[IO].read(out).unsafeRunSync(), summary)
    }

  private def remove(
    wb: Workbook,
    name: String,
    scope: Option[SheetName] = None,
    policy: WritePolicy = WritePolicy.default
  ): (Workbook, String) =
    withTemp { out =>
      val summary = SheetCommands
        .nameRemove(wb, scope, name, out, WriterConfig.default, false, policy)
        .unsafeRunSync()
      (ExcelIO.instance[IO].read(out).unsafeRunSync(), summary)
    }

  /** `Tax` = 0.1; A1 reads it (cached 10), B1 reads A1 (cached 20), D1 reads another name. */
  private val book: Workbook = Workbook(
    Sheet("Data")
      .put(ref"A1", formula("Tax*100", 10))
      .put(ref"B1", formula("A1*2", 20))
      .put(ref"D1", formula("OtherRate*10", 999))
  ).withDefinedName("Tax", "0.1").withDefinedName("OtherRate", "0.5")

  test("name add replacing a name's binding refreshes its readers and their dependents") {
    val (result, summary) = add(book, "Tax", "0.2")
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(20), summary)
    assertEquals(result.sheets(0)(ref"B1").effectiveValue, CellValue.Number(40), summary)
    // an unrelated name's reader keeps its cache
    assertEquals(result.sheets(0)(ref"D1").effectiveValue, CellValue.Number(999), summary)
    assertEquals(
      summary.linesIterator.toList.take(2),
      List("Added named range 'Tax' -> 0.2", "Recalculated 2 formulas")
    )
  }

  test("name add under --no-recalc keeps every original cache and says so") {
    val (result, summary) = add(book, "Tax", "0.2", policy = WritePolicy(noRecalc = true))
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(10), summary)
    assertEquals(result.sheets(0)(ref"B1").effectiveValue, CellValue.Number(20), summary)
    assert(
      summary.contains(
        "Recalculation skipped (--no-recalc): every existing cached value preserved"
      ),
      summary
    )
    assert(!summary.contains("Recalculated"), summary)
  }

  test("name rm withdraws the caches of the removed name's readers and their dependents") {
    val (result, summary) = remove(book, "Tax")
    for ref <- List(ref"A1", ref"B1") do
      result.sheets(0)(ref).value match
        case CellValue.Formula(_, cache, _) => assertEquals(cache, None, summary)
        case other => fail(s"expected formula at $ref, got $other")
    assertEquals(result.sheets(0)(ref"D1").effectiveValue, CellValue.Number(999), summary)
    assert(summary.startsWith("Removed named range 'Tax'\nRecalculated "), summary)
  }

  test("name rm under --no-recalc keeps the stale caches, as the batch twin does") {
    val (result, summary) = remove(book, "Tax", policy = WritePolicy(noRecalc = true))
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(10), summary)
    assert(summary.contains("Recalculation skipped (--no-recalc)"), summary)
  }

  test("name add -s of a local name refreshes only that sheet's readers") {
    val wb = Workbook(
      Sheet("Data").put(ref"A1", formula("Tax*100", 10)),
      Sheet("Local").put(ref"A1", formula("Tax*100", 10))
    ).withDefinedName("Tax", "0.1")
    val (result, summary) = add(wb, "Tax", "0.5", scope = Some(SheetName.unsafe("Local")))
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(10), summary)
    assertEquals(result.sheets(1)(ref"A1").effectiveValue, CellValue.Number(50), summary)
  }

  test("a name add that changes nothing recalculates nothing and prints no recalc line") {
    val (result, summary) = add(book, "Tax", "0.1")
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(10), summary)
    assertEquals(summary.linesIterator.toList.head, "Added named range 'Tax' -> 0.1")
    assert(!summary.contains("Recalculated"), summary)
  }

  test("the CLI verb prints the batch twin's summary shape and writes the refreshed cache") {
    val dir = Files.createTempDirectory("name-verb-cli")
    try
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      ExcelIO.instance[IO].write(book, in).unsafeRunSync()
      val run = CliHarness
        .run("-f", in.toString, "-o", out.toString, "name", "add", "Tax", "0.2")
        .unsafeRunSync()
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(
        run.stdout.linesIterator.toList.take(2),
        List("Added named range 'Tax' -> 0.2", "Recalculated 2 formulas")
      )
      val written = ExcelIO.instance[IO].read(out).unsafeRunSync()
      assertEquals(written.sheets(0)(ref"A1").effectiveValue, CellValue.Number(20))
      // the twin, byte-for-byte the same caches
      val twinOut = dir.resolve("twin.xlsx")
      val twin = CliHarness
        .run(
          List("-f", in.toString, "-o", twinOut.toString, "batch", "-"),
          """[{"op":"define-name","name":"Tax","refersTo":"0.2"}]"""
        )
        .unsafeRunSync()
      assertEquals(twin.exit, 0, twin.stderr)
      assert(twin.stdout.contains("Recalculated 2 formulas"), twin.stdout)
      val twinWb = ExcelIO.instance[IO].read(twinOut).unsafeRunSync()
      assertEquals(twinWb.sheets(0).cells, written.sheets(0).cells)
    finally
      val entries = Files.list(dir)
      try entries.forEach(p => Files.deleteIfExists(p))
      finally entries.close()
      Files.deleteIfExists(dir)
  }

  test("name rm --strict exits 1 when a reader is left uncached, like the batch twin") {
    val dir = Files.createTempDirectory("name-verb-strict")
    try
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      ExcelIO.instance[IO].write(book, in).unsafeRunSync()
      val run = CliHarness
        .run("-f", in.toString, "-o", out.toString, "--strict", "--json", "name", "rm", "Tax")
        .unsafeRunSync()
      assertEquals(run.exit, 1, run.stdout)
      val envelope = ujson.read(run.stdout)
      assertEquals(envelope("error")("code"), ujson.Str("RECALC_GATE"))
    finally
      val entries = Files.list(dir)
      try entries.forEach(p => Files.deleteIfExists(p))
      finally entries.close()
      Files.deleteIfExists(dir)
  }
