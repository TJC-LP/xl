package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{IO, Ref, Resource}
import fs2.Stream
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cli.contract.{
  CliHarness,
  CliRun,
  Outcome,
  OutputMode,
  Payload,
  StreamedBody,
  TestFixtures,
  Warning,
  WarningCode
}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.{SstPolicy, WriterConfig}

/**
 * The streaming contract of the 0.21.1 fixes, end to end through the harness.
 *
 *   - GH-635: `--stream view --limit 0` streams — the same bytes as the bounded and the in-memory
 *     runs, for csv, json (bare and in the envelope) and markdown; `--limit -1` is a usage error; a
 *     failure while the rows stream is the run's failure after whatever was written; a reader that
 *     closes stdout (`| head`) ends the run quietly.
 *   - GH-638: every verb answers `--stream` from one table — `sheets --stats` is refused before any
 *     read like `describe --full`; `names` and `lint` run under it; `diff`, `eval`, `schema` and
 *     the other refusals are `UNSUPPORTED_IN_STREAM` and never decline's "Unexpected argument"; the
 *     reader's size limit is `SECURITY_ERROR` with the `--max-size`/`--stream` hint, not `IO_READ`.
 *   - GH-640: the shared-string table is held to `--max-size` under `--stream`, with its own hint.
 */
class StreamingContractSpec extends CatsEffectSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "streaming-contract-fixtures",
    Resource.make(TestFixtures.materialize.flatTap(extras))(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString

  /** The harness over an argv built from pieces. */
  private def run(args: Vector[String]): IO[CliRun] = CliHarness.run(args.toList, "")

  /**
   * `big.xlsx`: 40,000 distinct strings, so the package inflates past 1 MB and its shared-string
   * table alone does too — the smallest `--max-size` the flag accepts.
   */
  private def extras(dir: Path): IO[Unit] =
    val sheet = (0 until 40000).foldLeft(Sheet("Data")) { (s, i) =>
      s.put(ARef.from0(i % 8, i / 8), f"distinct-text-$i%06d-padding-padding")
    }
    ExcelIO
      .instance[IO]
      .writeWith(
        Workbook(Vector(sheet)),
        dir.resolve("big.xlsx"),
        WriterConfig(sstPolicy = SstPolicy.Always)
      )

  private def envelope(run: CliRun): ujson.Value = ujson.read(run.stdout)

  private def assertRefused(run: CliRun, messageStart: String): Unit =
    assertEquals(run.exit, 2, run.stderr)
    assertEquals(run.stdout, "")
    val first = run.stderr.linesIterator.toVector.headOption.getOrElse("")
    assert(first.startsWith(s"Error: $messageStart"), run.stderr)
    assert(run.stderr.contains("  code: UNSUPPORTED_IN_STREAM"), run.stderr)
    assert(run.stderr.contains("  hint: "), run.stderr)
    assert(run.stderr.contains("omit --stream"), run.stderr)

  // ---------------------------------------------------------------------------------------------
  // GH-635: an unbounded streamed view
  // ---------------------------------------------------------------------------------------------

  test("--stream view --limit 0 prints what the bounded run prints and what memory prints: csv") {
    val base = Vector("-f", file("simple.xlsx"), "-s", "Data")
    for
      streamed <- run(base ++ Vector("--stream", "view", "--limit", "0", "--format", "csv"))
      bounded <- run(base ++ Vector("--stream", "view", "A1:C4", "--format", "csv"))
      memory <- run(base ++ Vector("view", "--limit", "0", "--format", "csv"))
    yield
      assertEquals(streamed.exit, 0, streamed.stderr)
      assertEquals(streamed.stdout, bounded.stdout)
      assertEquals(streamed.stdout, memory.stdout)
      assertEquals(streamed.stderr, "")
      assertEquals(streamed.stdout.linesIterator.size, 4)
  }

  test("--stream view --limit 0: json bare and in the envelope, and markdown, all as in memory") {
    val base = Vector("-f", file("simple.xlsx"), "-s", "Data")
    for
      json <- run(base ++ Vector("--stream", "view", "--limit", "0", "--format", "json"))
      jsonMemory <- run(base ++ Vector("view", "--limit", "0", "--format", "json"))
      wrapped <- run(
        base ++ Vector("--json", "--stream", "view", "--limit", "0", "--format", "json")
      )
      wrappedMemory <- run(
        base ++ Vector("--json", "view", "--limit", "0", "--format", "json")
      )
      markdown <- run(base ++ Vector("--stream", "view", "--limit", "0"))
      markdownMemory <- run(base ++ Vector("view", "--limit", "0"))
      csvWrapped <- run(
        base ++ Vector("--json", "--stream", "view", "--limit", "0", "--format", "csv")
      )
      csvWrappedMemory <- run(
        base ++ Vector("--json", "view", "--limit", "0", "--format", "csv")
      )
    yield
      assertEquals(json.stdout, jsonMemory.stdout)
      assertEquals(ujson.read(json.stdout)("rows").arr.size, 4)
      assertEquals(wrapped.stdout, wrappedMemory.stdout)
      val e = envelope(wrapped)
      assertEquals(e("ok"), ujson.True)
      assertEquals(e("data")("rows").arr.size, 4)
      assertEquals(markdown.stdout, markdownMemory.stdout)
      assert(markdown.stdout.startsWith("|   | A     |"), markdown.stdout)
      assertEquals(csvWrapped.stdout, csvWrappedMemory.stdout)
      assert(envelope(csvWrapped)("data")("text").str.startsWith("Hello,10"), csvWrapped.stdout)
  }

  test("view --limit -1 is a usage error (exit 2), not silently no limit") {
    for
      memory <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "view", "--limit", "-1")
      streamed <- CliHarness.run(
        "-f",
        file("simple.xlsx"),
        "-s",
        "Data",
        "--stream",
        "--json",
        "view",
        "--limit",
        "-1"
      )
    yield
      assertEquals(memory.exit, 2, memory.stderr)
      assertEquals(memory.stdout, "")
      assert(memory.stderr.startsWith("Error: --limit must be 0 or more (got -1)"), memory.stderr)
      assert(memory.stderr.contains("  code: USAGE"), memory.stderr)
      assertEquals(streamed.exit, 2)
      assertEquals(envelope(streamed)("error")("code"), ujson.Str("USAGE"))
  }

  test("a failure while the rows stream stops the output and is the run's failure") {
    val rows = Stream.emits(Vector("a,1", "b,2")).covary[IO] ++
      Stream.raiseError[IO](new IllegalStateException("the reader gave up"))
    val outcome = Outcome.ok(
      "view",
      Payload.Streamed(StreamedBody.Lines(Vector(",A,B"), rows, Vector.empty)),
      Vector(Warning(WarningCode.TRUNCATED, "… showing 2 of 4 rows"))
    )
    for
      captured <- CliIO.capturing("")
      (io, collect) = captured
      exit <- Main.emit(outcome, OutputMode.Text, io)
      (out, err) <- collect
      capturedJson <- CliIO.capturing("")
      (ioJson, collectJson) = capturedJson
      exitJson <- Main.emit(outcome, OutputMode.Json, ioJson)
      (outJson, errJson) <- collectJson
    yield
      assertEquals(exit.code, 3)
      // the failed chunk is never written: the lines before it were the whole chunk
      assertEquals(out, "")
      assert(err.startsWith("Error: the reader gave up"), err)
      assert(err.contains("  code: INTERNAL"), err)
      assert(err.contains("Warning[TRUNCATED]"), err)
      assertEquals(exitJson.code, 3)
      assertEquals(outJson, "")
      assertEquals(errJson.trim, "Error: the reader gave up")
  }

  test("a reader that closes stdout ends a streamed run quietly with exit 0 (xl … | head)") {
    val rows = Stream.range(0, 2000).map(i => s"row $i").covary[IO]
    val outcome =
      Outcome.ok("view", Payload.Streamed(StreamedBody.Lines(Vector.empty, rows, Vector.empty)))
    for
      writes <- Ref.of[IO, Int](0)
      captured <- CliIO.capturing("")
      (io, collect) = captured
      // the pipe closes after the first write
      closing = io.copy(
        out = line => io.out(line) *> writes.update(_ + 1),
        outFailed = writes.get.map(_ >= 1)
      )
      exit <- Main.emit(outcome, OutputMode.Text, closing)
      (out, err) <- collect
      written <- writes.get
    yield
      assertEquals(exit.code, 0)
      assertEquals(err, "")
      assertEquals(written, 1)
      assertEquals(out.linesIterator.size, 256)
  }

  // ---------------------------------------------------------------------------------------------
  // GH-638: every verb's --stream answer comes from one table
  // ---------------------------------------------------------------------------------------------

  test("--stream sheets --stats is refused before any read, like describe --full") {
    for
      text <- CliHarness.run("-f", file("simple.xlsx"), "--stream", "sheets", "--stats")
      json <- CliHarness.run("-f", file("simple.xlsx"), "--stream", "--json", "sheets", "--stats")
      quick <- CliHarness.run("-f", file("simple.xlsx"), "--stream", "sheets")
      plain <- CliHarness.run("-f", file("simple.xlsx"), "sheets")
    yield
      assertRefused(text, "sheets --stats is not supported with --stream")
      assertEquals(json.exit, 2)
      assertEquals(envelope(json)("error")("code"), ujson.Str("UNSUPPORTED_IN_STREAM"))
      assertEquals(quick.exit, 0, quick.stderr)
      assertEquals(quick.stdout, plain.stdout)
  }

  test("--stream names and --stream lint run, and print what the plain verbs print") {
    for
      names <- CliHarness.run("-f", file("named.xlsx"), "--stream", "names")
      plainNames <- CliHarness.run("-f", file("named.xlsx"), "names")
      namesJson <- CliHarness.run("-f", file("named.xlsx"), "--stream", "--json", "names")
      lint <- CliHarness.run("-f", file("simple.xlsx"), "--stream", "lint")
      plainLint <- CliHarness.run("-f", file("simple.xlsx"), "lint")
      lintPositional <- CliHarness.run("--stream", "lint", file("simple.xlsx"))
    yield
      assertEquals(names.exit, 0, names.stderr)
      assertEquals(names.stdout, plainNames.stdout)
      assert(names.stdout.contains("Total"), names.stdout)
      assertEquals(envelope(namesJson)("ok"), ujson.True)
      assertEquals(lint.exit, 0, lint.stderr)
      assertEquals(lint.stdout, plainLint.stdout)
      assertEquals(lintPositional.stdout, plainLint.stdout)
  }

  test(
    "verbs that refuse --stream say so with UNSUPPORTED_IN_STREAM, never 'Unexpected argument'"
  ) {
    for
      diff <- CliHarness.run(
        "-f",
        file("simple.xlsx"),
        "--stream",
        "diff",
        "-g",
        file("changed.xlsx")
      )
      eval <- CliHarness.run("--stream", "eval", "=1+1")
      schema <- CliHarness.run("--stream", "schema")
      audit <- CliHarness.run("-f", file("dirty.xlsx"), "--stream", "audit")
      deps <- CliHarness.run("-f", file("linked.xlsx"), "--stream", "--json", "deps", "Sheet2!A1")
      newBook <- CliHarness.run("--stream", "new", fixtures().resolve("never.xlsx").toString)
      help <- CliHarness.run("--stream", "audit", "--help")
    yield
      assertRefused(
        diff,
        "diff is not supported with --stream (the two workbooks are compared in memory)"
      )
      assert(diff.stderr.contains("--max-size"), diff.stderr)
      assertRefused(
        eval,
        "eval is not supported with --stream (formula evaluation needs the loaded workbook)"
      )
      assertRefused(schema, "schema is not supported with --stream (it reads no workbook)")
      assert(!schema.stderr.contains("--max-size"), schema.stderr)
      assertRefused(
        audit,
        "audit is not supported with --stream (the analysis needs the whole workbook)"
      )
      assertEquals(deps.exit, 2)
      val e = envelope(deps)
      assertEquals(e("error")("code"), ujson.Str("UNSUPPORTED_IN_STREAM"))
      assertEquals(e("verb"), ujson.Str("deps"))
      assertEquals(
        e("error")("message"),
        ujson.Str("deps is not supported with --stream (the graph needs the whole workbook)")
      )
      assertRefused(
        newBook,
        "new is not supported with --stream (it writes a new file and reads none)"
      )
      assert(!Files.exists(fixtures().resolve("never.xlsx")), "the refusal wrote nothing")
      // --help still helps
      assertEquals(help.exit, 0, help.stderr)
      assert(help.stderr.contains("Usage"), help.stderr)
  }

  test("the reader's size limit is SECURITY_ERROR (exit 3) with the --max-size/--stream hint") {
    for
      text <- CliHarness.run("-f", file("big.xlsx"), "--max-size", "1", "sheets", "--stats")
      json <- CliHarness.run("-f", file("big.xlsx"), "--max-size", "1", "--json", "view", "A1:B2")
      lifted <- CliHarness.run("-f", file("big.xlsx"), "--max-size", "0", "view", "A1:B2")
    yield
      assertEquals(text.exit, 3, text.stderr)
      assertEquals(text.stdout, "")
      assert(text.stderr.startsWith("Error: Security error: Total uncompressed size"), text.stderr)
      assert(text.stderr.contains("  code: SECURITY_ERROR"), text.stderr)
      assert(text.stderr.contains("  hint: re-run with --max-size 0 or --stream"), text.stderr)
      assertEquals(json.exit, 3)
      val e = envelope(json)
      assertEquals(e("error")("code"), ujson.Str("SECURITY_ERROR"))
      assertEquals(e("error")("location")("file"), ujson.Str(file("big.xlsx")))
      assertEquals(lifted.exit, 0, lifted.stderr)
  }

  // ---------------------------------------------------------------------------------------------
  // GH-640: the shared-string table under --stream
  // ---------------------------------------------------------------------------------------------

  test("--stream holds the shared-string table to --max-size: SECURITY_ERROR with its own hint") {
    for
      capped <- CliHarness.run(
        "-f",
        file("big.xlsx"),
        "--stream",
        "--max-size",
        "1",
        "view",
        "A1:B2"
      )
      cell <- CliHarness.run(
        "-f",
        file("big.xlsx"),
        "--stream",
        "--max-size",
        "1",
        "--json",
        "cell",
        "A1"
      )
      allowed <- CliHarness.run(
        "-f",
        file("big.xlsx"),
        "--stream",
        "view",
        "A1:B2",
        "--format",
        "csv"
      )
      lifted <- CliHarness.run("-f", file("big.xlsx"), "--stream", "--max-size", "0", "cell", "A1")
    yield
      assertEquals(capped.exit, 3, capped.stderr)
      assertEquals(capped.stdout, "")
      assert(
        capped.stderr.startsWith(
          "Error: Security error: Uncompressed size of 'xl/sharedStrings.xml'"
        ),
        capped.stderr
      )
      assert(capped.stderr.contains("  code: SECURITY_ERROR"), capped.stderr)
      assert(capped.stderr.contains("  hint: raise --max-size <MB> (0 = unlimited)"), capped.stderr)
      assertEquals(cell.exit, 3)
      assertEquals(envelope(cell)("error")("code"), ujson.Str("SECURITY_ERROR"))
      // the default 100 MB admits the table; the streamed read then works as it always has
      assertEquals(allowed.exit, 0, allowed.stderr)
      assert(allowed.stdout.startsWith("distinct-text-000000-padding-padding,"), allowed.stdout)
      assertEquals(lifted.exit, 0, lifted.stderr)
      assert(lifted.stdout.contains("distinct-text-000000"), lifted.stdout)
  }
