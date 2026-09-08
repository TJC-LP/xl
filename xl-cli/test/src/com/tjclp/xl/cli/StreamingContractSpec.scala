package com.tjclp.xl.cli

import java.io.IOException
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
  Render,
  StreamedBody,
  TestFixtures,
  Warning,
  WarningCode
}
import com.tjclp.xl.cli.read.ReadTestKit
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.{SstPolicy, WriterConfig}

/**
 * The streaming contract of the 0.22.0 fixes, end to end through the harness.
 *
 *   - GH-635: `--stream view --limit 0` streams — the same bytes as the bounded and the in-memory
 *     runs, for csv, json (bare and in the envelope) and markdown; `--limit -1` is a usage error; a
 *     failure before the first row is the ordinary failure (the envelope under `--json`), one after
 *     bytes went out stops the output and reports on stderr; a reader that closes stdout (`| head`)
 *     ends the run quietly, a stdout that fails otherwise is `IO_WRITE`.
 *   - GH-638: every verb answers `--stream` from one table — `sheets --stats` is refused before any
 *     read like `describe --full`; `names` and `lint` run under it; `diff`, `eval`, `schema` and
 *     the other refusals are `UNSUPPORTED_IN_STREAM` and never decline's "Unexpected argument",
 *     while `--help` still helps; the reader's size limit is `SECURITY_ERROR` with the
 *     `--max-size`/`--stream` hint, not `IO_READ`; a worksheet the streaming parser rejects is
 *     `IO_READ` with the file and sheet, as the loaded reader reports the part.
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
   * table alone does too — the smallest `--max-size` the flag accepts. `corrupt-sheet.xlsx`:
   * `simple.xlsx` with the last 30 bytes of its first worksheet part cut off, so the parser rejects
   * `Data` once it reads past its rows.
   */
  private def extras(dir: Path): IO[Unit] =
    val sheet = (0 until 40000).foldLeft(Sheet("Data")) { (s, i) =>
      s.put(ARef.from0(i % 8, i / 8), f"distinct-text-$i%06d-padding-padding")
    }
    val corrupt = dir.resolve("corrupt-sheet.xlsx")
    ExcelIO
      .instance[IO]
      .writeWith(
        Workbook(Vector(sheet)),
        dir.resolve("big.xlsx"),
        WriterConfig(sstPolicy = SstPolicy.Always)
      ) *>
      IO.blocking(Files.copy(dir.resolve("simple.xlsx"), corrupt)).void *>
      ReadTestKit.rewriteZip(corrupt) { (name, bytes) =>
        if name == "xl/worksheets/sheet1.xml" then bytes.dropRight(30) else bytes
      }

  private def envelope(run: CliRun): ujson.Value = ujson.read(run.stdout)

  private def assertRefused(run: CliRun, messageStart: String): Unit =
    assertEquals(run.exit, 2, run.stderr)
    assertEquals(run.stdout, "")
    val first = run.stderr.linesIterator.toVector.headOption.getOrElse("")
    assert(first.startsWith(s"Error: $messageStart"), run.stderr)
    assert(run.stderr.contains("  code: UNSUPPORTED_IN_STREAM"), run.stderr)
    assert(run.stderr.contains("  hint: "), run.stderr)
    assert(run.stderr.contains("omit --stream"), run.stderr)

  /** `Main.emit` over an in-memory `CliIO`: `(exit, stdout, stderr)`. */
  private def emitted(
    outcome: Outcome,
    mode: OutputMode,
    adapt: CliIO => CliIO = identity
  ): IO[(Int, String, String)] =
    CliIO.capturing("").flatMap { (io, collect) =>
      Main.emit(outcome, mode, adapt(io)).flatMap { exit =>
        collect.map((out, err) => (exit.code, out, err))
      }
    }

  private def streamedLines(lines: Vector[String], warnings: Vector[Warning] = Vector.empty) =
    Outcome.ok(
      "view",
      Payload.Streamed(
        StreamedBody.Lines(Vector.empty, Stream.emits(lines).covary[IO], Vector.empty)
      ),
      warnings
    )

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
      // (--json without --format selects the JSON payload, so markdown is asked for by name)
      markdownWrapped <- run(
        base ++ Vector("--json", "--stream", "view", "--limit", "0", "--format", "markdown")
      )
      markdownWrappedMemory <- run(
        base ++ Vector("--json", "view", "--limit", "0", "--format", "markdown")
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
      // data.text streams too, escaped line by line: the same envelope as the gathered text
      assertEquals(csvWrapped.stdout, csvWrappedMemory.stdout)
      assert(envelope(csvWrapped)("data")("text").str.startsWith("Hello,10"), csvWrapped.stdout)
      assertEquals(markdownWrapped.stdout, markdownWrappedMemory.stdout)
      assert(
        envelope(markdownWrapped)("data")("text").str.startsWith("|   | A"),
        markdownWrapped.stdout
      )
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

  test("a 600-row streamed table is written as the whole, past one write of 256 fragments") {
    val rows = (1 to 600).toVector.map(i => s"row $i,$i.5")
    val text = rows.mkString("\n")
    for
      plain <- emitted(streamedLines(rows), OutputMode.Text)
      wrapped <- emitted(streamedLines(rows), OutputMode.Json)
    yield
      assertEquals(plain, (0, text + "\n", ""))
      val expected = Render.json(Outcome.ok("view", Payload.text(text)), BuildInfo.version)
      assertEquals(wrapped, (0, expected.stdout + "\n", ""))
      assertEquals(ujson.read(wrapped._2)("data")("text"), ujson.Str(text))
  }

  test("a failure before the first row is the ordinary failure: the envelope under --json") {
    val failing = Stream.raiseError[IO](new IllegalStateException("the reader gave up"))
    val warnings = Vector(Warning(WarningCode.TRUNCATED, "… showing 2 of 4 rows"))
    val outcome = Outcome.ok(
      "view",
      Payload.Streamed(StreamedBody.Lines(Vector(",A,B"), failing, Vector.empty)),
      warnings
    )
    for
      plain <- emitted(outcome, OutputMode.Text)
      wrapped <- emitted(outcome, OutputMode.Json)
    yield
      val (exit, out, err) = plain
      assertEquals(exit, 3)
      assertEquals(out, "")
      assert(err.startsWith("Error: the reader gave up"), err)
      assert(err.contains("  code: INTERNAL"), err)
      assert(err.contains("Warning[TRUNCATED]"), err)
      val (exitJson, outJson, errJson) = wrapped
      assertEquals(exitJson, 3)
      val e = ujson.read(outJson)
      assertEquals(e("ok"), ujson.False)
      assertEquals(e("error")("code"), ujson.Str("INTERNAL"))
      assertEquals(e("warnings").arr.size, 1)
      assertEquals(errJson.trim, "Error: the reader gave up")
  }

  test("a failure after bytes went out stops the output there and reports on stderr") {
    val rows = Stream.emits((1 to 600).toVector.map(i => s"row $i")).covary[IO] ++
      Stream.raiseError[IO](new IllegalStateException("the reader gave up"))
    val outcome =
      Outcome.ok("view", Payload.Streamed(StreamedBody.Lines(Vector.empty, rows, Vector.empty)))
    for
      plain <- emitted(outcome, OutputMode.Text)
      wrapped <- emitted(outcome, OutputMode.Json)
    yield
      val (exit, out, err) = plain
      assertEquals(exit, 3)
      // two writes of 256 fragments went out; the third, which the failure cut short, never did
      assertEquals(out.linesIterator.size, 512)
      assert(out.startsWith("row 1\nrow 2\n"), out.take(40))
      assert(err.startsWith("Error: the reader gave up"), err)
      assert(err.contains("  code: INTERNAL"), err)
      val (exitJson, outJson, errJson) = wrapped
      assertEquals(exitJson, 3)
      // under --json the envelope is left unterminated: the exit code and stderr carry the failure
      assert(outJson.startsWith("{\n  \"ok\": true"), outJson.take(40))
      assert(scala.util.Try(ujson.read(outJson)).isFailure, "not a complete envelope")
      assertEquals(errJson.trim, "Error: the reader gave up")
  }

  test("a reader that closes stdout ends a streamed run quietly with exit 0 (xl … | head)") {
    val rows = (0 until 2000).toVector.map(i => s"row $i")
    for
      writes <- Ref.of[IO, Int](0)
      result <- emitted(
        streamedLines(rows),
        OutputMode.Text,
        io =>
          io.copy(write =
            text =>
              // the pipe closes after the first write
              writes.getAndUpdate(_ + 1).flatMap {
                case 0 => io.write(text)
                case _ => IO.raiseError(CliIO.StdoutClosed)
              }
          )
      )
      written <- writes.get
    yield
      val (exit, out, err) = result
      assertEquals(exit, 0)
      assertEquals(err, "")
      assertEquals(written, 2)
      assertEquals(out.linesIterator.size, 256)
  }

  test("a stdout that fails for another reason is IO_WRITE (exit 3), never a truncated success") {
    val rows = (0 until 2000).toVector.map(i => s"row $i")
    val noSpace = new CliIO.StdoutFailed(new IOException("No space left on device"))
    for
      atOnce <- emitted(
        streamedLines(rows),
        OutputMode.Text,
        io => io.copy(write = _ => IO.raiseError(noSpace))
      )
      writes <- Ref.of[IO, Int](0)
      later <- emitted(
        streamedLines(rows),
        OutputMode.Json,
        io =>
          io.copy(write =
            text =>
              writes.getAndUpdate(_ + 1).flatMap {
                case 0 => io.write(text)
                case _ => IO.raiseError(noSpace)
              }
          )
      )
    yield
      val (exit, out, err) = atOnce
      assertEquals(exit, 3)
      assertEquals(out, "")
      assert(err.startsWith("Error: cannot write to stdout: No space left on device"), err)
      assert(err.contains("  code: IO_WRITE"), err)
      val (exitLater, _, errLater) = later
      assertEquals(exitLater, 3)
      assertEquals(errLater.trim, "Error: cannot write to stdout: No space left on device")
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
  }

  test(
    "--help still helps under --stream, for a verb that refuses the flag and one that takes it"
  ) {
    for
      diff <- CliHarness.run("-f", file("simple.xlsx"), "--stream", "diff", "--help")
      eval <- CliHarness.run("--stream", "eval", "--help")
      audit <- CliHarness.run("--stream", "audit", "--help")
      view <- CliHarness.run("--stream", "view", "--help")
    yield Vector(diff, eval, audit, view).foreach { run =>
      // GH-620: help is a result, on stdout
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(run.stderr, "")
      assert(run.stdout.startsWith("Usage:"), run.stdout)
      assert(!run.stdout.contains("Unexpected argument"), run.stdout)
      assert(!(run.stdout + run.stderr).contains("not supported with --stream"), run.stdout)
    }
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

  test("a worksheet the streaming parser rejects is IO_READ, located at the file and sheet") {
    for
      streamed <- CliHarness.run(
        "-f",
        file("corrupt-sheet.xlsx"),
        "-s",
        "Data",
        "--stream",
        "view",
        "A1:C4"
      )
      wrapped <- CliHarness.run(
        "-f",
        file("corrupt-sheet.xlsx"),
        "-s",
        "Data",
        "--stream",
        "--json",
        "view",
        "A1:C4",
        "--format",
        "csv"
      )
      searched <- CliHarness.run(
        "-f",
        file("corrupt-sheet.xlsx"),
        "--stream",
        "--json",
        "search",
        "Hello"
      )
      loaded <- CliHarness.run(
        "-f",
        file("corrupt-sheet.xlsx"),
        "-s",
        "Data",
        "--json",
        "view",
        "A1:C4"
      )
    yield
      assertEquals(streamed.exit, 3, streamed.stderr)
      assertEquals(streamed.stdout, "")
      assert(streamed.stderr.startsWith("Error: Parse error at worksheet 'Data':"), streamed.stderr)
      assert(streamed.stderr.contains("(line "), streamed.stderr)
      assert(streamed.stderr.contains("  code: IO_READ"), streamed.stderr)
      // before the first row: the failure envelope, not an unterminated one
      assertEquals(wrapped.exit, 3)
      val e = envelope(wrapped)
      assertEquals(e("ok"), ujson.False)
      assertEquals(e("error")("code"), ujson.Str("IO_READ"))
      assertEquals(e("error")("location")("sheet"), ujson.Str("Data"))
      assertEquals(e("error")("location")("file"), ujson.Str(file("corrupt-sheet.xlsx")))
      assertEquals(envelope(searched)("error")("code"), ujson.Str("IO_READ"))
      // the loaded reader agrees on the code
      assertEquals(envelope(loaded)("error")("code"), ujson.Str("IO_READ"))
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

  test("the table's limit under --json --stream view is a failure envelope for every format") {
    val base = Vector("-f", file("big.xlsx"), "-s", "Data", "--stream", "--json", "--max-size", "1")
    for
      csv <- run(base ++ Vector("view", "A1:B2", "--format", "csv"))
      json <- run(base ++ Vector("view", "A1:B2", "--format", "json"))
      markdown <- run(base ++ Vector("view", "A1:B2"))
      unlimited <- run(base ++ Vector("view", "--limit", "0", "--format", "csv"))
    yield Vector(csv, json, markdown, unlimited).foreach { run =>
      assertEquals(run.exit, 3, run.stderr)
      val e = envelope(run)
      assertEquals(e("ok"), ujson.False)
      assertEquals(e("data"), ujson.Null)
      assertEquals(e("error")("code"), ujson.Str("SECURITY_ERROR"))
      assertEquals(e("verb"), ujson.Str("view"))
      assert(run.stderr.startsWith("Error: Security error"), run.stderr)
    }
  }
