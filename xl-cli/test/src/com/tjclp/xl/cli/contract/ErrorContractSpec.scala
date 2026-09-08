package com.tjclp.xl.cli.contract

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.cli.BuildInfo

/**
 * The error contract end to end (ADR-017 §2.3), through the in-process harness: every failure
 * leaves stdout EMPTY and puts `Error: <message>` plus a `code:` line on stderr; exit 1 is reserved
 * for findings and gates, 2 for a wrong command line, 3 for an operation that could not complete.
 */
class ErrorContractSpec extends CatsEffectSuite:

  private val nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private val nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  private val fixtures = ResourceSuiteLocalFixture(
    "error-contract-fixtures",
    Resource.make(TestFixtures.materialize.flatTap(extras))(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  /**
   * Beyond the golden fixtures (which include `circular.xlsx`): a non-zip, a styles-less package, a
   * lint incident.
   */
  private def extras(dir: Path): IO[Unit] =
    for
      _ <- IO.blocking(
        Files.write(dir.resolve("corrupt.xlsx"), "not a zip".getBytes(StandardCharsets.UTF_8))
      )
      _ <- IO.blocking(withoutStyles(dir.resolve("simple.xlsx"), dir.resolve("nostyles.xlsx")))
      _ <- IO.blocking(writeZip(dir.resolve("incident.xlsx"), incidentParts))
    yield ()

  /** `source` with `xl/styles.xml` dropped: the reader falls back and reports MissingStylesXml. */
  private def withoutStyles(source: Path, target: Path): Unit =
    val in = new ZipInputStream(Files.newInputStream(source))
    val out = new ZipOutputStream(Files.newOutputStream(target))
    try
      Iterator
        .continually(in.getNextEntry)
        .takeWhile(_ != null)
        .filterNot(_.getName == "xl/styles.xml")
        .foreach { entry =>
          out.putNextEntry(new ZipEntry(entry.getName))
          out.write(in.readAllBytes())
          out.closeEntry()
        }
    finally
      in.close()
      out.close()

  private def writeZip(target: Path, parts: Map[String, String]): Unit =
    val bytes = new ByteArrayOutputStream()
    val zos = new ZipOutputStream(bytes)
    parts.foreach { (name, content) =>
      zos.putNextEntry(new ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()
    }
    zos.close()
    Files.write(target, bytes.toByteArray)

  /**
   * The GH-397 field incident: externalReferences after extLst with a dangling r:id (lint
   * findings).
   */
  private def incidentParts: Map[String, String] = Map(
    "[Content_Types].xml" ->
      """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""",
    "_rels/.rels" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="$nsRel/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
    "xl/workbook.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="$nsMain" xmlns:r="$nsRel">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <extLst/>
  <externalReferences><externalReference r:id="rId5"/></externalReferences>
</workbook>""",
    "xl/_rels/workbook.xml.rels" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="$nsRel/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
    "xl/worksheets/sheet1.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain"><sheetData/></worksheet>"""
  )

  private def file(name: String): String = fixtures().resolve(name).toString

  private def lines(text: String): Vector[String] = text.split("\n", -1).toVector

  /**
   * Every failure: empty stdout, `Error: <message>` first on stderr, then the indented code line.
   */
  private def assertFailure(run: CliRun, exit: Int, message: String, code: String): Unit =
    assertEquals(run.exit, exit, s"exit code\n${run.stdout}\n${run.stderr}")
    assertEquals(run.stdout, "", "stdout must be empty on failure")
    val errLines = lines(run.stderr)
    assertEquals(errLines.headOption, Some(s"Error: $message"), run.stderr)
    assert(errLines.contains(s"  code: $code"), s"missing code line:\n${run.stderr}")

  test("sheet not found: stdout empty, Error + code + hint on stderr, exit 3, nothing written") {
    val out = fixtures().resolve("nope-out.xlsx")
    CliHarness
      .run("-f", file("simple.xlsx"), "-s", "Nope", "-o", out.toString, "put", "A1", "1")
      .map { run =>
        assertFailure(run, 3, "Sheet not found: Nope. Available: Data, Summary", "SHEET_NOT_FOUND")
        assert(run.stderr.contains("  hint: list sheets with `xl -f <file> sheets`"), run.stderr)
        assert(
          !run.stderr.contains("did you mean"),
          s"Nope is not close to any sheet:\n${run.stderr}"
        )
        assert(!Files.exists(out), "a failed write must not create the output")
      }
  }

  test("sheet not found: a near miss gets a did-you-mean line") {
    CliHarness.run("-f", file("simple.xlsx"), "-s", "Dat", "view", "A1:B2").map { run =>
      assertFailure(run, 3, "Sheet not found: Dat. Available: Data, Summary", "SHEET_NOT_FOUND")
      assert(run.stderr.contains("  did you mean: Data"), run.stderr)
    }
  }

  test("sheet required: unqualified ref on a multi-sheet book exits 3 with the sheet names") {
    CliHarness.run("-f", file("simple.xlsx"), "cell", "A1").map { run =>
      assertFailure(
        run,
        3,
        "cell with unqualified ref 'A1' requires --sheet or qualified ref (e.g., Sheet1!A1). Available sheets: Data, Summary",
        "SHEET_REQUIRED"
      )
      assert(run.stderr.contains("  did you mean: Data, Summary"), run.stderr)
      assert(
        run.stderr.contains("  hint: use -s <name> or a qualified ref like 'Name'!A1"),
        run.stderr
      )
    }
  }

  test("missing -o exits 2 with OUTPUT_REQUIRED") {
    CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "put", "A5", "42").map { run =>
      assertFailure(
        run,
        2,
        "put requires -o <out.xlsx> (or -i to modify in place)",
        "OUTPUT_REQUIRED"
      )
    }
  }

  test("-i with -o exits 2 (usage) and touches nothing") {
    val input = fixtures().resolve("inplace.xlsx")
    for
      before <- IO.blocking(Files.readAllBytes(input))
      run <- CliHarness.run(
        "-f",
        input.toString,
        "-s",
        "Data",
        "-i",
        "-o",
        file("conflict-out.xlsx"),
        "put",
        "A1",
        "1"
      )
      after <- IO.blocking(Files.readAllBytes(input))
    yield
      assertFailure(run, 2, "--in-place (-i) and --output (-o) are mutually exclusive", "USAGE")
      assert(java.util.Arrays.equals(before, after), "input must be untouched")
      assert(!Files.exists(fixtures().resolve("conflict-out.xlsx")))
  }

  test("diff: differing files exit 1 with the report on stdout; an unreadable file exits 3") {
    for
      differs <- CliHarness.run("-f", file("simple.xlsx"), "diff", "-g", file("changed.xlsx"))
      unreadable <- CliHarness.run("-f", file("simple.xlsx"), "diff", "-g", file("missing.xlsx"))
    yield
      assertEquals(differs.exit, 1)
      // B1 changed and B4's cache followed it (GH-607)
      assert(differs.stdout.contains("Summary: 2 changed"), differs.stdout)
      assertEquals(differs.stderr, "")
      assertEquals(unreadable.exit, 3, unreadable.stderr)
      assertEquals(unreadable.stdout, "")
      assert(unreadable.stderr.startsWith("Error: "), unreadable.stderr)
      assert(unreadable.stderr.contains("  code: IO_READ"), unreadable.stderr)
  }

  test("lint: findings exit 1 on stdout; usage exits 2; a corrupt zip and a missing file exit 3") {
    for
      findings <- CliHarness.run("lint", file("incident.xlsx"))
      both <- CliHarness.run("-f", "a.xlsx", "lint", "b.xlsx")
      none <- CliHarness.run("lint")
      corrupt <- CliHarness.run("lint", file("corrupt.xlsx"))
      missing <- CliHarness.run("lint", file("missing.xlsx"))
    yield
      assertEquals(findings.exit, 1, findings.stderr)
      assert(findings.stdout.contains("finding(s)"), findings.stdout)
      assertEquals(findings.stderr, "")
      assertFailure(
        both,
        2,
        "lint takes exactly one file — got both -f 'a.xlsx' and positional 'b.xlsx'",
        "USAGE"
      )
      assertFailure(none, 2, "lint requires a file: xl lint <file> (or xl -f <file> lint)", "USAGE")
      assertEquals(corrupt.exit, 3, corrupt.stderr)
      assertEquals(corrupt.stdout, "")
      assert(corrupt.stderr.startsWith("Error: "), corrupt.stderr)
      assert(corrupt.stderr.contains("  code: "), corrupt.stderr)
      assertFailure(missing, 3, s"No such file: ${file("missing.xlsx")}", "IO_READ")
  }

  test("a missing input file is IO_READ (exit 3) on every verb: sheets, names, view, cell, lint") {
    val missing = file("missing.xlsx")
    for
      sheets <- CliHarness.run("-f", missing, "sheets")
      names <- CliHarness.run("-f", missing, "names")
      view <- CliHarness.run("-f", missing, "-s", "Data", "view", "A1:B2")
      cell <- CliHarness.run("-f", missing, "cell", "Data!A1")
      lint <- CliHarness.run("lint", missing)
      diff <- CliHarness.run("-f", missing, "diff", "-g", file("simple.xlsx"))
      streamed <- CliHarness.run("-f", missing, "--stream", "-s", "Data", "view", "A1:B2")
      json <- CliHarness.run("-f", missing, "--json", "-s", "Data", "view", "A1:B2")
    yield
      List(
        "sheets" -> sheets,
        "names" -> names,
        "view" -> view,
        "cell" -> cell,
        "lint" -> lint,
        "diff" -> diff,
        "--stream view" -> streamed
      ).foreach { (verb, run) =>
        assertEquals(run.exit, 3, s"$verb exit\n${run.stderr}")
        assertEquals(run.stdout, "", s"$verb stdout must be empty")
        // GH-621: one prefix, the cause, a hint — never "Failed to read XLSX: IO error: Failed …"
        assert(run.stderr.startsWith(s"Error: No such file: $missing\n"), s"$verb: ${run.stderr}")
        assert(run.stderr.contains("  code: IO_READ"), s"$verb: ${run.stderr}")
        assert(
          run.stderr.contains("  hint: check the path; the previous write may have failed"),
          s"$verb: ${run.stderr}"
        )
        assert(!run.stderr.contains("Failed to read XLSX"), s"$verb: ${run.stderr}")
      }
      val error = ujson.read(json.stdout)("error")
      assertEquals(error("code"), ujson.Str("IO_READ"))
      assertEquals(error("message"), ujson.Str(s"No such file: $missing"))
      assertEquals(error("hint"), ujson.Str("check the path; the previous write may have failed"))
      assertEquals(error("location")("file"), ujson.Str(missing))
  }

  test("GH-621: a corrupt input keeps the reader's own message, never re-prefixed") {
    CliHarness.run("-f", file("corrupt.xlsx"), "-s", "Data", "view", "A1").map { run =>
      assertEquals(run.exit, 3, run.stderr)
      assert(run.stderr.startsWith("Error: "), run.stderr)
      assert(!run.stderr.contains("Failed to read XLSX: "), run.stderr)
      assert(!run.stderr.contains("No such file"), run.stderr)
      assert(run.stderr.contains("  code: IO_READ"), run.stderr)
    }
  }

  test(
    "standalone batch --dry-run: a non-array document is usage (exit 2), a missing source exit 3"
  ) {
    for
      notArray <- CliHarness.run(List("batch", "--dry-run", "-"), """{"op":"put","ref":"A1"}""")
      missing <- CliHarness.run("batch", "--dry-run", file("no-such-ops.json"))
    yield
      assertFailure(notArray, 2, "Batch input must be a JSON array", "BATCH_JSON_INVALID")
      assertEquals(missing.exit, 3, missing.stderr)
      assertEquals(missing.stdout, "")
      assert(missing.stderr.startsWith("Error: "), missing.stderr)
      assert(missing.stderr.contains("  code: "), missing.stderr)
  }

  test("view --eval --strict on an unevaluable sheet is a gate: exit 1, RECALC_GATE, no result") {
    for
      strict <- CliHarness.run(
        "-f",
        file("circular.xlsx"),
        "-s",
        "Data",
        "view",
        "A1:A1",
        "--eval",
        "--strict"
      )
      advisory <- CliHarness.run(
        "-f",
        file("circular.xlsx"),
        "-s",
        "Data",
        "view",
        "A1:A1",
        "--eval"
      )
    yield
      assertEquals(strict.exit, 1, strict.stderr)
      assertEquals(strict.stdout, "", "a gated view renders nothing")
      assert(
        lines(strict.stderr).headOption.exists(_.startsWith("Error: Formula evaluation failed: ")),
        strict.stderr
      )
      assert(strict.stderr.contains("  code: RECALC_GATE"), strict.stderr)
      assert(strict.stderr.contains("  hint: "), strict.stderr)
      // without --strict the same failure is advisory: the view renders and exits 0
      assertEquals(advisory.exit, 0, advisory.stderr)
      assert(advisory.stdout.contains("| A"), advisory.stdout)
  }

  test("strict gate still exits 1 with the summary on stdout; -o still writes, -i does not") {
    val out = fixtures().resolve("strict-out.xlsx")
    val input = fixtures().resolve("circular.xlsx")
    for
      before <- IO.blocking(Files.readAllBytes(input))
      viaOutput <- CliHarness.run("-f", input.toString, "-o", out.toString, "--strict", "recalc")
      written <- IO.blocking(Files.size(out))
      viaInPlace <- CliHarness.run("-f", input.toString, "-i", "--strict", "recalc")
      after <- IO.blocking(Files.readAllBytes(input))
    yield
      assertEquals(viaOutput.exit, 1, viaOutput.stderr)
      assert(viaOutput.stdout.contains(s"Saved: $out"), viaOutput.stdout)
      assertEquals(viaOutput.stderr, "", "a gate is not an error")
      assert(written > 0L, "-o still writes on a strict failure")
      assertEquals(viaInPlace.exit, 1, viaInPlace.stderr)
      assert(viaInPlace.stdout.contains("NOT saved (--strict failure)"), viaInPlace.stdout)
      assertEquals(viaInPlace.stderr, "")
      assert(java.util.Arrays.equals(before, after), "-i leaves the input untouched")
  }

  test(
    "reader warnings surface on stderr as Warning[READER_WARNING] lines; the result stays on stdout"
  ) {
    for
      view <- CliHarness.run("-f", file("nostyles.xlsx"), "-s", "Data", "view", "A1:B1")
      eval <- CliHarness.run("-f", file("nostyles.xlsx"), "-s", "Data", "eval", "=B1*2")
    yield
      assertEquals(view.exit, 0, view.stderr)
      assert(view.stdout.contains("Hello"), view.stdout)
      assertEquals(view.stderr, "Warning[READER_WARNING]: MissingStylesXml\n")
      assertEquals(eval.exit, 0, eval.stderr)
      assert(eval.stdout.contains("Result: 20"), eval.stdout)
      assertEquals(eval.stderr, "Warning[READER_WARNING]: MissingStylesXml\n")
  }

  test("a wrong command line is usage: exit 2, one-line usage plus the error on stderr") {
    for
      unknown <- CliHarness.run("frob")
      flagAfterVerb <- CliHarness.run("-f", file("simple.xlsx"), "view", "A1:B2", "-s", "Data")
      // W2.4: `view` no longer needs a range (it defaults to the used range); `cell` still does
      missingArg <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "cell")
      noArgs <- CliHarness.run()
    yield
      assertFailure(unknown, 2, "unknown verb 'frob'", "UNKNOWN_VERB")
      // ADR-017 §2.2: a global after the verb is hoisted in front of it — a plain success
      assertEquals(flagAfterVerb.exit, 0, flagAfterVerb.stderr)
      assertFailure(missingArg, 2, "Missing expected positional argument!", "USAGE")
      assert(missingArg.stderr.contains("usage: xl "), missingArg.stderr)
      assertEquals(noArgs.exit, 2)
      assertEquals(noArgs.stdout, "")
  }

  test(
    "--stream on an unsupported combination is usage (2) with UNSUPPORTED_IN_STREAM and a hint"
  ) {
    CliHarness
      .run("-f", file("simple.xlsx"), "-s", "Data", "--stream", "view", "A1:B2", "--eval")
      .map { run =>
        assertFailure(
          run,
          2,
          "--eval is not supported with --stream (streaming view uses cached values only)",
          "UNSUPPORTED_IN_STREAM"
        )
        assert(run.stderr.contains("  hint: "), run.stderr)
      }
  }

  test("a value-count mismatch is VALUE_COUNT_MISMATCH (exit 3) with today's message and a hint") {
    CliHarness
      .run(
        "-f",
        file("simple.xlsx"),
        "-s",
        "Data",
        "-o",
        file("mismatch.xlsx"),
        "put",
        "A1:B2",
        "1",
        "2",
        "3"
      )
      .map { run =>
        assertEquals(run.exit, 3, run.stderr)
        assertEquals(run.stdout, "")
        assert(
          run.stderr.startsWith("Error: Range A1:B2 has 4 cells but 3 values provided"),
          run.stderr
        )
        assert(run.stderr.contains("  code: VALUE_COUNT_MISMATCH"), run.stderr)
        assert(run.stderr.contains("  hint: provide exactly 4 values"), run.stderr)
      }
  }

  test("a formula that does not parse is FORMULA_ERROR with the caret context, on putf and eval") {
    for
      putf <- CliHarness
        .run(
          "-f",
          file("simple.xlsx"),
          "-s",
          "Data",
          "-o",
          file("bad-f.xlsx"),
          "putf",
          "A5",
          "=SUM("
        )
      eval <- CliHarness.run("eval", "=SUM(")
    yield
      assertFailure(putf, 3, "=SUM(", "FORMULA_ERROR")
      // the caret block: formula, pointer, the parser's diagnostic — the formula is not repeated
      assert(putf.stderr.contains("=SUM(\n    ^\n"), putf.stderr)
      assert(!putf.stderr.contains("Formula error in '=SUM(':"), putf.stderr)
      assert(putf.stderr.contains("  hint: check the formula with `xl eval`"), putf.stderr)
      assertEquals(eval.exit, 3, eval.stderr)
      assert(eval.stderr.contains("  code: FORMULA_ERROR"), eval.stderr)
  }

  test("a range where one cell is needed is INVALID_REFERENCE: cell, in memory and --stream") {
    val message = "Invalid reference: cell requires a single cell, not the range A1:B2"
    for
      memory <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "cell", "A1:B2")
      stream <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "--stream", "cell", "A1:B2")
    yield
      assertFailure(memory, 3, message, "INVALID_REFERENCE")
      assertFailure(stream, 3, message, "INVALID_REFERENCE")
  }

  test("an unknown sheet is SHEET_NOT_FOUND with candidates on --stream search and on diff") {
    for
      search <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Dat", "--stream", "search", "x")
      filter <- CliHarness
        .run("-f", file("simple.xlsx"), "--stream", "search", "x", "--sheets", "Data,Summry")
      diff <- CliHarness
        .run("-f", file("simple.xlsx"), "-s", "Summry", "diff", "-g", file("changed.xlsx"))
      badName <- CliHarness.run("-f", file("simple.xlsx"), "-s", "a[b", "--stream", "search", "x")
    yield
      assertFailure(search, 3, "Sheet not found: Dat. Available: Data, Summary", "SHEET_NOT_FOUND")
      assert(search.stderr.contains("  did you mean: Data"), search.stderr)
      assertFailure(
        filter,
        3,
        "Sheet not found: Summry. Available: Data, Summary",
        "SHEET_NOT_FOUND"
      )
      assert(filter.stderr.contains("  did you mean: Summary"), filter.stderr)
      assertFailure(diff, 3, "Sheet 'Summry' not found in either workbook", "SHEET_NOT_FOUND")
      assert(diff.stderr.contains("  did you mean: Summary"), diff.stderr)
      assertEquals(badName.exit, 3, badName.stderr)
      assert(badName.stderr.contains("  code: INVALID_SHEET_NAME"), badName.stderr)
  }

  test("an output directory that does not exist is IO_WRITE (exit 3), never a stack trace: put") {
    val target = fixtures().resolve("no").resolve("such").resolve("dir").resolve("out.xlsx")
    val message =
      s"cannot write $target: no such directory: ${target.getParent}"
    for
      text <- CliHarness
        .run("-f", file("simple.xlsx"), "-s", "Data", "-o", target.toString, "put", "A1", "1")
      json <- CliHarness.run(
        "-f",
        file("simple.xlsx"),
        "-s",
        "Data",
        "-o",
        target.toString,
        "--json",
        "put",
        "A1",
        "1"
      )
    yield
      assertFailure(text, 3, message, "IO_WRITE")
      assert(
        text.stderr.contains("  hint: check that the directory exists and is writable"),
        text.stderr
      )
      assertEquals(json.exit, 3, json.stderr)
      val e = ujson.read(json.stdout)
      assertEquals(e("ok"), ujson.False)
      assertEquals(e("exitCode"), ujson.Num(3))
      assertEquals(e("verb"), ujson.Str("put"))
      assertEquals(e("data"), ujson.Null)
      assertEquals(e("error")("code"), ujson.Str("IO_WRITE"))
      assertEquals(e("error")("message"), ujson.Str(message))
      assertEquals(e("error")("location")("file"), ujson.Str(target.toString))
      assertEquals(json.stdout.count(_ == '\n'), ujson.write(e, indent = 2).count(_ == '\n') + 1)
      assertEquals(json.stderr, s"Error: $message\n")
  }

  test("an output directory that does not exist is IO_WRITE (exit 3), never a stack trace: new") {
    val target = fixtures().resolve("no").resolve("such").resolve("dir").resolve("new.xlsx")
    for
      text <- CliHarness.run("new", target.toString)
      json <- CliHarness.run("--json", "new", target.toString)
    yield
      assertEquals(text.exit, 3, text.stderr)
      assertEquals(text.stdout, "")
      assert(text.stderr.startsWith(s"Error: cannot write $target: "), text.stderr)
      assert(text.stderr.contains("  code: IO_WRITE"), text.stderr)
      assert(!text.stderr.contains("\tat "), s"no stack trace:\n${text.stderr}")
      assertEquals(json.exit, 3, json.stderr)
      val e = ujson.read(json.stdout)
      assertEquals(e("error")("code"), ujson.Str("IO_WRITE"))
      assertEquals(e("verb"), ujson.Str("new"))
      assertEquals(e("data"), ujson.Null)
      assertEquals(json.stderr.linesIterator.size, 1, json.stderr)
  }

  test(
    "a `--json` that is the value of -o is a file name: the usage failure is text, not an envelope"
  ) {
    // the verb's positional is missing, so decline fails before anything is read or written
    CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "-o", "--json", "put").map { run =>
      assertFailure(run, 2, "Missing expected positional argument!", "USAGE")
      assert(!Files.exists(fixtures().resolve("--json")), "nothing is written")
    }
  }

  test("recalc with --no-recalc is a usage error (exit 2), not INTERNAL") {
    CliHarness
      .run("-f", file("simple.xlsx"), "-o", file("contradiction.xlsx"), "--no-recalc", "recalc")
      .map { run =>
        assertEquals(run.exit, 2, run.stderr)
        assert(
          run.stderr.startsWith("Error: recalc cannot be combined with --no-recalc"),
          run.stderr
        )
        assert(run.stderr.contains("  code: USAGE"), run.stderr)
      }
  }

  test("headless eval and standalone new report on stderr and exit 3") {
    for
      eval <- CliHarness.run("eval", "=SUM(")
      badSheet <- CliHarness.run("new", file("new-bad.xlsx"), "--sheet-name", "a[b")
    yield
      assertEquals(eval.exit, 3, eval.stderr)
      assertEquals(eval.stdout, "")
      assert(eval.stderr.startsWith("Error: "), eval.stderr)
      assert(eval.stderr.contains("  code: "), eval.stderr)
      assertEquals(badSheet.exit, 3, badSheet.stderr)
      assertEquals(badSheet.stdout, "")
      assert(badSheet.stderr.startsWith("Error: "), badSheet.stderr)
      assert(badSheet.stderr.contains("  code: "), badSheet.stderr)
  }

  test("--help exits 0 on stdout (GH-620) and --version prints the version on stdout") {
    for
      help <- CliHarness.run("--help")
      version <- CliHarness.run("--version")
    yield
      assertEquals(help.exit, 0)
      assertEquals(help.stderr, "")
      assert(help.stdout.startsWith("Usage:"), help.stdout)
      assert(help.stdout.contains("Exit codes:"), "the exit-code table is part of --help")
      assertEquals(version.exit, 0)
      assertEquals(version.stdout, s"${BuildInfo.version}\n")
      assertEquals(version.stderr, "")
  }
