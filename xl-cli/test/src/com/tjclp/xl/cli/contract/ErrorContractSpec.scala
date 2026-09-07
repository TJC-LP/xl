package com.tjclp.xl.cli.contract

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.BuildInfo
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

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
   * Beyond the golden fixtures: a non-zip, a styles-less package, a circular book, a lint incident.
   */
  private def extras(dir: Path): IO[Unit] =
    for
      _ <- IO.blocking(
        Files.write(dir.resolve("corrupt.xlsx"), "not a zip".getBytes(StandardCharsets.UTF_8))
      )
      _ <- IO.blocking(withoutStyles(dir.resolve("simple.xlsx"), dir.resolve("nostyles.xlsx")))
      _ <- ExcelIO
        .instance[IO]
        .write(
          Workbook(Vector(Sheet("Data").put(ref"A1", CellValue.Formula("A1+1", None)))),
          dir.resolve("circular.xlsx")
        )
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
      assert(differs.stdout.contains("Summary: 1 changed"), differs.stdout)
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
      assertFailure(
        missing,
        3,
        s"IO error: Failed to open file: ${file("missing.xlsx")}",
        "IO_ERROR"
      )
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

  test("a parse failure is usage: exit 2, help on stderr, stdout empty") {
    for
      unknown <- CliHarness.run("frob")
      flagAfterVerb <- CliHarness.run("-f", file("simple.xlsx"), "view", "A1:B2", "-s", "Data")
      noArgs <- CliHarness.run()
    yield
      assertEquals(unknown.exit, 2)
      assertEquals(unknown.stdout, "")
      assert(unknown.stderr.contains("Unexpected argument: frob"), unknown.stderr)
      assertEquals(flagAfterVerb.exit, 2)
      assertEquals(flagAfterVerb.stdout, "")
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

  test(
    "an un-migrated `new Exception(msg)` still yields a well-formed INTERNAL diagnostic, exit 3"
  ) {
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
        assert(run.stderr.contains("  code: INTERNAL"), run.stderr)
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

  test("--help still exits 0 (on stderr) and --version prints the version on stdout") {
    for
      help <- CliHarness.run("--help")
      version <- CliHarness.run("--version")
    yield
      assertEquals(help.exit, 0)
      assertEquals(help.stdout, "")
      assert(help.stderr.startsWith("Usage:"), help.stderr)
      assert(help.stderr.contains("Exit codes:"), "the exit-code table is part of --help")
      assertEquals(version.exit, 0)
      assertEquals(version.stdout, s"${BuildInfo.version}\n")
      assertEquals(version.stderr, "")
  }
