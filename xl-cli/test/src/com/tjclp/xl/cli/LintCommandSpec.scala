package com.tjclp.xl.cli

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path, Paths}
import java.util.zip.{ZipEntry, ZipOutputStream}

import scala.jdk.CollectionConverters.*

import cats.effect.{ExitCode, IO}
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{CellRange, Sheet, Workbook, given}
import com.tjclp.xl.addressing.{Column, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.LintCommands
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.formula.parser.UnparseableFormula
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.lint.{LintCategory, WorkbookLint}
import com.tjclp.xl.ops.FormulaSupport
import com.tjclp.xl.sheets.dataTableSyntax.*

/**
 * Tests for the lint command (GH-397).
 *
 * The structural checks themselves live in xl-ooxml (WorkbookLintSpec); this covers the CLI
 * surface: the exit-code convention (0 clean / 1 findings / 2 usage / 3 error) and the text/JSON
 * renderers.
 */
class LintCommandSpec extends CatsEffectSuite:

  private val nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private val nsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

  /**
   * The field-incident package (GH-397): externalReferences zip-patched AFTER extLst, carrying an
   * r:id that does not exist in the rels — Excel repair-dialogs, xl <= 0.12.6 read it silently.
   */
  private def incidentZip(): Path =
    val parts = Map(
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
    val baos = ByteArrayOutputStream()
    val zos = ZipOutputStream(baos)
    parts.foreach { case (name, content) =>
      zos.putNextEntry(ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()
    }
    zos.close()
    val path = Files.createTempFile("lint-cli-incident", ".xlsx")
    Files.write(path, baos.toByteArray)
    path

  private def cleanXlsx(): IO[Path] =
    for
      path <- IO(Files.createTempFile("lint-cli-clean", ".xlsx"))
      wb = Workbook(Vector(Sheet("Data").put(ref"A1" -> "hello")))
      _ <- ExcelIO.instance[IO].write(wb, path)
    yield path

  /**
   * GH-442: a seeded data table with a formula put into its interior. `put` carries no data-table
   * awareness (the authoring guards only cover `sheet.dataTable` itself), so this is exactly the
   * silent tear the lint exists to report.
   */
  private def tornDataTableXlsx(): IO[Path] =
    for
      path <- IO(Files.createTempFile("lint-cli-dt-torn", ".xlsx"))
      interior <- IO.fromEither(CellRange.parse("D5:F6").left.map(new Exception(_)))
      seeds = Seq.fill(2)(Seq.fill(3)(CellValue.Number(BigDecimal(1))))
      authored <- IO.fromEither(
        Sheet("Data")
          .put(ref"C4", CellValue.Formula("B1*B2"))
          .put(ref"D4" -> 1, ref"E4" -> 2, ref"F4" -> 3)
          .put(ref"C5" -> 10, ref"C6" -> 20)
          .dataTable(interior, ref"B1", ref"B2", seeds)
          .left
          .map(err => new Exception(err.message))
      )
      wb = Workbook(Vector(authored.put(ref"E5", CellValue.Formula("C4*2"))))
      _ <- ExcelIO.instance[IO].write(wb, path)
    yield path

  // ========== Exit codes (end-to-end through Main.runLint) ==========

  test("lint: clean file written by xl exits 0") {
    for
      path <- cleanXlsx()
      code <- Main.runLint(path, LintFormat.Text)
      _ <- IO(Files.deleteIfExists(path))
    yield assertEquals(code, ExitCode.Success)
  }

  test("lint: field-incident file (child order + dangling r:id) exits 1") {
    for
      path <- IO(incidentZip())
      code <- Main.runLint(path, LintFormat.Text)
      _ <- IO(Files.deleteIfExists(path))
    yield assertEquals(code, ExitCode(1))
  }

  test("lint: torn data table exits 1 and renders the data-table-torn slug (GH-442)") {
    for
      path <- tornDataTableXlsx()
      code <- Main.runLint(path, LintFormat.Text)
      findings <- IO(
        WorkbookLint.lint(path).fold(err => fail(s"lint errored: $err"), identity)
      )
      _ <- IO(Files.deleteIfExists(path))
    yield
      assertEquals(code, ExitCode(1))
      val text = LintCommands.renderText(path.toString, findings)
      assert(text.contains("[data-table-torn]"), text)
      val parsed = ujson.read(LintCommands.renderJson(path.toString, findings))
      assertEquals(parsed("clean").bool, false)
      assertEquals(parsed("findings").arr.map(_("category").str).toSet, Set("data-table-torn"))
  }

  /**
   * GH-460 addendum: openpyxl's serialization of `value=""` — a childless `<c t="inlineStr"/>`. A
   * valid package otherwise, so the ONLY finding is the new class.
   */
  private def emptyInlineStrZip(): Path =
    val parts = Map(
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
</workbook>""",
      "xl/_rels/workbook.xml.rels" ->
        s"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="$nsRel/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
      "xl/worksheets/sheet1.xml" ->
        s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain"><sheetData><row r="1"><c r="A1" t="inlineStr"/><c r="B1" t="inlineStr"><is><t>text</t></is></c></row></sheetData></worksheet>"""
    )
    val baos = ByteArrayOutputStream()
    val zos = ZipOutputStream(baos)
    parts.foreach { case (name, content) =>
      zos.putNextEntry(ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()
    }
    zos.close()
    val path = Files.createTempFile("lint-cli-empty-inline", ".xlsx")
    Files.write(path, baos.toByteArray)
    path

  test(
    "lint: openpyxl-style empty inlineStr book exits 1 with empty-inline-str, and reads (GH-460)"
  ) {
    for
      path <- IO(emptyInlineStrZip())
      code <- Main.runLint(path, LintFormat.Text)
      findings <- IO(WorkbookLint.lint(path).fold(err => fail(s"lint errored: $err"), identity))
      read <- ExcelIO.instance[IO].read(path).attempt
      _ <- IO(Files.deleteIfExists(path))
    yield
      assertEquals(code, ExitCode(1))
      val text = LintCommands.renderText(path.toString, findings)
      assert(text.contains("[empty-inline-str]"), text)
      assert(text.contains("""<c r="A1" t="inlineStr"/>"""), text)
      val parsed = ujson.read(LintCommands.renderJson(path.toString, findings))
      assertEquals(parsed("clean").bool, false)
      assertEquals(
        parsed("findings").arr.map(_("category").str).toSet,
        Set("empty-inline-str")
      )
      // the addendum's other half: the in-memory reader now opens the book lint flags
      assert(read.isRight, s"xl read must tolerate a childless inlineStr, got $read")
  }

  // ========== Severity tiers (PR #659 review): hygiene findings do not fail the gate ==========

  /**
   * An Excel-shaped book: the string lives in xl/sharedStrings.xml (`t="s"`), as every
   * Excel/LibreOffice/openpyxl-authored file stores text. Lints clean.
   */
  private def sstZip(): Path =
    val parts = Map(
      "[Content_Types].xml" ->
        """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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
</workbook>""",
      "xl/_rels/workbook.xml.rels" ->
        s"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="$nsRel/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="$nsRel/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>""",
      "xl/sharedStrings.xml" ->
        s"""<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="$nsMain" count="2" uniqueCount="2"><si><t>plain</t></si><si><t>other</t></si></sst>""",
      "xl/worksheets/sheet1.xml" ->
        s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData></worksheet>"""
    )
    val baos = ByteArrayOutputStream()
    val zos = ZipOutputStream(baos)
    parts.foreach { case (name, content) =>
      zos.putNextEntry(ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()
    }
    zos.close()
    val path = Files.createTempFile("lint-cli-sst", ".xlsx")
    Files.write(path, baos.toByteArray)
    path

  /**
   * The skeptics' reproduction: `put A1 42` on an SST book writes a new file whose shared-string
   * table still carries the replaced text (the writer never prunes a preserved table), so the ONE
   * finding on xl's own output is `shared-string-orphan` — a hygiene finding.
   */
  private def sstBookEditedByPut(): IO[Path] =
    for
      source <- IO(sstZip())
      out <- IO(Files.createTempFile("lint-cli-sst-put", ".xlsx"))
      excel = ExcelIO.instance[IO]
      wb <- excel.read(source)
      sheet <- IO.fromEither(wb("Sheet1").left.map(e => new Exception(e.message)))
      _ <- excel.write(wb.put(sheet.put(ref"A1" -> 42)), out)
      _ <- IO(Files.deleteIfExists(source))
    yield out

  test("lint: an SST book edited by put exits 0 — the orphan is a hygiene finding (PR #659)") {
    for
      path <- sstBookEditedByPut()
      code <- Main.runLint(path, LintFormat.Json)
      findings <- IO(WorkbookLint.lint(path).fold(err => fail(s"lint errored: $err"), identity))
      _ <- IO(Files.deleteIfExists(path))
    yield
      assertEquals(code, ExitCode.Success)
      val parsed = ujson.read(LintCommands.renderJson(path.toString, findings))
      assertEquals(parsed("clean").bool, false) // "clean" stays "no findings at all"
      assertEquals(
        parsed("findings").arr.map(f => (f("category").str, f("severity").str)).toList,
        List(("shared-string-orphan", "hygiene"))
      )
      val text = LintCommands.renderText(path.toString, findings)
      assert(text.contains("(0 repair, 1 hygiene)"), text)
      assert(text.contains("[shared-string-orphan] (hygiene)"), text)
      assert(text.contains("pass --strict"), text)
      // under --strict the trailer explaining the exit 0 is gone
      assert(!LintCommands.renderText(path.toString, findings, strict = true).contains("exit 0"))
  }

  test("lint --strict: the same book exits 1 (hygiene promoted to the gate)") {
    for
      path <- sstBookEditedByPut()
      code <- Main.runLint(path, LintFormat.Text, strict = true)
      _ <- IO(Files.deleteIfExists(path))
    yield assertEquals(code, ExitCode(1))
  }

  test("lint: the pristine SST book lints clean, exit 0 in both gates") {
    for
      path <- IO(sstZip())
      lenient <- Main.runLint(path, LintFormat.Text)
      strict <- Main.runLint(path, LintFormat.Text, strict = true)
      _ <- IO(Files.deleteIfExists(path))
    yield
      assertEquals(lenient, ExitCode.Success)
      assertEquals(strict, ExitCode.Success)
  }

  test("lint: a repair finding exits 1 with or without --strict") {
    for
      path <- IO(incidentZip())
      lenient <- Main.runLint(path, LintFormat.Text)
      strict <- Main.runLint(path, LintFormat.Text, strict = true)
      findings <- IO(WorkbookLint.lint(path).fold(err => fail(s"lint errored: $err"), identity))
      _ <- IO(Files.deleteIfExists(path))
    yield
      assertEquals(lenient, ExitCode(1))
      assertEquals(strict, ExitCode(1))
      val severities = ujson
        .read(LintCommands.renderJson(path.toString, findings))("findings")
        .arr
        .map(_("severity").str)
        .toSet
      assertEquals(severities, Set("repair"))
      // a repair-only report has no tier breakdown and no hygiene tag
      val text = LintCommands.renderText(path.toString, findings)
      assert(!text.contains("hygiene"), text)
  }

  test("lint: --strict is the global flag — accepted before and after the verb, end to end") {
    for
      path <- sstBookEditedByPut()
      after <- contract.CliHarness.run("lint", path.toString, "--strict")
      before <- contract.CliHarness.run("--strict", "lint", path.toString)
      lenient <- contract.CliHarness.run("--json", "lint", path.toString)
      _ <- IO(Files.deleteIfExists(path))
    yield
      assertEquals(after.exit, 1, after.toString)
      assertEquals(before.exit, 1, before.toString)
      assert(after.stdout.contains("[shared-string-orphan] (hygiene)"), after.stdout)
      // the lenient envelope: ok, exit 0, the finding in data, and the LINT_HYGIENE warning
      assertEquals(lenient.exit, 0, lenient.toString)
      val envelope = ujson.read(lenient.stdout)
      assertEquals(envelope("ok").bool, true)
      assertEquals(envelope("data")("findings").arr.map(_("severity").str).toList, List("hygiene"))
      assertEquals(envelope("warnings").arr.map(_("code").str).toList, List("LINT_HYGIENE"))
  }

  test("lint: unreadable file exits 3 (a failure, not usage — ADR-017)") {
    for code <- Main.runLint(Paths.get("/nonexistent/no-such-file.xlsx"), LintFormat.Text)
    yield assertEquals(code, ExitCode(3))
  }

  test("lint: json format also drives the findings exit code") {
    for
      path <- IO(incidentZip())
      code <- Main.runLint(path, LintFormat.Json)
      _ <- IO(Files.deleteIfExists(path))
    yield assertEquals(code, ExitCode(1))
  }

  // ========== Renderers ==========

  test("lint: text output names category slug, part, and locator") {
    val path = incidentZip()
    try
      val findings = WorkbookLint
        .lint(path)
        .fold(err => fail(s"lint errored: $err"), identity)
      assert(findings.nonEmpty, "incident fixture must produce findings")
      val text = LintCommands.renderText(path.toString, findings)
      assert(text.contains("finding(s)"), text)
      assert(text.contains("[child-order]"), text)
      assert(text.contains("[unresolved-rel-id]"), text)
      assert(text.contains("xl/workbook.xml"), text)
      assert(text.contains("externalReferences"), text)
    finally Files.deleteIfExists(path)
  }

  test("lint: clean text output is a single clean line") {
    assertEquals(
      LintCommands.renderText("f.xlsx", Vector.empty),
      "f.xlsx: clean (no findings)"
    )
  }

  test("lint: json output is parseable with the stable schema") {
    val path = incidentZip()
    try
      val findings = WorkbookLint
        .lint(path)
        .fold(err => fail(s"lint errored: $err"), identity)
      val parsed = ujson.read(LintCommands.renderJson(path.toString, findings))
      assertEquals(parsed("clean").bool, false)
      assertEquals(parsed("file").str, path.toString)
      val arr = parsed("findings").arr
      assertEquals(arr.size, findings.size)
      val categories = arr.map(_("category").str).toSet
      assert(categories.contains("child-order"), categories.toString)
      assert(categories.contains("unresolved-rel-id"), categories.toString)
      arr.foreach { f =>
        assert(f("part").str.nonEmpty)
        assert(f("locator").str.nonEmpty)
        assert(f("message").str.nonEmpty)
      }
    finally Files.deleteIfExists(path)
  }

  test("lint: json clean output has clean=true and empty findings") {
    val parsed = ujson.read(LintCommands.renderJson("f.xlsx", Vector.empty))
    assertEquals(parsed("clean").bool, true)
    assertEquals(parsed("findings").arr.size, 0)
  }

  // ========== Arg shape: positional file form (GH-422) ==========

  private def xlCommand = com.monovore.decline.Command("xl", "test")(Main.main)

  private def parsedIO(args: Seq[String]): IO[IO[ExitCode]] =
    IO.fromEither(
      xlCommand.parse(args, Map.empty).left.map(help => new Exception(help.toString))
    )

  test("lint: positional file form parses (GH-422)") {
    val result = xlCommand.parse(Seq("lint", "book.xlsx"), Map.empty)
    assert(result.isRight, s"xl lint <file> must parse, got: $result")
  }

  test("lint: positional file form works end-to-end, exit 0 on clean file (GH-422)") {
    for
      path <- cleanXlsx()
      io <- parsedIO(Seq("lint", path.toString))
      code <- io
      _ <- IO(Files.deleteIfExists(path))
    yield assertEquals(code, ExitCode.Success)
  }

  test("lint: -f flag form still works end-to-end (GH-422)") {
    for
      path <- cleanXlsx()
      io <- parsedIO(Seq("-f", path.toString, "lint"))
      code <- io
      _ <- IO(Files.deleteIfExists(path))
    yield assertEquals(code, ExitCode.Success)
  }

  test("lint: giving the file both ways is rejected with exit 2 (GH-422)") {
    for
      io <- parsedIO(Seq("-f", "a.xlsx", "lint", "b.xlsx"))
      code <- io
    yield assertEquals(code, ExitCode(2))
  }

  test("lint: no file at all exits 2 with a hint instead of decline noise (GH-422)") {
    for
      io <- parsedIO(Seq("lint"))
      code <- io
    yield assertEquals(code, ExitCode(2))
  }

  /**
   * GH-486: docs/reference/cli.md's `xl lint` "What it flags" list must enumerate EVERY
   * LintCategory. It listed 5 of 9 for two releases — an agent reading the docs concluded xl could
   * not detect data-table tearing when it could. This test is the anti-drift gate: adding a
   * LintCategory without documenting its slug fails here.
   */
  private def repoRoot: Path =
    val start = Paths.get(sys.props.getOrElse("user.dir", ".")).toAbsolutePath
    Iterator
      .iterate(start)(_.getParent)
      .takeWhile(_ != null)
      .find(p => Files.isRegularFile(p.resolve("build.mill")))
      .getOrElse(fail(s"could not locate the repo root (no build.mill at or above $start)"))

  test("GH-486: cli.md documents every LintCategory slug in 'What it flags'") {
    val cliDoc = repoRoot.resolve("docs/reference/cli.md")
    assert(Files.isRegularFile(cliDoc), s"missing $cliDoc")
    val body = Files.readString(cliDoc, StandardCharsets.UTF_8)
    val section = body.split("\\*\\*What it flags\\*\\*", -1).lift(1).getOrElse {
      fail("cli.md has no '**What it flags**' section in the xl lint docs")
    }
    // The bullet list runs until the next '**' heading line ("**Exit codes**").
    val bullets = section.split("\n\\*\\*", -1).headOption.getOrElse("")
    val undocumented = LintCategory.values.toList
      .map(_.slug)
      .filterNot(slug => bullets.contains(s"`$slug`"))
    assertEquals(
      undocumented,
      List.empty[String],
      s"lint categories missing from cli.md 'What it flags': ${undocumented.mkString(", ")}"
    )
  }

  test("GH-486: the generated verb table's lint one-liner names the same category families") {
    // The command table is generated from Schema.verbs (docs/reference/generated/cli-verbs.md,
    // rendered by DocsGenSpec); the anti-drift gate follows it there.
    val body = Files.readString(
      repoRoot.resolve("docs/reference/generated/cli-verbs.md"),
      StandardCharsets.UTF_8
    )
    val summaryLine = body.linesIterator
      .find(l => l.startsWith("| `lint`"))
      .getOrElse(fail("generated/cli-verbs.md has no `lint` row in the verb table"))
    // The summary is prose, not a slug list, but it must not claim a narrower scope than reality.
    assert(
      summaryLine.contains("data-table") && summaryLine.contains("content-type"),
      s"lint command-table summary is stale: $summaryLine"
    )
  }

  test("resolveLintFile: exactly-one-file resolution and hint messages (GH-422)") {
    val a = Paths.get("a.xlsx")
    val b = Paths.get("b.xlsx")
    assertEquals(Main.resolveLintFile(Some(a), None), Right(a))
    assertEquals(Main.resolveLintFile(None, Some(b)), Right(b))
    Main.resolveLintFile(Some(a), Some(b)) match
      case Left(msg) => assert(msg.contains("exactly one file"), msg)
      case Right(p) => fail(s"Expected rejection when the file is given twice, got $p")
    Main.resolveLintFile(None, None) match
      case Left(msg) =>
        assert(msg.contains("lint requires a file"), msg)
        assert(msg.contains("-f"), msg)
      case Right(p) => fail(s"Expected rejection when no file is given, got $p")
  }

  // ========== GH-663: formula-unparseable through the real parser ==========

  /** A structurally clean one-sheet package whose A3 holds the given `<f>` text. */
  private def formulaZip(fText: String): Path =
    val parts = Map(
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
</workbook>""",
      "xl/_rels/workbook.xml.rels" ->
        s"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="$nsRel/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
      "xl/worksheets/sheet1.xml" ->
        s"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="$nsMain"><sheetData>
  <row r="1"><c r="A1"><v>1</v></c></row>
  <row r="2"><c r="A2"><v>2</v></c></row>
  <row r="3"><c r="A3"><f>${fText.replace("&", "&amp;").replace("<", "&lt;")}</f><v>3</v></c></row>
</sheetData></worksheet>"""
    )
    val baos = ByteArrayOutputStream()
    val zos = ZipOutputStream(baos)
    parts.foreach { case (name, content) =>
      zos.putNextEntry(ZipEntry(name))
      zos.write(content.getBytes(StandardCharsets.UTF_8))
      zos.closeEntry()
    }
    zos.close()
    val path = Files.createTempFile("lint-cli-formula", ".xlsx")
    Files.write(path, baos.toByteArray)
    path

  private def unparseableFindings(path: Path, stream: Boolean = false) =
    val lint =
      if stream then WorkbookLint.lintStream(path, LintCommands.formulaCheck)
      else WorkbookLint.lint(path, LintCommands.formulaCheck)
    lint
      .fold(e => fail(s"lint errored: $e"), identity)
      .filter(_.category == LintCategory.FormulaUnparseable)

  test(
    "GH-663: `<f>SUM(A1:A2</f>` exits 1 with formula-unparseable carrying the parser's message"
  ) {
    for
      path <- IO(formulaZip("SUM(A1:A2"))
      code <- Main.runLint(path, LintFormat.Text)
      codeStream <- Main.runLint(path, LintFormat.Text, stream = true)
      findings <- IO(unparseableFindings(path))
      streamed <- IO(unparseableFindings(path, stream = true))
      text = LintCommands.renderText(path.toString, findings)
      _ <- IO(Files.deleteIfExists(path))
    yield
      assertEquals(code, ExitCode(1))
      assertEquals(codeStream, ExitCode(1))
      assertEquals(findings.size, 1, findings.mkString("\n"))
      assertEquals(streamed, findings)
      assert(text.contains("[formula-unparseable]"), text)
      assert(text.contains("A3"), text)
      assert(text.contains("<f>SUM(A1:A2</f>: Unexpected end of formula at position"), text)
  }

  test("GH-663: an unknown function is #NAME? on recalculation, not a repair — lint stays clean") {
    for
      path <- IO(formulaZip("FOOBAR(1)"))
      code <- Main.runLint(path, LintFormat.Text)
      findings <- IO(unparseableFindings(path))
      _ <- IO(Files.deleteIfExists(path))
    yield
      assertEquals(code, ExitCode.Success)
      assertEquals(findings, Vector.empty)
  }

  test("GH-663: formulas Excel opens that xl's parser refuses are not repairs either") {
    // the arity model is the registry's — a known name with an odd argument count opens intact
    // (at worst #VALUE!) — so it may not fail the ship gate. The grammar these texts once hit —
    // `TRUE()`, Excel's union ',' and intersection ' ' operators, a flat 130-term chain — parses
    // since #669/#680 (LibreOffice evaluates them: SUM((A1,A2)) = 3, AREAS((A1,B1)) = 2, ...);
    // they stay pinned as non-findings, and so does a 130-deep SUM nest, past Excel's own 64 and
    // the parser's 128-level budget — a parser bound, not a certain repair.
    val flatChain = (2 to 131).map(i => s"B$i").mkString("+")
    val flatConcat = (1 to 130).map(i => s"A$i").mkString("&")
    val deepNest = "SUM(" * 130 + "1" + ")" * 130
    // Nor may a complete text: `NOT` is a legal defined name Excel resolves (a name to the parser
    // since #669; before, its prefix operator awaiting an operand) — the oracle judges truncation
    // from the text, not from the diagnostic class (PR #679 review).
    val texts = Vector(
      "NOT",
      "not",
      "NOT ",
      flatChain,
      flatConcat,
      deepNest,
      "SUM()",
      "TRUE()",
      "FALSE()",
      "SUM((A1,A2))",
      "SUM((A1:A2,B1:B2))",
      "INDEX((A1:B2,A1:C2),1,1,2)",
      "AREAS((A1,B1))",
      "(A1:B2 B1:C2)",
      "SUM((A1:B2 B1:C2))"
    )
    for
      paths <- IO(texts.map(formulaZip))
      findings <- IO(paths.flatMap(unparseableFindings(_)))
      streamed <- IO(paths.flatMap(unparseableFindings(_, stream = true)))
      codes <- paths.traverse(Main.runLint(_, LintFormat.Text))
      codesStream <- paths.traverse(Main.runLint(_, LintFormat.Text, stream = true))
      _ <- IO(paths.foreach(Files.deleteIfExists))
    yield
      assertEquals(findings, Vector.empty)
      assertEquals(streamed, findings)
      assertEquals(codes, Vector.fill(texts.size)(ExitCode.Success))
      assertEquals(codesStream, codes)
  }

  test("GH-663: the certain classes are all findings — unterminated string, wrong closer, limits") {
    // the limit arm of the oracle: 9001 chars (Excel's 8192) — and its finding must stay one
    // readable line, not echo the formula. The parser's depth budget is NOT an arm (see the
    // "not repairs either" test: it refuses flat chains Excel opens).
    val tooLong = "1+" * 4500 + "1"
    // (`A1:` is an InvalidCellRef to the parser, not an EOF, so it is audit's, not lint's)
    val texts = Vector("\"abc", "(A1]", "(A1}", "A1+", "SUM((A1)", tooLong, "'Sheet 1")
    for
      paths <- IO(texts.map(formulaZip))
      findings <- IO(paths.map(unparseableFindings(_)))
      streamed <- IO(paths.map(unparseableFindings(_, stream = true)))
      _ <- IO(paths.foreach(Files.deleteIfExists))
    yield
      assertEquals(streamed, findings)
      texts.zip(findings).foreach { (text, found) =>
        assertEquals(found.size, 1, s"'${text.take(40)}' should be exactly one finding: $found")
        assert(
          found.head.message.length < 400,
          s"finding must stay one readable line (${found.head.message.length}): ${found.head}"
        )
      }
      assert(findings(1).head.message.contains("']'"), findings(1))
      assert(findings(2).head.message.contains("'}'"), findings(2))
      assert(
        findings(5).head.message.contains("Formula too long: 9001 characters (max 8192)"),
        findings(5)
      )
      assert(findings(5).head.message.contains(s"<f>${tooLong.take(80)}…</f>"), findings(5))
  }

  test("PR #679 review: certainTruncation reads the text, not the parser's diagnostic") {
    val truncated = Vector(
      "SUM(A1:A2",
      "\"abc",
      "A1+",
      "A1:",
      "Sheet1!",
      "IF(A1,\"x\",",
      "{1,2",
      "Table1[Col",
      "'Rev (Q1",
      "\"a\"\"b",
      "SUM((A1)",
      "A1&",
      "A1=",
      "A1<"
    )
    val complete = Vector(
      "NOT",
      "NOT ",
      "TotalRev",
      "\"a(b\"",
      "'P&L (adj)'!A1",
      "\"a\"\"b\"",
      "SUM(A1:A2)",
      "A1%",
      "A1#",
      "{1,2;3,4}",
      "''",
      "\"\"",
      "SUM((A1,A2))",
      "(A1:B2 B1:C2)"
    )
    truncated.foreach(t =>
      assert(UnparseableFormula.certainTruncation(t), s"should be truncated: $t")
    )
    complete.foreach(t =>
      assert(!UnparseableFormula.certainTruncation(t), s"should be complete: $t")
    )
  }

  test("PR #679 review: the #669 grammar gaps and error literals are never formula-unparseable") {
    // pinned so a future parser change cannot promote one into a repair-tier claim about a file
    // Excel opens (array constants parse since #669; structured and 3-D references and LAMBDA
    // calls remain gaps). Only the parser-backed category is asserted: on this minimal fixture the
    // storage-form rules still speak (xlfn-missing for a bare x#, @ or LAMBDA; external-ref-dangling
    // for [1]), which is their job, not this oracle's.
    val texts = Vector(
      "SUM(Table1[Amount])",
      "(Table1[Amount])",
      "Table1[[#This Row],[Amount]]",
      "SUM({1,2,3})",
      "{1,2}+A1",
      "(A1:A3={1;2;3})",
      "SUM([1]Sheet1!A1)",
      "([1]Sheet1!A1)",
      "SUM(Sheet1:Sheet3!A1)",
      "A1#",
      "(A1#)",
      "@A1",
      "#REF!",
      "#DIV/0!",
      "LAMBDA(x,x+1)(2)"
    )
    for
      paths <- IO(texts.map(formulaZip))
      findings <- IO(paths.flatMap(unparseableFindings(_)))
      streamed <- IO(paths.flatMap(unparseableFindings(_, stream = true)))
      _ <- IO(paths.foreach(Files.deleteIfExists))
    yield
      assertEquals(findings, Vector.empty)
      assertEquals(streamed, findings)
  }

  test("PR #679 review: the prefiltered oracle answers exactly as the parse-backed one") {
    val corpus = Vector(
      "SUM(A1:A2",
      "\"abc",
      "A1+",
      "'Sheet 1",
      "(A1]",
      "(A1}",
      "SUM((A1)",
      "1+" * 4500 + "1",
      "SUM(" * 130 + "1" + ")" * 130,
      "NOT",
      "TotalRev",
      "SUM(A1:A2)",
      "SUM((A1,A2))",
      "(A1:B2 B1:C2)",
      "TRUE()",
      "SUM()",
      "A1:",
      "Sheet1!",
      "#REF!",
      "\"a]b\"",
      "\"a}b\"",
      "A1]",
      "SUM(A1:A2]",
      "{1,2;3,4}",
      "Table1[Col]",
      "IF(A1,\"x\",",
      ""
    )
    corpus.foreach { t =>
      assertEquals(
        LintCommands.formulaCheck(t),
        UnparseableFormula.checkSlow(t),
        s"'${t.take(40)}'"
      )
    }
  }

  test("GH-663: every real-file fixture in the repo lints free of formula-unparseable") {
    val fixtures = repoRoot.resolve("xl-ooxml/test/resources/fixtures")
    val books = Files.list(fixtures)
    val paths =
      try books.toList.asScala.toVector.filter(_.toString.endsWith(".xlsx")).sorted
      finally books.close()
    assert(paths.nonEmpty, s"no fixtures under $fixtures")
    val flagged = paths.flatMap { p =>
      WorkbookLint.lint(p, LintCommands.formulaCheck) match
        case Right(findings) =>
          findings.filter(_.category == LintCategory.FormulaUnparseable).map(f => s"$p: $f")
        case Left(_) => Vector.empty // the deliberately malformed fixture cannot be linted at all
    }
    assertEquals(flagged, Vector.empty[String], flagged.mkString("\n"))
  }

  test("GH-663: a book xl writes from parsed formulas lints clean") {
    for
      path <- IO(Files.createTempFile("lint-cli-formulas", ".xlsx"))
      wb = Workbook(
        Vector(
          Sheet("Data")
            .put(ref"A1" -> 1, ref"A2" -> 2)
            .put(ref"A3", CellValue.Formula("SUM(A1:A2)"))
            .put(ref"B1", CellValue.Formula("IF(A1>1,\"big\",\"small\")"))
            .put(ref"B2", CellValue.Formula("_xlfn.XLOOKUP(1,A1:A2,A1:A2)"))
            .put(ref"B3", CellValue.Formula("'Data'!A1+Data!A2"))
        )
      )
      _ <- ExcelIO.instance[IO].write(wb, path)
      code <- Main.runLint(path, LintFormat.Text)
      findings <- IO(unparseableFindings(path))
      _ <- IO(Files.deleteIfExists(path))
    yield
      assertEquals(code, ExitCode.Success)
      assertEquals(findings, Vector.empty)
  }

  // ========== #460 item 5: xl's own structural edits and the _FilterDatabase name ==========

  test("#460: xl's own row/column edits keep a filtered book's _FilterDatabase name in sync") {
    // the stale name is the class a range edit leaves behind; xl's structural verbs must not
    // produce it (the fixture carries a matching name for its A1:C6 filter)
    val source = repoRoot.resolve("xl-ooxml/test/resources/fixtures/autofilter.xlsx")
    val sheet = SheetName.unsafe("Filtered")
    val support = summon[FormulaSupport]
    // (label, edit, the filter range the edit must move A1:C6 to)
    val edits: Vector[(String, Workbook => XLResult[Workbook], String)] = Vector(
      ("insert rows inside", wb => support.insertRows(wb, sheet, Row.from1(3), 2), "A1:C8"),
      ("delete rows inside", wb => support.deleteRows(wb, sheet, Row.from1(3), 1), "A1:C5"),
      ("insert columns left", wb => support.insertCols(wb, sheet, Column.from0(0), 1), "B1:D6")
    )
    def entry(zip: Path, name: String): String =
      val file = new java.util.zip.ZipFile(zip.toFile)
      try new String(file.getInputStream(file.getEntry(name)).readAllBytes(), "UTF-8")
      finally file.close()
    edits.traverse_ { (label, edit, moved) =>
      for
        out <- IO(Files.createTempFile("lint-cli-filter", ".xlsx"))
        wb <- ExcelIO.instance[IO].read(source)
        edited <- IO.fromEither(edit(wb).left.map(e => new Exception(e.message)))
        _ <- ExcelIO.instance[IO].write(edited, out)
        findings <- IO(WorkbookLint.lint(out).fold(err => fail(s"lint errored: $err"), identity))
        sheetXml <- IO(entry(out, "xl/worksheets/sheet1.xml"))
        workbookXml <- IO(entry(out, "xl/workbook.xml"))
        _ <- IO(Files.deleteIfExists(out))
      yield
        // the edit really moved both sides — a vacuous pass would not
        assert(sheetXml.contains(s"""<autoFilter ref="$moved""""), s"$label: $sheetXml")
        val (c0, r0) = (moved.takeWhile(_.isLetter), moved.drop(1).takeWhile(_.isDigit))
        assert(workbookXml.contains(s"$$$c0$$$r0:"), s"$label: $workbookXml")
        val stale = findings.filter(_.category == LintCategory.AutoFilterNameMismatch)
        assertEquals(stale, Vector.empty, s"$label: $stale")
    }
  }
