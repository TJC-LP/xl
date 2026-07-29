package com.tjclp.xl.cli

import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path, Paths}
import java.util.zip.{ZipEntry, ZipOutputStream}

import cats.effect.{ExitCode, IO}
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.cli.commands.LintCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.lint.WorkbookLint

/**
 * Tests for the lint command (GH-397).
 *
 * The structural checks themselves live in xl-ooxml (WorkbookLintSpec); this covers the CLI
 * surface: the exit-code convention (0 clean / 1 findings / 2 error) and the text/JSON renderers.
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

  test("lint: unreadable file exits 2") {
    for code <- Main.runLint(Paths.get("/nonexistent/no-such-file.xlsx"), LintFormat.Text)
    yield assertEquals(code, ExitCode(2))
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
