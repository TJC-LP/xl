package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.cli.contract.CliHarness

/**
 * GH-785: a workbook whose `<sheets>` lists an Excel 4.0 macro sheet (relationship type
 * `xlMacrosheet`) is readable from every CLI entry point; the macro sheet is not listed and its
 * part is copied unchanged by a write.
 */
class NonWorksheetSheetCliSpec extends CatsEffectSuite:

  private val ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private val rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  private val macroPart = "xl/macrosheets/sheet1.xml"
  private val macroXml =
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xm:macrosheet xmlns="$ns" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData></xm:macrosheet>"""

  private def fixture(dir: Path): Path =
    val path = dir.resolve("in.xlsm")
    val parts = Vector(
      "[Content_Types].xml" ->
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/$macroPart" ContentType="application/vnd.ms-excel.macrosheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>""",
      "_rels/.rels" ->
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="$rel/officeDocument" Target="xl/workbook.xml"/></Relationships>""",
      "xl/workbook.xml" ->
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="$ns" xmlns:r="$rel"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Macro1" sheetId="2" r:id="rId2"/></sheets></workbook>""",
      "xl/_rels/workbook.xml.rels" ->
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="$rel/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/><Relationship Id="rId3" Type="$rel/sharedStrings" Target="sharedStrings.xml"/></Relationships>""",
      "xl/worksheets/sheet1.xml" ->
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="$ns"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>5</v></c></row></sheetData></worksheet>""",
      macroPart -> macroXml,
      "xl/sharedStrings.xml" ->
        s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="$ns" count="2" uniqueCount="2"><si><t>Alpha</t></si><si><t>MacroText</t></si></sst>"""
    )
    val zip = new ZipOutputStream(Files.newOutputStream(path))
    try
      parts.foreach { (name, content) =>
        zip.putNextEntry(new ZipEntry(name))
        zip.write(content.getBytes(StandardCharsets.UTF_8))
        zip.closeEntry()
      }
    finally zip.close()
    path

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try new String(zip.getInputStream(zip.getEntry(name)).readAllBytes(), StandardCharsets.UTF_8)
    finally zip.close()

  private val tmp = ResourceFunFixture(
    cats.effect.Resource.make(IO.blocking(Files.createTempDirectory("gh785-cli")))(dir =>
      IO.blocking {
        Files.list(dir).forEach(p => Files.deleteIfExists(p))
        Files.deleteIfExists(dir)
        ()
      }
    )
  )

  tmp.test("GH-785: sheets, view and put read a workbook with a macro sheet") { dir =>
    val in = fixture(dir).toString
    val out = dir.resolve("out.xlsm")
    for
      sheets <- CliHarness.run("-f", in, "sheets")
      streamed <- CliHarness.run("-f", in, "--stream", "sheets")
      full <- CliHarness.run("-f", in, "sheets", "--stats")
      view <- CliHarness.run("-f", in, "view", "A1:B1")
      put <- CliHarness.run("-f", in, "-o", out.toString, "put", "A2", "NewText")
      after <- CliHarness.run("-f", out.toString, "view", "A1:A2")
    yield
      List(sheets, streamed, full, view, put, after).foreach(r => assertEquals(r.exit, 0, r.stderr))
      List(sheets, streamed, full).foreach { r =>
        assert(r.stdout.contains("Data"), r.stdout)
        assert(!r.stdout.contains("Macro1"), r.stdout)
      }
      assert(view.stdout.contains("Alpha"), view.stdout)
      assert(after.stdout.contains("NewText"), after.stdout)
      assertEquals(entryText(out, macroPart), macroXml)
      assert(entryText(out, "xl/workbook.xml").contains("""name="Macro1""""))
  }
