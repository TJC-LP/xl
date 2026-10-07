package com.tjclp.xl.cli

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * GH-694, end to end on the book shape Excel leaves after a target is deleted: a sheet-scoped
 * `_xlnm._FilterDatabase` and a user name both `Support!#REF!`, and cells `=Support!#REF!` (cached
 * `#REF!`) and `=IFERROR(Support!#REF!,0)` (cached 0). The parser refused the spelling, so
 * structural edits on Support exited 3, `recalc` dropped both caches with RECALC_ERRORS, and `putf`
 * refused the text.
 */
class QualifiedErrorLiteralCliSpec extends CatsEffectSuite:

  private val names =
    """<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Support!#REF!</definedName><definedName name="Gone">Support!#REF!</definedName></definedNames>"""

  private val parts: Vector[(String, String)] = Vector(
    "[Content_Types].xml" ->
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""".stripMargin,
    "_rels/.rels" ->
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""".stripMargin,
    "xl/workbook.xml" ->
      s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Support" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets>$names</workbook>""".stripMargin,
    "xl/_rels/workbook.xml.rels" ->
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>""".stripMargin,
    "xl/worksheets/sheet1.xml" ->
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
        |<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="e"><f>Support!#REF!</f><v>#REF!</v></c></row>
        |<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f>IFERROR(Support!#REF!,0)</f><v>0</v></c></row>
        |</sheetData></worksheet>""".stripMargin,
    "xl/worksheets/sheet2.xml" ->
      """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        |<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
        |<row r="1"><c r="A1"><v>7</v></c></row>
        |</sheetData></worksheet>""".stripMargin
  )

  private def source(): IO[Path] = IO.blocking {
    val out = Files.createTempFile("xl-gh694-src-", ".xlsx")
    val zos = new ZipOutputStream(Files.newOutputStream(out))
    try
      parts.foreach { (name, text) =>
        zos.putNextEntry(new ZipEntry(name))
        zos.write(text.getBytes(StandardCharsets.UTF_8))
        zos.closeEntry()
      }
    finally zos.close()
    out
  }

  private def outPath(): IO[Path] = IO.blocking(Files.createTempFile("xl-gh694-out-", ".xlsx"))

  private def workbookXml(path: Path): String =
    val zip = new ZipFile(path.toFile)
    try
      new String(
        zip.getInputStream(zip.getEntry("xl/workbook.xml")).readAllBytes(),
        StandardCharsets.UTF_8
      )
    finally zip.close()

  private def cell(path: Path, sheet: String, at: ARef): IO[CellValue] =
    ExcelIO.instance[IO].read(path).map { wb =>
      wb.sheets.find(_.name.value == sheet).getOrElse(fail(s"no sheet $sheet"))(at).value
    }

  private def assertOk(run: CliRun): Unit =
    assertEquals(run.exit, 0, run.toString)
    assert(!run.toString.contains("RECALC_ERRORS"), run.toString)

  private def assertFormula(value: CellValue, text: String, cached: CellValue): Unit =
    value match
      case CellValue.Formula(f, c, _) =>
        assertEquals(f.stripPrefix("="), text)
        assertEquals(c, Some(cached))
      case other => fail(s"expected =$text cached $cached, got $other")

  List(
    List("insert-rows", "1"),
    List("delete-cols", "C")
  ).foreach { edit =>
    test(s"${edit.head} on a sheet with a Support!#REF! name succeeds, names byte-identical") {
      for
        src <- source()
        out <- outPath()
        run <- CliHarness.run(
          List("-f", src.toString, "-s", "Support", "-o", out.toString) ++ edit,
          ""
        )
      yield
        assertOk(run)
        assert(workbookXml(out).contains(names), workbookXml(out))
    }
  }

  test("recalc keeps Excel's values for =Support!#REF! and =IFERROR(Support!#REF!,0)") {
    for
      src <- source()
      out <- outPath()
      run <- CliHarness.run("-f", src.toString, "-s", "Support", "-o", out.toString, "recalc")
      b1 <- cell(out, "Support", ref"B1")
      b2 <- cell(out, "Support", ref"B2")
    yield
      assertOk(run)
      assertFormula(b1, "Support!#REF!", CellValue.Error(CellError.Ref))
      assertFormula(b2, "IFERROR(Support!#REF!,0)", CellValue.Number(BigDecimal(0)))
  }

  test("putf and batch putf accept the spelling and evaluate it") {
    for
      src <- source()
      out <- outPath()
      single <- CliHarness.run(
        "-f",
        src.toString,
        "-s",
        "Support",
        "-o",
        out.toString,
        "putf",
        "D2",
        "=Support!#REF!+1"
      )
      d2 <- cell(out, "Support", ref"D2")
      twinOut <- outPath()
      batch <- CliHarness.run(
        List("-f", src.toString, "-s", "Support", "-o", twinOut.toString, "batch", "-"),
        """[{"op":"putf","ref":"D2","formula":"=IFERROR(Support!#REF!,5)"}]"""
      )
      twinD2 <- cell(twinOut, "Support", ref"D2")
    yield
      assertOk(single)
      assertFormula(d2, "Support!#REF!+1", CellValue.Error(CellError.Ref))
      assertOk(batch)
      assertFormula(twinD2, "IFERROR(Support!#REF!,5)", CellValue.Number(BigDecimal(5)))
  }

  test("rename-sheet renames the qualifier in cells and names, as Excel does") {
    for
      src <- source()
      out <- outPath()
      run <- CliHarness.run(
        "-f",
        src.toString,
        "-o",
        out.toString,
        "rename-sheet",
        "Support",
        "Help Desk"
      )
      b2 <- cell(out, "Help Desk", ref"B2")
      wb <- ExcelIO.instance[IO].read(out)
    yield
      assertOk(run)
      assertFormula(b2, "IFERROR('Help Desk'!#REF!,0)", CellValue.Number(BigDecimal(0)))
      assertEquals(
        wb.metadata.definedNames.map(dn => (dn.name, dn.formula)).toSet,
        Set("_xlnm._FilterDatabase" -> "'Help Desk'!#REF!", "Gone" -> "'Help Desk'!#REF!")
      )
  }
