package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import munit.FunSuite

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.lint.{Finding, LintCategory, WorkbookLint}
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * GH-555: the source `xl/calcChain.xml` leaves the package (part, Override, Relationship) on any
 * write that rewrites a worksheet, and stays on a clean copy. Fixture: datatable-excel.xlsx, an
 * Excel-authored book whose chain names A21, F9 and E1 on sheetId 1.
 */
class CalcChainDropSpec extends FunSuite:

  private val calcChain = "xl/calcChain.xml"
  private val contentTypes = "[Content_Types].xml"
  private val workbookRels = "xl/_rels/workbook.xml.rels"
  private val sheet1 = "xl/worksheets/sheet1.xml"

  private def readFixture(name: String): (java.nio.file.Path, Workbook) =
    val in = TestFixtures.copyToTemp(name)
    (in, XlsxReader.read(in).fold(err => fail(err.message), identity))

  private def writeToBytes(wb: Workbook, config: WriterConfig = WriterConfig.default): Array[Byte] =
    val out = Files.createTempFile("xl-calcchain-", ".xlsx")
    try
      XlsxWriter.writeWith(wb, out, config).fold(err => fail(err.message), identity)
      Files.readAllBytes(out)
    finally Files.deleteIfExists(out)

  private def assertNoChain(output: Map[String, Array[Byte]]): Unit =
    assert(!output.contains(calcChain), output.keys.toVector.sorted.mkString(", "))
    assert(!text(output, contentTypes).contains("calcChain"), text(output, contentTypes))
    assert(!text(output, workbookRels).contains("calcChain"), text(output, workbookRels))

  private def entries(bytes: Array[Byte]): Map[String, Array[Byte]] =
    val zin = new ZipInputStream(new ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(Option(zin.getNextEntry))
        .takeWhile(_.isDefined)
        .flatten
        .map(e => e.getName -> zin.readAllBytes())
        .toMap
    finally zin.close()

  private def text(parts: Map[String, Array[Byte]], name: String): String =
    new String(parts.getOrElse(name, fail(s"missing $name")), StandardCharsets.UTF_8)

  private def rezip(parts: Map[String, Array[Byte]]): Array[Byte] =
    val baos = new ByteArrayOutputStream()
    val zos = new ZipOutputStream(baos)
    parts.toSeq.sortBy(_._1).foreach { case (name, bytes) =>
      zos.putNextEntry(new ZipEntry(name))
      zos.write(bytes)
      zos.closeEntry()
    }
    zos.close()
    baos.toByteArray

  private def lintOf(bytes: Array[Byte]): Vector[Finding] =
    WorkbookLint.lintBytes(bytes).fold(err => fail(err.message), identity)

  test("the fixture carries a chain naming A21, F9 and E1, and lints clean") {
    val (in, _) = readFixture("datatable-excel.xlsx")
    val source = entries(Files.readAllBytes(in))
    val chain = text(source, calcChain)
    List("A21", "F9", "E1").foreach(r => assert(chain.contains(s"""<c r="$r""""), chain))
    assert(text(source, contentTypes).contains("calcChain"))
    assert(text(source, workbookRels).contains("calcChain"))
    assertEquals(lintOf(Files.readAllBytes(in)), Vector.empty[Finding])
  }

  test(
    "datatable-excel-edge.xlsx (two sheets, chain on sheetId 1) lints clean and keeps its chain on a clean write"
  ) {
    val (in, wb) = readFixture("datatable-excel-edge.xlsx")
    val source = entries(Files.readAllBytes(in))
    assert(text(source, calcChain).contains("""i="1""""), text(source, calcChain))
    assertEquals(lintOf(Files.readAllBytes(in)), Vector.empty[Finding])
    assertEquals(text(entries(writeToBytes(wb)), calcChain), text(source, calcChain))
  }

  test("removing a sheet drops the chain, its Override and its Relationship") {
    val (_, wb) = readFixture("datatable-excel-edge.xlsx")
    val removed = wb.remove(SheetName.unsafe("Sheet2")).fold(err => fail(err.message), identity)
    val bytes = writeToBytes(removed)
    assertNoChain(entries(bytes))
    assertEquals(lintOf(bytes), Vector.empty[Finding])
  }

  test("WriterConfig.secure rewrites every sheet, so an otherwise clean write drops the chain") {
    val (_, wb) = readFixture("datatable-excel.xlsx")
    val bytes = writeToBytes(wb, WriterConfig.secure)
    assertNoChain(entries(bytes))
    assertEquals(lintOf(bytes), Vector.empty[Finding])
  }

  test("a clean read -> write keeps the source chain byte-identically") {
    val (in, wb) = readFixture("datatable-excel.xlsx")
    val source = entries(Files.readAllBytes(in))
    val output = entries(writeToBytes(wb))
    assertEquals(text(output, calcChain), text(source, calcChain))
    assert(text(output, contentTypes).contains("calcChain"))
    assert(text(output, workbookRels).contains("calcChain"))
  }

  test("replacing a formula with a value drops the chain, its Override and its Relationship") {
    val (_, wb) = readFixture("datatable-excel.xlsx")
    val sheet = wb.sheets.headOption.getOrElse(fail("missing Sheet1"))
    val edited = wb.put(sheet.put(ref"F9", CellValue.Number(BigDecimal(5))))
    val bytes = writeToBytes(edited)
    val output = entries(bytes)
    assert(!output.contains(calcChain), output.keys.toVector.sorted.mkString(", "))
    assert(!text(output, contentTypes).contains("calcChain"), text(output, contentTypes))
    assert(!text(output, workbookRels).contains("calcChain"), text(output, workbookRels))
    assert(!text(output, sheet1).contains("""<c r="F9"><f>"""), text(output, sheet1))
    assertEquals(lintOf(bytes), Vector.empty[Finding])
  }

  test("a value edit on a cell outside the chain drops the chain too") {
    val (_, wb) = readFixture("datatable-excel.xlsx")
    val sheet = wb.sheets.headOption.getOrElse(fail("missing Sheet1"))
    val output = entries(writeToBytes(wb.put(sheet.put(ref"A30", CellValue.Text("dirty")))))
    assert(!output.contains(calcChain))
    assert(!text(output, contentTypes).contains("calcChain"))
    assert(!text(output, workbookRels).contains("calcChain"))
  }

  test("adding a sheet drops the chain") {
    val (_, wb) = readFixture("datatable-excel.xlsx")
    val added = wb.put(Sheet(SheetName.unsafe("New")).put(ref"A1", CellValue.Number(BigDecimal(1))))
    val output = entries(writeToBytes(added))
    assert(!output.contains(calcChain))
    assert(!text(output, contentTypes).contains("calcChain"))
    assert(!text(output, workbookRels).contains("calcChain"))
  }

  test("the stale chain a dirty write used to carry is what calc-chain-stale flags") {
    val (in, _) = readFixture("datatable-excel.xlsx")
    val source = entries(Files.readAllBytes(in))
    val cleared = text(source, sheet1)
      .replaceAll("""<c r="F9"><f>[^<]*</f><v>[^<]*</v></c>""", """<c r="F9"><v>5</v></c>""")
    assert(cleared != text(source, sheet1), "fixture F9 formula cell shape changed")
    val findings = lintOf(rezip(source + (sheet1 -> cleared.getBytes(StandardCharsets.UTF_8))))
    assertEquals(
      findings.map(f => (f.part, f.category)),
      Vector((calcChain, LintCategory.CalcChainStale))
    )
    val finding = findings.headOption.getOrElse(fail("expected one finding"))
    assertEquals(finding.locator, """<c r="F9" i="1">""")
    assert(finding.message.contains("1 of 3 entries"), finding.message)
  }
