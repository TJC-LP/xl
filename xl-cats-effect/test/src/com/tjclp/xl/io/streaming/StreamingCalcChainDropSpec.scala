package com.tjclp.xl.io.streaming

import java.io.InputStream
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path, StandardCopyOption}
import java.util.zip.ZipFile
import scala.jdk.CollectionConverters.*

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.ooxml.lint.{Finding, WorkbookLint}
import com.tjclp.xl.styles.CellStyle

/**
 * GH-555 on the streaming writer: a `--stream` transform drops the source `xl/calcChain.xml` with
 * its Override and Relationship, exactly as the in-memory writer does. Fixture:
 * datatable-excel.xlsx (Excel-authored, chain naming A21, F9 and E1).
 */
class StreamingCalcChainDropSpec extends CatsEffectSuite:

  private val calcChain = "xl/calcChain.xml"

  private def fixturePath(name: String): Path =
    val resource = s"/fixtures/$name"
    val stream: InputStream = Option(getClass.getResourceAsStream(resource)) match
      case Some(s) => s
      case None => fail(s"Fixture not on classpath: $resource")
    try
      val target =
        Files.createTempFile(s"xl-stream-calcchain-${name.stripSuffix(".xlsx")}-", ".xlsx")
      target.toFile.deleteOnExit()
      Files.copy(stream, target, StandardCopyOption.REPLACE_EXISTING)
      target
    finally stream.close()

  private def output(label: String): Path =
    val p = Files.createTempFile(s"xl-stream-calcchain-out-$label-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def entryNames(path: Path): Set[String] =
    val zip = new ZipFile(path.toFile)
    try zip.entries().asIterator().asScala.map(_.getName).toSet
    finally zip.close()

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(name)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $name not found")
    finally zip.close()

  private def assertChainDropped(out: Path): Unit =
    assert(!entryNames(out).contains(calcChain), entryNames(out).toVector.sorted.mkString(", "))
    assert(!entryText(out, "[Content_Types].xml").contains("calcChain"))
    assert(!entryText(out, "xl/_rels/workbook.xml.rels").contains("calcChain"))
    assertEquals(
      WorkbookLint.lint(out).fold(err => fail(err.message), identity),
      Vector.empty[Finding]
    )

  private def aref(a1: String): ARef = ARef.parse(a1).fold(fail(_), identity)

  test("the fixture carries the chain the streaming writer used to copy") {
    val src = fixturePath("datatable-excel.xlsx")
    assert(entryNames(src).contains(calcChain))
    assert(entryText(src, "[Content_Types].xml").contains("calcChain"))
    assert(entryText(src, "xl/_rels/workbook.xml.rels").contains("calcChain"))
  }

  test("transformValues (streaming put) drops the chain, its Override and its Relationship") {
    val src = fixturePath("datatable-excel.xlsx")
    val out = output("values")
    ZipTransformer
      .transformValues[IO](
        src,
        out,
        "xl/worksheets/sheet1.xml",
        Map(aref("F9") -> CellValue.Number(BigDecimal(5)))
      )
      .map { _ =>
        assertChainDropped(out)
        assert(!entryText(out, "xl/worksheets/sheet1.xml").contains("""<c r="F9"><f>"""))
      }
  }

  test("transformStyle (streaming style) drops the chain too") {
    val src = fixturePath("datatable-excel.xlsx")
    val out = output("style")
    ZipTransformer
      .transformStyle[IO](
        src,
        out,
        "xl/worksheets/sheet1.xml",
        com.tjclp.xl.addressing.CellRange.parse("A1:A2").fold(fail(_), identity),
        CellStyle.default,
        replace = true
      )
      .map(_ => assertChainDropped(out))
  }

  test("a package without a chain copies its structural parts byte-for-byte") {
    val src = fixturePath("small-values.xlsx")
    assert(!entryNames(src).contains(calcChain))
    val out = output("nochain")
    ZipTransformer
      .transformValues[IO](
        src,
        out,
        "xl/worksheets/sheet1.xml",
        Map(aref("A1") -> CellValue.Number(BigDecimal(1)))
      )
      .map { _ =>
        assertEquals(entryText(out, "[Content_Types].xml"), entryText(src, "[Content_Types].xml"))
        assertEquals(
          entryText(out, "xl/_rels/workbook.xml.rels"),
          entryText(src, "xl/_rels/workbook.xml.rels")
        )
      }
  }
