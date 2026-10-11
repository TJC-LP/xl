package com.tjclp.xl.ooxml

import scala.xml.Elem

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.{ArrayMode, CellValue, FormulaKind}
import com.tjclp.xl.macros.ref

import munit.FunSuite

/** Independent package/namespace probes for PR #721. */
class ReviewMetadataIntegrationSpec extends FunSuite:
  import DynamicArrayFixtures.*

  private val ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  private val plain: ArrayMode.Dynamic = ArrayMode.Dynamic()
  private val collapsed: ArrayMode.Dynamic = ArrayMode.Dynamic(collapsed = true)

  test("metadata append keeps SpreadsheetML namespaces in a prefix-qualified source part") {
    val qualified = metadataXml
      .replace(s"""xmlns="$ns"""", s"""xmlns:m="$ns"""")
      .replaceAll(
        "<(/?)(metadata|metadataTypes|metadataType|futureMetadata|cellMetadata|bk|rc|extLst|ext)([ />])",
        "<$1m:$2$3"
      )
    assertEquals(SheetMetadataPart.parse(qualified).byCm, Map(1 -> plain))
    val plan = SheetMetadataPart.plan(Some(qualified), Set(plain, collapsed))
    val appended = plan.output match
      case SheetMetadataPart.MetadataOutput.Appended(xml) => xml
      case other => fail(s"expected append, got $other")
    val root = XmlSecurity.parseSafe(appended, "metadata.xml").fold(e => fail(e.message), identity)
    val labels = Set(
      "metadata",
      "metadataTypes",
      "metadataType",
      "futureMetadata",
      "cellMetadata",
      "bk",
      "rc",
      "extLst",
      "ext"
    )
    root.descendant_or_self.collect { case e: Elem if labels.contains(e.label) => e }.foreach { e =>
      assertEquals(Option(e.namespace), Some(ns), s"element ${e.label}: $appended")
    }
  }

  test("dynamic metadata can follow another record in the same cellMetadata block") {
    // CT_MetadataBlock contains one or more rc elements, not just its first rc.
    val mixed = metadataXml
      .replace(
        "<metadataTypes count=\"1\">",
        "<metadataTypes count=\"2\"><metadataType name=\"REVIEWINFO\" minSupportedVersion=\"120000\" cellMeta=\"1\"/>"
      )
      .replace(
        "<futureMetadata name=\"XLDAPR\"",
        "<futureMetadata name=\"REVIEWINFO\" count=\"1\"><bk/></futureMetadata><futureMetadata name=\"XLDAPR\""
      )
      .replace("<rc t=\"1\" v=\"0\"/>", "<rc t=\"1\" v=\"0\"/><rc t=\"2\" v=\"0\"/>")
    assertEquals(SheetMetadataPart.parse(mixed).byCm, Map(1 -> plain))
  }

  test("relationship targets identify metadata even when the part name omits metadata") {
    val renamed = parts().map { case (path, xml) =>
      val nextPath = if path == "xl/metadata.xml" then "xl/custom/state.xml" else path
      val nextXml = xml
        .replace("/xl/metadata.xml", "/xl/custom/state.xml")
        .replace("Target=\"metadata.xml\"", "Target=\"custom/state.xml\"")
      nextPath -> nextXml
    }
    val workbook = XlsxReader.readFromBytes(xlsx(renamed)).fold(e => fail(e.message), identity)
    val actual = workbook.sheets.headOption.flatMap(_.cells.get(ref"B1")).map(_.value).collect {
      case CellValue.Formula(_, _, kind) => kind
    }
    assertEquals(actual, Some(FormulaKind.dynamicArray(CellRange(ref"B1", ref"B5"))))
  }
