package com.tjclp.xl.ooxml

import scala.xml.Elem

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAll

import com.tjclp.xl.cells.ArrayMode

/**
 * GH-714: the cell-metadata part — `cm` resolution through XLDAPR, and the append-only plan that
 * never moves an index another cell (or a rich value) points at.
 */
class SheetMetadataPartSpec extends ScalaCheckSuite:
  import SheetMetadataPart.{CellMetadataPlan, MetadataOutput}

  private val plain: ArrayMode.Dynamic = ArrayMode.Dynamic()
  private val collapsed: ArrayMode.Dynamic = ArrayMode.Dynamic(collapsed = true)

  private val head =
    """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      |""".stripMargin
  private val ns = """xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""""
  private val xdaNs =
    """xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray""""
  private val xldaprType =
    """<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>"""
  private val richType =
    """<metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>"""
  private def dapr(fDynamic: String, fCollapsed: String) =
    s"""<bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="$fDynamic" fCollapsed="$fCollapsed"/></ext></extLst></bk>"""
  private val richFuture =
    """<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>"""
  private val valueMetadata =
    """<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>"""
  private val xlrdNs =
    """xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata""""

  /** Images in cells only: a rich-value part with no XLDAPR at all. */
  private val richOnly =
    s"""$head<metadata $ns $xlrdNs><metadataTypes count="1">$richType</metadataTypes>$richFuture$valueMetadata</metadata>"""

  /** Rich values first, then XLDAPR as type 2 with a plain and a non-dynamic block. */
  private val mixed =
    s"""$head<metadata $ns $xlrdNs $xdaNs><metadataTypes count="2">$richType$xldaprType</metadataTypes>$richFuture<futureMetadata name="XLDAPR" count="2">${dapr(
        "0",
        "0"
      )}${dapr(
        "1",
        "0"
      )}</futureMetadata><cellMetadata count="2"><bk><rc t="2" v="0"/></bk><bk><rc t="2" v="1"/></bk></cellMetadata>$valueMetadata</metadata>"""

  /** Excel's part with a collapsed block only. */
  private val collapsedOnly =
    s"""$head<metadata $ns $xdaNs><metadataTypes count="1">$xldaprType</metadataTypes><futureMetadata name="XLDAPR" count="1">${dapr(
        "1",
        "1"
      )}</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>"""

  private val sources: Vector[String] =
    Vector(DynamicArrayFixtures.metadataXml, richOnly, mixed, collapsedOnly)

  private def resolve(xml: String): Map[Int, ArrayMode.Dynamic] = SheetMetadataPart.parse(xml).byCm

  private def outputXml(source: String, plan: CellMetadataPlan): String = plan.output match
    case MetadataOutput.Appended(xml) => xml
    case MetadataOutput.Generated(xml) => xml
    case MetadataOutput.Untouched => source

  private def subtree(xml: String, label: String): Option[String] =
    XmlSecurity.parseSafe(xml, "metadata").toOption.flatMap { root =>
      root.child.collectFirst { case e: Elem if e.label == label => XmlUtil.compact(e) }
    }

  test("the generated part is Excel's, byte for byte") {
    assertEquals(SheetMetadataPart.canonicalXml, DynamicArrayFixtures.metadataXml)
    assertEquals(
      SheetMetadataPart.plan(None, Set(plain)),
      CellMetadataPlan(Map(plain -> 1), MetadataOutput.Generated(DynamicArrayFixtures.metadataXml))
    )
  }

  test("parse resolves cm through the XLDAPR type and its fDynamic block") {
    assertEquals(resolve(DynamicArrayFixtures.metadataXml), Map(1 -> plain))
    assertEquals(resolve(collapsedOnly), Map(1 -> collapsed))
    // cm 1 names an fDynamic=0 block: not a dynamic array
    assertEquals(resolve(mixed), Map(2 -> plain))
    assertEquals(resolve(richOnly), Map.empty)
  }

  test("parse is lenient: junk, a dangling v, a non-XLDAPR t") {
    assertEquals(resolve("<metadata"), Map.empty)
    assertEquals(resolve(""), Map.empty)
    val danglingV =
      s"""$head<metadata $ns $xdaNs><metadataTypes count="1">$xldaprType</metadataTypes><futureMetadata name="XLDAPR" count="1">${dapr(
          "1",
          "0"
        )}</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="7"/></bk></cellMetadata></metadata>"""
    assertEquals(resolve(danglingV), Map.empty)
    val wrongType =
      s"""$head<metadata $ns $xdaNs><metadataTypes count="2">$xldaprType$richType</metadataTypes><futureMetadata name="XLDAPR" count="1">${dapr(
          "1",
          "0"
        )}</futureMetadata><cellMetadata count="1"><bk><rc t="2" v="0"/></bk></cellMetadata></metadata>"""
    assertEquals(resolve(wrongType), Map.empty)
  }

  test("a part that already has the variant is reused untouched") {
    assertEquals(
      SheetMetadataPart.plan(Some(DynamicArrayFixtures.metadataXml), Set(plain)),
      CellMetadataPlan(Map(plain -> 1), MetadataOutput.Untouched)
    )
    assertEquals(
      SheetMetadataPart.plan(Some(mixed), Set(plain)),
      CellMetadataPlan(Map(plain -> 2), MetadataOutput.Untouched)
    )
    assertEquals(SheetMetadataPart.plan(Some(richOnly), Set.empty), CellMetadataPlan.none)
    assertEquals(SheetMetadataPart.plan(None, Set.empty), CellMetadataPlan.none)
  }

  test("an images-in-cells part gains XLDAPR by appending") {
    val p = SheetMetadataPart.plan(Some(richOnly), Set(plain))
    val xml = p.output match
      case MetadataOutput.Appended(x) => x
      case other => fail(s"expected Appended, got $other")
    assertEquals(p.cmOf, Map(plain -> 1))
    assertEquals(resolve(xml), Map(1 -> plain))
    assert(xml.contains(s"""<metadataTypes count="2">$richType$xldaprType</metadataTypes>"""), xml)
    assert(
      xml.contains("""<cellMetadata count="1"><bk><rc t="2" v="0"/></bk></cellMetadata>"""),
      xml
    )
    // CT_Metadata order: futureMetadata*, cellMetadata, valueMetadata
    assert(
      xml.indexOf("XLRICHVALUE\" count") < xml.indexOf("futureMetadata name=\"XLDAPR\"") &&
        xml.indexOf("futureMetadata name=\"XLDAPR\"") < xml.indexOf("<cellMetadata") &&
        xml.indexOf("<cellMetadata") < xml.indexOf("<valueMetadata"),
      xml
    )
    assertEquals(subtree(xml, "valueMetadata"), subtree(richOnly, "valueMetadata"))
  }

  test("a generated part for both variants allocates them in a fixed order") {
    val p = SheetMetadataPart.plan(None, Set(collapsed, plain))
    assertEquals(p.cmOf, Map(plain -> 1, collapsed -> 2))
    assertEquals(resolve(outputXml("", p)), Map(1 -> plain, 2 -> collapsed))
  }

  property("append keeps every index and rich value, and a second plan is untouched") {
    val genNeeded = Gen.someOf(Vector(plain, collapsed)).map(_.toSet)
    forAll(Gen.oneOf(sources), genNeeded) { (source, needed) =>
      val p = SheetMetadataPart.plan(Some(source), needed)
      val out = outputXml(source, p)
      val before = resolve(source)
      val after = resolve(out)
      // existing indices never move
      before.foreach { case (cm, mode) => assertEquals(after.get(cm), Some(mode)) }
      // every needed variant is resolvable at the index the plan allocated
      needed.foreach(m => assertEquals(p.cmOf.get(m).flatMap(after.get), Some(m)))
      assertEquals(subtree(out, "valueMetadata"), subtree(source, "valueMetadata"))
      // idempotent: the planned part already has everything
      val again = SheetMetadataPart.plan(Some(out), needed)
      assertEquals(again.output, MetadataOutput.Untouched)
      assertEquals(again.cmOf, p.cmOf)
    }
  }
