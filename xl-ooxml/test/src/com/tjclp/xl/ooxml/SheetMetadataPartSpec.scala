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

  test(
    "a cell block shared with another record is read, but new anchors get a block of their own"
  ) {
    // cm 1's block holds two records: OTHER (t=1) and XLDAPR (t=2, v=0)
    val source = DynamicArrayFixtures.multiRecordMetadataXml
    assert(source.contains("""<bk><rc t="1" v="0"/><rc t="2" v="0"/></bk>"""), source)
    assertEquals(resolve(source), Map(1 -> plain))
    val plan = SheetMetadataPart.plan(Some(source), Set(plain))
    val xml = plan.output match
      case MetadataOutput.Appended(x) => x
      case other => fail(s"expected Appended, got $other")
    // the existing plain XLDAPR block is pointed at again (no duplicate), through a new cell block
    assertEquals(plan.cmOf, Map(plain -> 2))
    assert(xml.contains("""<futureMetadata name="XLDAPR" count="1">"""), xml)
    assert(
      xml.contains(
        """<cellMetadata count="2"><bk><rc t="1" v="0"/><rc t="2" v="0"/></bk><bk><rc t="2" v="0"/></bk></cellMetadata>"""
      ),
      xml
    )
    assertEquals(resolve(xml), Map(1 -> plain, 2 -> plain))
    assertEquals(
      SheetMetadataPart.plan(Some(xml), Set(plain)),
      CellMetadataPlan(Map(plain -> 2), MetadataOutput.Untouched)
    )
    val both = SheetMetadataPart.plan(Some(source), Set(plain, collapsed))
    assertEquals(resolve(outputXml(source, both)), Map(1 -> plain, 2 -> plain, 3 -> collapsed))
    // a source that already has a block of its own for the variant reuses that one
    val withOwn = source.replace(
      """<cellMetadata count="1"><bk><rc t="1" v="0"/><rc t="2" v="0"/></bk>""",
      """<cellMetadata count="2"><bk><rc t="1" v="0"/><rc t="2" v="0"/></bk><bk><rc t="2" v="0"/></bk>"""
    )
    assert(withOwn != source, "fixture edit missed")
    assertEquals(
      SheetMetadataPart.plan(Some(withOwn), Set(plain)),
      CellMetadataPlan(Map(plain -> 2), MetadataOutput.Untouched)
    )
  }

  test("appending to a prefix-qualified part keeps new elements in SpreadsheetML") {
    val namespace = XmlUtil.nsSpreadsheetML
    val labels = Set(
      "metadata",
      "metadataTypes",
      "metadataType",
      "futureMetadata",
      "cellMetadata",
      "valueMetadata",
      "bk",
      "rc",
      "extLst",
      "ext"
    )
    Vector(DynamicArrayFixtures.metadataXml, richOnly).foreach { source =>
      val qualified = source
        .replace(ns, s"""xmlns:m="$namespace"""")
        .replaceAll(s"<(/?)(${labels.toVector.sorted.mkString("|")})([ />])", "<$1m:$2$3")
      val plan = SheetMetadataPart.plan(Some(qualified), Set(plain, collapsed))
      val appended = outputXml(qualified, plan)
      val root =
        XmlSecurity.parseSafe(appended, "metadata.xml").fold(e => fail(e.message), identity)
      root.descendant_or_self.collect { case e: Elem if labels.contains(e.label) => e }.foreach {
        e =>
          assertEquals(Option(e.namespace), Some(namespace), s"element ${e.label}: $appended")
      }
      assertEquals(resolve(appended), Map(1 -> plain, 2 -> collapsed))
      assertEquals(
        SheetMetadataPart.plan(Some(appended), Set(plain, collapsed)).output,
        MetadataOutput.Untouched
      )
    }
  }

  /**
   * What each `vm` index (1-based `valueMetadata/bk`) resolves to: the name of the metadata type
   * its `rc t` names and the `rvb i` of the `futureMetadata` block its `rc v` names.
   */
  private def vmTargets(xml: String): Vector[(String, String)] =
    val root = XmlSecurity.parseSafe(xml, "metadata.xml").fold(e => fail(e.message), identity)
    def kids(e: Elem, label: String): Vector[Elem] =
      e.child.collect { case c: Elem if c.label == label => c }.toVector
    def at(e: Elem, name: String): String = e.attributes.asAttrMap.getOrElse(name, "")
    val types = kids(root, "metadataTypes").flatMap(kids(_, "metadataType")).map(at(_, "name"))
    val future = kids(root, "futureMetadata").map(f => at(f, "name") -> kids(f, "bk")).toMap
    kids(root, "valueMetadata").flatMap(kids(_, "bk")).flatMap(kids(_, "rc")).map { rc =>
      val name = types.lift(at(rc, "t").toInt - 1).getOrElse(fail(s"no type ${at(rc, "t")}"))
      val block = future.get(name).flatMap(_.lift(at(rc, "v").toInt)).getOrElse(fail("no block"))
      val rvb = (block \\ "rvb").collectFirst { case e: Elem => at(e, "i") }
      name -> rvb.getOrElse(fail("no rvb"))
    }

  test("appending to a prefix-bound rich-value part reparses and keeps every vm index") {
    val namespace = XmlUtil.nsSpreadsheetML
    def rich(i: Int) =
      s"""<x:bk><x:extLst><x:ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="$i"/></x:ext></x:extLst></x:bk>"""
    // SpreadsheetML bound only through the `x` prefix; two images in cells (vm 1 and vm 2)
    val source =
      s"""$head<x:metadata xmlns:x="$namespace" $xlrdNs><x:metadataTypes count="1">${richType
          .replace(
            "<metadataType",
            "<x:metadataType"
          )}</x:metadataTypes><x:futureMetadata name="XLRICHVALUE" count="2">${rich(
          0
        )}${rich(
          1
        )}</x:futureMetadata><x:valueMetadata count="2"><x:bk><x:rc t="1" v="0"/></x:bk><x:bk><x:rc t="1" v="1"/></x:bk></x:valueMetadata></x:metadata>"""
    val before = vmTargets(source)
    assertEquals(before, Vector("XLRICHVALUE" -> "0", "XLRICHVALUE" -> "1"))
    val plan = SheetMetadataPart.plan(Some(source), Set(plain, collapsed))
    val appended = plan.output match
      case MetadataOutput.Appended(xml) => xml
      case other => fail(s"expected Appended, got $other")
    // reparses; the rich values' vm indices resolve exactly as before
    assertEquals(vmTargets(appended), before)
    assertEquals(plan.cmOf, Map(plain -> 1, collapsed -> 2))
    assertEquals(resolve(appended), Map(1 -> plain, 2 -> collapsed))
    // the source's own elements keep their prefix; nothing was renumbered or rebound
    assert(appended.contains("""<x:valueMetadata count="2">"""), appended)
    assert(appended.contains("""<x:rc t="1" v="1"/>"""), appended)
  }

  test("a collapsed-only XLDAPR part gains the plain block under its existing type") {
    val plan = SheetMetadataPart.plan(Some(collapsedOnly), Set(plain))
    val xml = plan.output match
      case MetadataOutput.Appended(x) => x
      case other => fail(s"expected Appended, got $other")
    // the existing type is reused (no second XLDAPR type), its block list grows by one
    assert(xml.contains(s"""<metadataTypes count="1">$xldaprType</metadataTypes>"""), xml)
    assert(xml.contains("""<futureMetadata name="XLDAPR" count="2">"""), xml)
    assert(
      xml.contains(
        """<cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="1" v="1"/></bk></cellMetadata>"""
      ),
      xml
    )
    assertEquals(plan.cmOf, Map(plain -> 2))
    // cm 1 still names the collapsed block; cm 2 the new plain one
    assertEquals(resolve(xml), Map(1 -> collapsed, 2 -> plain))
    assertEquals(
      SheetMetadataPart.plan(Some(xml), Set(plain, collapsed)).output,
      MetadataOutput.Untouched
    )
  }

  test("external sheetMetadata relationships do not identify a package part") {
    val external = Relationship("rId2", SheetMetadataPart.relType, "metadata.xml", Some("External"))
    assertEquals(SheetMetadataPart.locate(Relationships(Seq(external))), None)
    // a target resolving above the package root names no part: never read, never written
    Vector("../../evil.xml", "../../../x/evil.xml", "/../../evil.xml", "..").foreach { target =>
      val escaping = Relationship("rId4", SheetMetadataPart.relType, target, None)
      assertEquals(SheetMetadataPart.locate(Relationships(Seq(escaping))), None, target)
    }
    // one that climbs and comes back stays inside: xl/custom/../metadata.xml is xl/metadata.xml,
    // and ../evil.xml (from xl/) is the package-root part evil.xml
    val inside = Relationship("rId5", SheetMetadataPart.relType, "custom/../metadata.xml", None)
    assertEquals(SheetMetadataPart.locate(Relationships(Seq(inside))), Some("xl/metadata.xml"))
    val root = Relationship("rId6", SheetMetadataPart.relType, "../evil.xml", None)
    assertEquals(SheetMetadataPart.locate(Relationships(Seq(root))), Some("evil.xml"))
    val internal = external.copy(id = "rId3", targetMode = None)
    assertEquals(
      SheetMetadataPart.locate(Relationships(Seq(external, internal))),
      Some("xl/metadata.xml")
    )
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

  /**
   * An arbitrary readable part: up to two non-XLDAPR types and an optional XLDAPR type at any
   * position, up to three XLDAPR blocks with any flags, and up to three cell blocks of one or two
   * records naming any type (so blocks shared between XLDAPR and other records occur too).
   */
  private val genPart: Gen[String] =
    for
      others <- Gen.choose(0, 2)
      withDapr <- if others == 0 then Gen.const(true) else Gen.oneOf(true, false)
      at <- Gen.choose(0, others)
      flags <- Gen
        .choose(1, 3)
        .flatMap(n => Gen.listOfN(n, Gen.zip(Gen.oneOf("0", "1"), Gen.oneOf("0", "1"))))
      names = {
        val other = Vector.tabulate(others)(i => s"XLOTHER$i")
        if withDapr then other.patch(at, Vector("XLDAPR"), 0) else other
      }
      rc = Gen.choose(1, names.size).flatMap { t =>
        val blocks = if names(t - 1) == "XLDAPR" then flags.size else 1
        Gen.choose(0, blocks - 1).map(v => s"""<rc t="$t" v="$v"/>""")
      }
      cells <- Gen
        .choose(0, 3)
        .flatMap(n => Gen.listOfN(n, Gen.choose(1, 2).flatMap(Gen.listOfN(_, rc))))
    yield
      val types =
        names.map(n => if n == "XLDAPR" then xldaprType else richType.replace("XLRICHVALUE", n))
      val future = names.map { n =>
        if n == "XLDAPR" then
          s"""<futureMetadata name="XLDAPR" count="${flags.size}">${flags
              .map((d, c) => dapr(d, c))
              .mkString}</futureMetadata>"""
        else
          s"""<futureMetadata name="$n" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>"""
      }
      val cellPart =
        if cells.isEmpty then ""
        else
          s"""<cellMetadata count="${cells.size}">${cells
              .map(b => s"<bk>${b.mkString}</bk>")
              .mkString}</cellMetadata>"""
      s"""$head<metadata $ns $xlrdNs $xdaNs><metadataTypes count="${names.size}">${types.mkString}</metadataTypes>${future.mkString}$cellPart</metadata>"""

  /**
   * What each `cm` index (1-based `cellMetadata/bk`) means, and what each `futureMetadata` block
   * holds: every record as its type's name plus the block it names, compacted. Independent of the
   * type and block numbering, so a renumbering that kept meanings would still pass, and any change
   * of meaning would not.
   */
  private def meanings(xml: String): (Vector[Vector[String]], Map[String, Vector[String]]) =
    val root = XmlSecurity.parseSafe(xml, "metadata.xml").fold(e => fail(e.message), identity)
    def kids(e: Elem, label: String): Vector[Elem] =
      e.child.collect { case c: Elem if c.label == label => c }.toVector
    def at(e: Elem, name: String): String = e.attributes.asAttrMap.getOrElse(name, "")
    val types = kids(root, "metadataTypes").flatMap(kids(_, "metadataType")).map(at(_, "name"))
    val future =
      kids(root, "futureMetadata")
        .map(f => at(f, "name") -> kids(f, "bk").map(XmlUtil.compact))
        .toMap
    val cells = kids(root, "cellMetadata").flatMap(kids(_, "bk")).map { bk =>
      kids(bk, "rc").map { rc =>
        val name = types.lift(at(rc, "t").toInt - 1).getOrElse(fail(s"no type ${at(rc, "t")}"))
        val block =
          future.get(name).flatMap(_.lift(at(rc, "v").toInt)).getOrElse(fail(s"no block: $xml"))
        s"$name:$block"
      }
    }
    (cells, future)

  property("plan over any part: existing cm/bk meanings never move, every planned cm resolves") {
    val genNeeded = Gen.nonEmptyContainerOf[Set, ArrayMode.Dynamic](Gen.oneOf(plain, collapsed))
    forAll(genPart, genNeeded) { (source, needed) =>
      val p = SheetMetadataPart.plan(Some(source), needed)
      val out = outputXml(source, p)
      val (cellsBefore, futureBefore) = meanings(source)
      val (cellsAfter, futureAfter) = meanings(out)
      // every existing cm index means what it meant; new ones only follow
      assertEquals(cellsAfter.take(cellsBefore.size), cellsBefore, out)
      // every existing futureMetadata block keeps its position (rc v) and content
      futureBefore.foreach { case (name, blocks) =>
        assertEquals(futureAfter.getOrElse(name, Vector.empty).take(blocks.size), blocks, out)
      }
      val before = resolve(source)
      val after = resolve(out)
      before.foreach { case (cm, mode) => assertEquals(after.get(cm), Some(mode), out) }
      // the plan covers exactly what was asked, and the written part resolves each to its variant
      assertEquals(p.cmOf.keySet, needed)
      p.cmOf.foreach { case (mode, cm) => assertEquals(after.get(cm), Some(mode), out) }
    }
  }
