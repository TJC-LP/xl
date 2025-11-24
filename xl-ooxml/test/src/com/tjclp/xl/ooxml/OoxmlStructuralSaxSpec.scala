package com.tjclp.xl.ooxml

import munit.FunSuite
import scala.xml.Elem
import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.richtext.RichText

class OoxmlStructuralSaxSpec extends FunSuite:

  private def roundTrip[A <: SaxSerializable & XmlWritable](part: A): Elem =
    val writer = ScalaXmlSaxWriter()
    part.writeSax(writer)
    writer.result.getOrElse(fail("expected root element"))

  test("ContentTypes.writeSax matches toXml structure") {
    val ct = ContentTypes(
      defaults = Map(
        "rels" -> XmlUtil.ctRelationships,
        "xml" -> "application/xml"
      ),
      overrides = Map(
        "/xl/workbook.xml" -> XmlUtil.ctWorkbook,
        "/xl/worksheets/sheet1.xml" -> XmlUtil.ctWorksheet
      )
    )

    val sax = roundTrip(ct)
    val dom = ct.toXml

    val saxDefaults = (sax \ "Default").map(_.attributes.asAttrMap).toSet
    val domDefaults = (dom \ "Default").map(_.attributes.asAttrMap).toSet
    assertEquals(saxDefaults, domDefaults)

    val saxOverrides = (sax \ "Override").map(_.attributes.asAttrMap).toSet
    val domOverrides = (dom \ "Override").map(_.attributes.asAttrMap).toSet
    assertEquals(saxOverrides, domOverrides)
  }

  test("Relationships.writeSax preserves ids and targets") {
    val rels = Relationships(
      Seq(
        Relationship("rId2", "typeB", "targetB", Some("External")),
        Relationship("rId1", "typeA", "targetA")
      )
    )

    val sax = roundTrip(rels)

    val ids = (sax \ "Relationship").map(_ \ "@Id").map(_.text)
    assertEquals(ids, Seq("rId1", "rId2"))

    val targets = (sax \ "Relationship").map(_ \ "@Target").map(_.text)
    assertEquals(targets, Seq("targetA", "targetB"))

    val targetModes = (sax \ "Relationship").map(_ \ "@TargetMode").map(_.text)
    assertEquals(targetModes, Seq("", "External"))
  }

  test("Workbook.writeSax emits sheets with r:id and state") {
    val wb = OoxmlWorkbook(
      sheets = Seq(
        SheetRef(SheetName.unsafe("Sheet1"), 1, "rId1"),
        SheetRef(SheetName.unsafe("Hidden"), 2, "rId2", Some("hidden"))
      )
    )

    val sax = roundTrip(wb)

    val sheets = (sax \\ "sheet").toVector
    if sheets.isEmpty then fail(s"no sheets in SAX output: ${sax.toString}")
    val names = sheets.map(_ \ "@name").map(_.text)
    assertEquals(names, Vector("Sheet1", "Hidden"))

    val rels = sheets.map(_.attributes.asAttrMap.getOrElse("r:id", ""))
    assertEquals(rels, Vector("rId1", "rId2"))

    val states = sheets.map(_ \ "@state").map(_.text)
    assertEquals(states, Vector("", "hidden"))
  }

  test("OoxmlTable.writeSax emits table columns and refs") {
    val range = CellRange.parse("A1:B2").toOption.getOrElse(fail("expected valid range"))
    val table = OoxmlTable(
      id = 1,
      name = "Table1",
      displayName = "Table1",
      ref = range,
      headerRowCount = 1,
      totalsRowCount = 0,
      totalsRowShown = false,
      columns = Vector(OoxmlTableColumn(1, "Col1"), OoxmlTableColumn(2, "Col2")),
      autoFilter = None,
      styleInfo = None
    )

    val sax = roundTrip(table)
    val dom = table.toXml

    assertEquals((sax \ "table" \ "@ref").map(_.text), (dom \ "table" \ "@ref").map(_.text))
    assertEquals((sax \ "tableColumn" \ "@name").map(_.text), (dom \ "tableColumn" \ "@name").map(_.text))
  }

  test("OoxmlComments.writeSax preserves refs and text") {
    val ref = ARef.parse("A1").toOption.getOrElse(fail("expected ref"))
    val comment = OoxmlComment(ref, authorId = 0, text = RichText.plain("hello"))
    val comments = OoxmlComments(Vector("Author"), Vector(comment))

    val sax = roundTrip(comments)
    val dom = comments.toXml

    assertEquals((sax \ "comment" \ "@ref").map(_.text), (dom \ "comment" \ "@ref").map(_.text))
    assertEquals((sax \ "t").map(_.text), (dom \ "t").map(_.text))
  }
