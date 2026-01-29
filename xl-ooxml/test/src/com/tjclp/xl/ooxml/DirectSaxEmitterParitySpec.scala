package com.tjclp.xl.ooxml

import munit.FunSuite
import java.io.ByteArrayOutputStream
import java.nio.charset.StandardCharsets
import java.time.LocalDateTime
import scala.xml.{Elem, MetaData, Node, Null, PrefixedAttribute, Text, TopScope, UnprefixedAttribute, XML}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.richtext.RichText.*
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties, Sheet}
// DirectSaxEmitter is in the same package

class DirectSaxEmitterParitySpec extends FunSuite:

  test("direct SAX emission matches worksheet XML structure") {
    val sheet = buildSheet()
    val tableParts = Some(buildTableParts())

    val expected = OoxmlWorksheet.fromDomainWithSST(sheet, None, Map.empty, tableParts)
    val actual = emitDirect(sheet, tableParts)

    assertEquals(normalize(expected.toXml), normalize(actual))
  }

  test("formula without cached value omits type attribute (no Excel corruption)") {
    // Regression test: formulas without cached values were written with t="str"
    // which caused Excel to show a "We found a problem" recovery dialog.
    // The fix is to omit the type attribute entirely.
    val a1 = ARef.from1(1, 1)
    val sheet = Sheet(SheetName.unsafe("Test"))
      .put(a1, CellValue.Formula("SUM(B1:B10)", None)) // No cached value

    val xml = emitDirect(sheet, None)
    val cellElem = (xml \\ "c").head

    // Should NOT have t="str" - either no t attribute or empty
    val typeAttr = cellElem.attribute("t").map(_.text)
    assert(
      typeAttr.isEmpty || typeAttr.contains(""),
      s"Formula without cached value should not have type attribute, got: $typeAttr"
    )

    // Should have formula element
    val formulaElem = (cellElem \ "f").headOption
    assert(formulaElem.isDefined, "Cell should have <f> element")
    assertEquals(formulaElem.get.text, "SUM(B1:B10)")
  }

  private def emitDirect(sheet: Sheet, tableParts: Option[Elem]): Elem =
    val output = new ByteArrayOutputStream()
    val writer = StaxSaxWriter.create(output)
    DirectSaxEmitter.emitWorksheet(writer, sheet, None, Map.empty, tableParts)
    val xml = new String(output.toByteArray, StandardCharsets.UTF_8)
    XML.loadString(xml)

  private def buildSheet(): Sheet =
    val a1 = ARef.from1(1, 1)
    val b1 = ARef.from1(2, 1)
    val c1 = ARef.from1(3, 1)
    val d1 = ARef.from1(4, 1)
    val a2 = ARef.from1(1, 2)
    val b2 = ARef.from1(2, 2)

    val rich = "Bold".bold.red + " plain"

    Sheet(SheetName.unsafe("Sheet1"))
      .put(a1, CellValue.Text("  spaced"))
      .put(b1, CellValue.Number(BigDecimal(42)))
      .put(c1, CellValue.Bool(true))
      .put(
        d1,
        CellValue.Formula("SUM(A1:B1)", Some(CellValue.Number(BigDecimal(3))))
      )
      .put(a2, CellValue.RichText(rich))
      .put(b2, CellValue.DateTime(LocalDateTime.of(2024, 1, 2, 3, 4)))
      .merge(CellRange(ARef.from1(1, 3), ARef.from1(2, 3)))
      .comment(a1, Comment.plainText("Note", Some("QA")))
      .setRowProperties(
        Row.from1(2),
        RowProperties(height = Some(24.0), hidden = true, outlineLevel = Some(1), collapsed = true)
      )
      .setColumnProperties(
        Column.from1(1),
        ColumnProperties(width = Some(12.5), outlineLevel = Some(1))
      )
      .setColumnProperties(
        Column.from1(2),
        ColumnProperties(width = Some(12.5), outlineLevel = Some(1))
      )
      .setColumnProperties(
        Column.from1(3),
        ColumnProperties(width = Some(20.0), hidden = true, outlineLevel = Some(2), collapsed = true)
      )

  private def buildTableParts(): Elem =
    val tablePart =
      scala.xml.Elem(
        prefix = null,
        label = "tablePart",
        attributes = new PrefixedAttribute("r", "id", "rId1", Null),
        scope = scala.xml.NamespaceBinding("r", XmlUtil.nsRelationships, TopScope),
        minimizeEmpty = true
      )
    XmlUtil.elem("tableParts", "count" -> "1")(tablePart)

  private def normalize(elem: Elem): Elem =
    val attrs = normalizeAttributes(elem.attributes)
    val children = elem.child.flatMap {
      case t: Text if t.text.forall(_.isWhitespace) => None
      case e: Elem => Some(normalize(e))
      case other => Some(other)
    }
    elem.copy(prefix = null, scope = TopScope, attributes = attrs, child = children)

  private def normalizeAttributes(attrs: MetaData): MetaData =
    val sorted = attributePairs(attrs)
      .filterNot { case (key, _) => key == "xmlns" || key.startsWith("xmlns:") }
      .sortBy(_._1)
    sorted.foldLeft(Null: MetaData) { case (acc, (key, value)) =>
      key.split(":", 2) match
        case Array(prefix, local) => new PrefixedAttribute(prefix, local, value, acc)
        case _ => new UnprefixedAttribute(key, value, acc)
    }

  private def attributePairs(attrs: MetaData): List[(String, String)] =
    attrs match
      case Null => Nil
      case PrefixedAttribute(pre, key, value, next) =>
        (s"$pre:$key" -> value.text) :: attributePairs(next)
      case UnprefixedAttribute(key, value, next) =>
        (key -> value.text) :: attributePairs(next)
