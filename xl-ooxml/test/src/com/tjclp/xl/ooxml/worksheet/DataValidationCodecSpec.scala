package com.tjclp.xl.ooxml.worksheet

import scala.xml.Elem

import munit.FunSuite

import com.tjclp.xl.api.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.sheets.{DataValidation, DvKind}

/**
 * GH-375: total data-validation codec — typed parse of the list family with entry-level Preserved
 * fallback, canonical single-container emission, and `parseAll(toElem(dvs)) == dvs` for typed
 * content (the CfCodecSpec contract).
 */
class DataValidationCodecSpec extends FunSuite:

  private def xml(s: String): Elem =
    XmlSecurity.parseSafe(s, "test").fold(e => fail(s"parse failed: ${e.message}"), identity)

  private def roundTrip(dvs: Vector[DataValidation]): Vector[DataValidation] =
    DataValidationCodec.parseAll(DataValidationCodec.toElem(dvs, None).map(e => xml(e.toString)))

  private def typedList(
    range: String,
    formula: String,
    allowBlank: Boolean = true,
    showDropdown: Boolean = true
  ): DataValidation.Rules =
    DataValidation.Rules(
      Vector(CellRange.parse(range).fold(e => fail(s"bad range: $e"), identity)),
      DvKind.List(formula),
      allowBlank,
      showDropdown
    )

  // ===== typed round-trips =====

  test("list round-trips: inline values, range ref, flag combinations") {
    val dvs: Vector[DataValidation] = Vector(
      typedList("B2:B10", "\"Low,Med,High\""),
      typedList("C2:C10", "$Z$1:$Z$3", allowBlank = false),
      typedList("D2:D10", "Lists!$A$1:$A$5", showDropdown = false),
      typedList("E2:E10", "\"1,2\"", allowBlank = false, showDropdown = false)
    )
    assertEquals(roundTrip(dvs), dvs)
  }

  test("multi-range sqref round-trips (space-separated)") {
    val dv = DataValidation.Rules(
      Vector(ref"A1:A5": CellRange, ref"C1:C5": CellRange),
      DvKind.List("\"x\"")
    )
    assertEquals(roundTrip(Vector(dv)), Vector(dv))
    val emitted = DataValidationCodec.toElem(Vector(dv), None).getOrElse(fail("no container"))
    assertEquals((emitted \ "dataValidation").headOption.map(_ \@ "sqref"), Some("A1:A5 C1:C5"))
  }

  test("the showDropDown INVERSION maps both directions") {
    // raw attr present=1 -> model false; absent -> model true
    val container = xml(
      """<dataValidations count="2">
        |  <dataValidation type="list" sqref="A1:A3" showDropDown="1"><formula1>"a"</formula1></dataValidation>
        |  <dataValidation type="list" sqref="B1:B3"><formula1>"b"</formula1></dataValidation>
        |</dataValidations>""".stripMargin
    )
    val parsed = DataValidationCodec.parseAll(Some(container))
    assertEquals(
      parsed.collect { case r: DataValidation.Rules => r.showDropdown },
      Vector(false, true)
    )
    // model false -> attr "1"; model true -> attr absent
    val emitted = DataValidationCodec.toElem(parsed, None).getOrElse(fail("no container"))
    val entries = (emitted \ "dataValidation").collect { case e: Elem => e }
    assertEquals(entries.map(e => e.attribute("showDropDown").map(_.text)), Seq(Some("1"), None))
  }

  test("single-cell sqref promotes to a 1x1 range") {
    val container = xml(
      """<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>"""
    )
    DataValidationCodec.parseAll(Some(container)) match
      case Vector(DataValidation.Rules(ranges, DvKind.List(f), allowBlank, showDropdown)) =>
        assertEquals(ranges.map(_.toA1), Vector("A1:A1"))
        assertEquals(f, "\"Yes,No\"")
        assertEquals(allowBlank, false) // schema default when the attribute is absent
        assertEquals(showDropdown, true)
      case other => fail(s"expected one typed list, got $other")
  }

  // ===== Preserved fallback (total parse) =====

  test("unmodeled entries ride Preserved: foreign type, extra attrs, formula2, no formula") {
    val container = xml(
      """<dataValidations count="4">
        |  <dataValidation type="whole" operator="between" sqref="A1"><formula1>1</formula1><formula2>9</formula2></dataValidation>
        |  <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B1"><formula1>"a,b"</formula1></dataValidation>
        |  <dataValidation type="list" sqref="C1"><formula1>"a"</formula1><formula2>"b"</formula2></dataValidation>
        |  <dataValidation type="list" sqref="D1"/>
        |</dataValidations>""".stripMargin
    )
    val parsed = DataValidationCodec.parseAll(Some(container))
    assertEquals(parsed.size, 4)
    assert(
      parsed.forall(_.isInstanceOf[DataValidation.Preserved]),
      s"all four must degrade to Preserved: $parsed"
    )
    // and they re-emit verbatim (payload is canonical self-contained XML)
    val emitted = DataValidationCodec.toElem(parsed, None).getOrElse(fail("no container"))
    assertEquals((emitted \ "dataValidation").size, 4)
    assertEquals(emitted \@ "count", "4")
    assert(emitted.toString.contains("showInputMessage=\"1\""), emitted.toString)
    assertEquals(DataValidationCodec.parseAll(Some(xml(emitted.toString))), parsed)
  }

  test("corrupt sqref degrades to Preserved (never fails the read)") {
    val container = xml(
      """<dataValidations count="1"><dataValidation type="list" sqref="NOT A REF"><formula1>"a"</formula1></dataValidation></dataValidations>"""
    )
    DataValidationCodec.parseAll(Some(container)) match
      case Vector(p: DataValidation.Preserved) => assert(p.xml.contains("NOT A REF"))
      case other => fail(s"expected Preserved, got $other")
  }

  // ===== emission =====

  test("empty model emits no container; empty-ranges typed entries are dropped") {
    assertEquals(DataValidationCodec.toElem(Vector.empty, None), None)
    assertEquals(
      DataValidationCodec.toElem(Vector(DataValidation.list("\"a\"")), None),
      None,
      "an entry with no ranges is unexpressible and must drop"
    )
  }

  test("a hand-built non-XML Preserved payload drops silently (CfCodec contract)") {
    val dvs: Vector[DataValidation] =
      Vector(DataValidation.Preserved("not xml <<<"), typedList("A1:A2", "\"a\""))
    val emitted = DataValidationCodec.toElem(dvs, None).getOrElse(fail("no container"))
    assertEquals((emitted \ "dataValidation").size, 1)
    assertEquals(emitted \@ "count", "1")
  }

  test("dirty rebuild overlays unmodeled container attrs from the source and restamps count") {
    val base = xml(
      """<dataValidations count="7" disablePrompts="1" xWindow="120"><dataValidation type="list" sqref="A1"><formula1>"a"</formula1></dataValidation></dataValidations>"""
    )
    val emitted = DataValidationCodec
      .toElem(Vector(typedList("B2:B4", "\"x,y\"")), Some(base))
      .getOrElse(fail("no container"))
    assertEquals(emitted \@ "count", "1", "count restamped to the emitted children")
    assertEquals(emitted \@ "disablePrompts", "1", "unmodeled container attr dropped")
    assertEquals(emitted \@ "xWindow", "120", "unmodeled container attr dropped")
    assertEquals((emitted \ "dataValidation").size, 1)
    assertEquals((emitted \ "dataValidation").headOption.map(_ \@ "sqref"), Some("B2:B4"))
  }

  test("allowBlank defaults: absent attr parses false; model true emits allowBlank=\"1\"") {
    val emitted = DataValidationCodec
      .toElem(Vector(typedList("A1:A2", "\"a\"", allowBlank = true)), None)
      .getOrElse(fail("no container"))
    assertEquals(
      (emitted \ "dataValidation").headOption.map(_ \@ "allowBlank"),
      Some("1")
    )
    val emittedFalse = DataValidationCodec
      .toElem(Vector(typedList("A1:A2", "\"a\"", allowBlank = false)), None)
      .getOrElse(fail("no container"))
    assertEquals(
      (emittedFalse \ "dataValidation").flatMap(_.attribute("allowBlank")).isEmpty,
      true,
      "false is the schema default and must omit the attribute"
    )
  }
