package com.tjclp.xl.ooxml.worksheet

import scala.xml.Elem

import munit.ScalaCheckSuite
import org.scalacheck.Prop.forAll

import com.tjclp.xl.api.*
import com.tjclp.xl.Generators.{genDvRules, genDvText}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.{XmlSecurity, XmlUtil}
import com.tjclp.xl.sheets.{
  DataValidation,
  DvBoundedType,
  DvErrorStyle,
  DvKind,
  DvMessages,
  DvOperator
}

/**
 * GH-375 (widened by GH-429): total data-validation codec — typed parse of the
 * list/custom/anyValue/bounded families with prompt/error messages, entry-level Preserved fallback,
 * canonical single-container emission, and `parseAll(toElem(dvs)) == dvs` for typed content (the
 * CfCodecSpec contract).
 */
class DataValidationCodecSpec extends ScalaCheckSuite:

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
      case Vector(DataValidation.Rules(ranges, DvKind.List(f), allowBlank, showDropdown, msgs)) =>
        assertEquals(ranges.map(_.toA1), Vector("A1:A1"))
        assertEquals(f, "\"Yes,No\"")
        assertEquals(allowBlank, false) // schema default when the attribute is absent
        assertEquals(showDropdown, true)
        assertEquals(msgs, DvMessages.default)
      case other => fail(s"expected one typed list, got $other")
  }

  // ===== GH-429: the widened typed subset =====

  test("GH-429: the field-repro entry (Excel-stamped prompt/error flags) parses TYPED") {
    val container = xml(
      """<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="H5:H10"><formula1>"yes,no"</formula1></dataValidation></dataValidations>"""
    )
    DataValidationCodec.parseAll(Some(container)) match
      case Vector(r: DataValidation.Rules) =>
        assertEquals(r.ranges.map(_.toA1), Vector("H5:H10"))
        assertEquals(r.kind, DvKind.List("\"yes,no\""))
        assertEquals(r.allowBlank, true)
        assertEquals(r.showDropdown, true)
        assertEquals(
          r.messages,
          DvMessages(showInputMessage = true, showErrorMessage = true)
        )
      case other => fail(s"the Excel field shape must parse typed, got $other")
  }

  test("GH-429: bounded, custom, and message-only anyValue entries parse typed and round-trip") {
    val dvs: Vector[DataValidation] = Vector(
      DataValidation.Rules(
        Vector(ref"A1:A9": CellRange),
        DvKind.Bounded(DvBoundedType.Whole, DvOperator.Between, "1", Some("9")),
        allowBlank = false
      ),
      DataValidation.Rules(
        Vector(ref"B1:B9": CellRange),
        DvKind.Bounded(DvBoundedType.Date, DvOperator.GreaterThan, "DATE(2020,1,1)", None),
        messages = DvMessages(showErrorMessage = true, errorStyle = DvErrorStyle.Warning)
      ),
      DataValidation.Rules(Vector(ref"C1:C9": CellRange), DvKind.Custom("ISNUMBER(C1)")),
      DataValidation.Rules(
        Vector(ref"D1:D9": CellRange),
        DvKind.AnyValue,
        messages = DvMessages(
          showInputMessage = true,
          promptTitle = Some("Hint"),
          prompt = Some("Anything goes")
        )
      )
    )
    assertEquals(roundTrip(dvs), dvs)
    // operator/errorStyle/type omit their schema defaults on the wire
    val emitted = DataValidationCodec.toElem(dvs, None).getOrElse(fail("no container"))
    val entries = (emitted \ "dataValidation").collect { case e: Elem => e }
    assertEquals(entries.map(_ \@ "operator"), Seq("", "greaterThan", "", ""))
    assertEquals(entries.map(_ \@ "errorStyle"), Seq("", "warning", "", ""))
    assertEquals(entries.map(_ \@ "type"), Seq("whole", "date", "custom", ""))
  }

  property("GH-429: kind x operator x messages round-trip (parseAll . toElem = id)") {
    forAll(genDvRules) { rules =>
      roundTrip(Vector(rules)) == Vector[DataValidation](rules)
    }
  }

  test("GH-429: a multiline prompt round-trips through _x000A_ (attr normalization survives)") {
    val dv = DataValidation.Rules(
      Vector(ref"A1:A2": CellRange),
      DvKind.List("\"a,b\""),
      messages = DvMessages(
        showInputMessage = true,
        promptTitle = Some("Two"),
        prompt = Some("line1\nline2\tand tab")
      )
    )
    val emitted = DataValidationCodec.toElem(Vector(dv), None).getOrElse(fail("no container"))
    assert(emitted.toString.contains("_x000A_"), emitted.toString)
    assert(emitted.toString.contains("_x0009_"), emitted.toString)
    assertEquals(roundTrip(Vector(dv)), Vector[DataValidation](dv))
  }

  property("GH-429: decodeXstring . escapeXstringAttr = id") {
    forAll { (s: String) =>
      XmlUtil.decodeXstring(XmlUtil.escapeXstringAttr(s)) == s
    }
  }

  property("GH-429: escaped attr text round-trips generator hazards") {
    forAll(genDvText) { s =>
      XmlUtil.decodeXstring(XmlUtil.escapeXstringAttr(s)) == s
    }
  }

  // ===== Preserved fallback (total parse) =====

  test("unmodeled entries ride Preserved: foreign attrs, list formula2, no formula") {
    val container = xml(
      """<dataValidations count="4" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
        |  <dataValidation type="whole" operator="between" imeMode="hiragana" sqref="A1"><formula1>1</formula1><formula2>9</formula2></dataValidation>
        |  <dataValidation type="list" xr:uid="{00000000-0002-0000-0000-000000000000}" sqref="B1"><formula1>"a,b"</formula1></dataValidation>
        |  <dataValidation type="list" sqref="C1"><formula1>"a"</formula1><formula2>"b"</formula2></dataValidation>
        |  <dataValidation type="unknownKind" sqref="D1"/>
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
    assert(emitted.toString.contains("imeMode=\"hiragana\""), emitted.toString)
    assertEquals(DataValidationCodec.parseAll(Some(xml(emitted.toString))), parsed)
  }

  test("GH-429: an operator on a non-bounded type stays Preserved (no parse-and-drop)") {
    val container = xml(
      """<dataValidations count="1"><dataValidation type="list" operator="equal" sqref="A1"><formula1>"a"</formula1></dataValidation></dataValidations>"""
    )
    DataValidationCodec.parseAll(Some(container)) match
      case Vector(p: DataValidation.Preserved) => assert(p.xml.contains("operator"))
      case other => fail(s"expected Preserved, got $other")
  }

  // ===== GH-429: commuting square — textual Preserved shift == typed envelope shift =====

  property("GH-429: parse(shiftPayload(emit(rules))) == shiftTyped(rules) over rules x edits") {
    import org.scalacheck.Gen
    import com.tjclp.xl.addressing.{ARef as XARef, Row as XRow, SheetName}
    import com.tjclp.xl.sheets.{Sheet as CoreSheet, SqrefShift}
    val genEdit = for
      at <- Gen.chooseNum(0, 120)
      count <- Gen.chooseNum(1, 5)
      deleting <- Gen.oneOf(true, false)
    yield (at, count, deleting)
    forAll(genDvRules, genEdit) { case (rules, (at, count, deleting)) =>
      // typed side: the Sheet envelope shift
      val base = CoreSheet(SheetName.unsafe("S")).copy(dataValidations = Vector(rules))
      val typedShifted =
        (if deleting then base.deleteRows(at, count)
         else base.insertRows(at, count)).dataValidations
      // textual side: the same span algebra applied to the EMITTED payload string
      def shift(r: CellRange): Option[CellRange] =
        CoreSheet
          .shiftSpan(r.start.row.index0, r.end.row.index0, at, count, deleting, XRow.MaxIndex0)
          .map((ns, ne) =>
            CellRange(
              XARef.from0(r.start.col.index0, ns),
              XARef.from0(r.end.col.index0, ne)
            )
          )
      // derive the payload EXACTLY like parseEntry's Preserved fallback does — declaration
      // prologue and all (a bare Elem.toString is not the production shape)
      val entryXml = DataValidationCodec
        .toElem(Vector(rules), None)
        .toList
        .flatMap(c =>
          (c \ "dataValidation").collect { case e: Elem => CfCodec.preservedXml(e, c.scope) }
        )
      entryXml match
        case entry :: Nil =>
          val textualShifted = SqrefShift
            .shiftPayload(entry, "dataValidation", shift)
            .toList
            .flatMap { x =>
              // payloads parse standalone (declaration first), then re-enter via a container
              val e = xml(x)
              DataValidationCodec.parseAll(Some(<dataValidations>{e}</dataValidations>))
            }
          textualShifted == typedShifted.toList
        case _ => false
    }
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
