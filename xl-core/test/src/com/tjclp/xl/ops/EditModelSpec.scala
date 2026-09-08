package com.tjclp.xl.ops

import scala.compiletime.constValueTuple
import scala.deriving.Mirror

import com.tjclp.xl.Generators.{genCellStyle, genSheetName}
import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row, SheetName}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.border.{BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen}
import org.scalacheck.Prop.forAll

/**
 * W2.1 (ADR-017 §2.12), the model half: `Edit` cases and their `EditSchema` rows agree one-to-one,
 * the target parsers are total over `RefType.parseToXLError`, `Scope` follows rule 4, `FormatHint`
 * enforces invariant 4 (GH-560), `StyleOverlay` is a monoid acting on `CellStyle`, and the
 * text-only `FormulaSupport` refuses what it cannot do.
 */
class EditModelSpec extends ScalaCheckSuite:

  private inline def caseLabels[T](using m: Mirror.SumOf[T]): List[String] =
    constValueTuple[m.MirroredElemLabels].toList.map(_.toString)

  private def sn(s: String): SheetName = SheetName.unsafe(s)
  private def a1(s: String): ARef = ARef.parse(s).fold(e => fail(e), identity)
  private def rng(s: String): CellRange = CellRange.parse(s).fold(e => fail(e), identity)

  // ========== EditSchema ==========

  test("EditSchema.all has exactly one row per Edit case, in declaration order") {
    val labels = caseLabels[Edit]
    assertEquals(EditSchema.all.size, labels.size)
    assertEquals(EditSchema.all.map(_.name).toList, labels.map(EditSchema.kebab))
  }

  test("EditSchema names are unique and resolve through find (name, batch op, any spelling)") {
    val names = EditSchema.all.map(_.name)
    assertEquals(names.distinct, names)
    assertEquals(EditSchema.find("put-formula").map(_.name), Some("put-formula"))
    assertEquals(EditSchema.find("putFormula").map(_.name), Some("put-formula"))
    assertEquals(EditSchema.find("PUT_FORMULA").map(_.name), Some("put-formula"))
    // The batch op name resolves to the FIRST case that is a form of it
    assertEquals(EditSchema.find("putf").map(_.name), Some("put-formula"))
    assertEquals(EditSchema.find("colwidth").map(_.name), Some("col-width"))
    assertEquals(EditSchema.find("no-such-op"), None)
  }

  test("EditSchema: every Wave 1 batch op name is the twin of at least one case") {
    val wave1 = Vector(
      "put",
      "putf",
      "style",
      "merge",
      "unmerge",
      "colwidth",
      "rowheight",
      "comment",
      "remove-comment",
      "hyperlink",
      "clear",
      "col-hide",
      "col-show",
      "row-hide",
      "row-show",
      "group-rows",
      "group-cols",
      "ungroup-rows",
      "ungroup-cols",
      "autofit",
      "add-sheet",
      "rename-sheet",
      "freeze",
      "unfreeze",
      "copy",
      "chart",
      "sheet-view",
      "tab-color",
      "autofilter",
      "page-setup",
      "header-footer",
      "cf"
    )
    val twins = EditSchema.all.flatMap(_.batchOp).toSet
    assertEquals(wave1.filterNot(twins.contains), Vector.empty)
    // and the 15 twin-less verbs are all reachable as cases
    val verbs = Vector(
      "insert-rows",
      "delete-rows",
      "insert-cols",
      "delete-cols",
      "fill",
      "sort",
      "remove-sheet",
      "move-sheet",
      "copy-sheet",
      "add-image",
      "name add",
      "name remove",
      "sheets hide",
      "sheets show"
    )
    val cliVerbs = EditSchema.all.flatMap(_.cliVerb).toSet
    assertEquals(verbs.filterNot(cliVerbs.contains), Vector.empty)
  }

  test("EditSchema flags: structural cases are never streamable; formula cases need a formula") {
    EditSchema.all.foreach { spec =>
      if spec.structural then assert(!spec.streamable, s"${spec.name} structural but streamable")
    }
    assert(EditSchema.find("drag-formula").exists(_.needsFormula))
    assert(EditSchema.find("copy").exists(_.needsFormula))
    assert(EditSchema.find("insert-rows").exists(_.structural))
    assert(EditSchema.find("rename-sheet").exists(_.structural))
    assert(EditSchema.find("put").exists(_.cellMutating))
    assert(EditSchema.find("style").exists(!_.cellMutating))
    assert(EditSchema.find("add-sheet").exists(!_.sheetScoped))
  }

  // ========== Loc / Area / ColSpan / RowSpan ==========

  test("Loc.parse: bare and qualified cells; a range is refused") {
    assertEquals(Loc.parse("B7"), Right(Loc(None, a1("B7"))))
    assertEquals(Loc.parse("'Q1 Data'!B7"), Right(Loc(Some(sn("Q1 Data")), a1("B7"))))
    assertEquals(Loc.parse("Data!b7"), Right(Loc(Some(sn("Data")), a1("B7"))))
    assert(Loc.parse("A1:B2").left.exists {
      case XLError.InvalidReference(_) => true
      case _ => false
    })
    assert(Loc.parse("garbage!!").isLeft)
  }

  test("Area.parse: ranges, qualified ranges, and a single cell as a 1x1 area") {
    assertEquals(Area.parse("A1:B2"), Right(Area(None, rng("A1:B2"))))
    assertEquals(Area.parse("Data!A1:B2"), Right(Area(Some(sn("Data")), rng("A1:B2"))))
    assertEquals(Area.parse("C3"), Right(Area(None, CellRange(a1("C3"), a1("C3")))))
    assert(Area.parse("").isLeft)
  }

  property("Loc and Area render/parse round-trip") {
    val genLoc = for
      sheet <- Gen.option(genSheetName)
      col <- Gen.choose(0, 25)
      row <- Gen.choose(0, 99)
    yield Loc(sheet, ARef.from0(col, row))
    forAll(genLoc) { (loc: Loc) =>
      assertEquals(Loc.parse(loc.toA1), Right(loc))
      val area = Area(loc.sheet, CellRange(loc.ref, loc.ref.shift(1, 1)))
      assertEquals(Area.parse(area.toA1), Right(area))
      true
    }
  }

  test("ColSpan.parse: a letter, a span, a reversed span; rows and cells are refused") {
    assertEquals(ColSpan.parse("E"), Right(ColSpan(Column.from0(4), Column.from0(4))))
    assertEquals(ColSpan.parse("E:H"), Right(ColSpan(Column.from0(4), Column.from0(7))))
    assertEquals(ColSpan.parse("H:E"), Right(ColSpan(Column.from0(4), Column.from0(7))))
    assertEquals(ColSpan.parse("e:h"), Right(ColSpan(Column.from0(4), Column.from0(7))))
    assert(ColSpan.parse("5").isLeft)
    assert(ColSpan.parse("A1:B2").isLeft)
    assert(ColSpan.parse("Data!A:B").isLeft)
    assertEquals(ColSpan(Column.from0(7), Column.from0(4)).start, Column.from0(4))
    assertEquals(ColSpan(Column.from0(4), Column.from0(7)).columns.size, 4)
  }

  test("RowSpan.parse: a number, a span, a reversed span; letters and cells are refused") {
    assertEquals(RowSpan.parse("10"), Right(RowSpan(Row.from1(10), Row.from1(10))))
    assertEquals(RowSpan.parse("10:20"), Right(RowSpan(Row.from1(10), Row.from1(20))))
    assertEquals(RowSpan.parse("20:10"), Right(RowSpan(Row.from1(10), Row.from1(20))))
    assert(RowSpan.parse("0").isLeft)
    assert(RowSpan.parse("E").isLeft)
    assert(RowSpan.parse("A1:B2").isLeft)
    assert(RowSpan.parse((Row.MaxIndex0 + 2).toString).isLeft)
    assertEquals(RowSpan(Row.from1(20), Row.from1(10)).start, Row.from1(10))
    assertEquals(RowSpan(Row.from1(10), Row.from1(12)).rows.size, 3)
  }

  // ========== Scope (rule 4) ==========

  test("Scope rule 4: a rename of the default sheet retargets the edits that follow") {
    val scope = Scope(Some(sn("Old")))
    assertEquals(scope.after(Edit.RenameSheet(sn("Old"), sn("New"))), Scope(Some(sn("New"))))
    assertEquals(scope.after(Edit.RenameSheet(sn("Other"), sn("New"))), scope)
    assertEquals(Scope.none.after(Edit.RenameSheet(sn("Old"), sn("New"))), Scope.none)
    assertEquals(scope.after(Edit.Unfreeze(None)), scope)
  }

  // ========== FormatHint (invariant 4, GH-560) ==========

  test("FormatHint: Explicit replaces any numFmt, Inferred only fills a General one") {
    val currency = CellStyle.default.withNumFmt(NumFmt.Currency).withFont(Font.default.withBold())
    val custom = CellStyle.default.withNumFmt(NumFmt.Custom("0.0x"))
    val plainBold = CellStyle.default.withFont(Font.default.withBold())
    assertEquals(FormatHint.Explicit(NumFmt.Date).resolve(custom).numFmt, NumFmt.Date)
    assertEquals(
      FormatHint.Explicit(NumFmt.Date).resolve(currency),
      currency.withNumFmt(NumFmt.Date)
    )
    assertEquals(FormatHint.Inferred(NumFmt.Date).resolve(currency), currency)
    assertEquals(
      FormatHint.Inferred(NumFmt.Date).resolve(plainBold),
      plainBold.withNumFmt(NumFmt.Date)
    )
    assertEquals(FormatHint.Inferred(NumFmt.General).resolve(plainBold), plainBold)
    assertEquals(
      FormatHint.Explicit(NumFmt.General).resolve(currency),
      currency.withNumFmt(NumFmt.General)
    )
  }

  // ========== StyleOverlay ==========

  private given Arbitrary[CellStyle] = Arbitrary(genCellStyle)

  private val genOverlay: Gen[StyleOverlay] =
    for
      full <- genCellStyle
      mask <- Gen.listOfN(17, Gen.oneOf(true, false))
    yield
      val o = StyleOverlay.of(full)
      StyleOverlay(
        fontName = o.fontName.filter(_ => mask(0)),
        fontSize = o.fontSize.filter(_ => mask(1)),
        bold = o.bold.filter(_ => mask(2)),
        italic = o.italic.filter(_ => mask(3)),
        underline = o.underline.filter(_ => mask(4)),
        fontColor = o.fontColor.filter(_ => mask(5)),
        fill = o.fill.filter(_ => mask(6)),
        numFmt = o.numFmt.filter(_ => mask(7)),
        hAlign = o.hAlign.filter(_ => mask(8)),
        vAlign = o.vAlign.filter(_ => mask(9)),
        wrap = o.wrap.filter(_ => mask(10)),
        indent = o.indent.filter(_ => mask(11)),
        textRotation = o.textRotation.filter(_ => mask(12)),
        borderTop = o.borderTop.filter(_ => mask(13)),
        borderRight = o.borderRight.filter(_ => mask(14)),
        borderBottom = o.borderBottom.filter(_ => mask(15)),
        borderLeft = o.borderLeft.filter(_ => mask(16))
      )

  private given Arbitrary[StyleOverlay] = Arbitrary(genOverlay)

  property("StyleOverlay monoid: identity and associativity") {
    forAll { (a: StyleOverlay, b: StyleOverlay, c: StyleOverlay) =>
      assertEquals(StyleOverlay.empty ++ a, a)
      assertEquals(a ++ StyleOverlay.empty, a)
      assertEquals((a ++ b) ++ c, a ++ (b ++ c))
      true
    }
  }

  property("StyleOverlay action law: (a ++ b).applyTo(s) == b.applyTo(a.applyTo(s))") {
    forAll { (a: StyleOverlay, b: StyleOverlay, s: CellStyle) =>
      assertEquals((a ++ b).applyTo(s), b.applyTo(a.applyTo(s)))
      assertEquals(StyleOverlay.empty.applyTo(s), s)
      true
    }
  }

  property("StyleOverlay.of(style) reproduces the style exactly from the default") {
    forAll { (s: CellStyle) =>
      assertEquals(StyleOverlay.of(s).applyTo(CellStyle.default), s)
      true
    }
  }

  test("StyleOverlay: border colour is per-side data, not a rule over every side") {
    val red = Color.Rgb(0xffff0000)
    val o = StyleOverlay(
      borderTop = Some(BorderSide(BorderStyle.Thin, Some(red))),
      borderBottom = Some(BorderSide(BorderStyle.Thick))
    )
    val styled = o.applyTo(CellStyle.default)
    assertEquals(styled.border.top, BorderSide(BorderStyle.Thin, Some(red)))
    assertEquals(styled.border.bottom, BorderSide(BorderStyle.Thick, None))
    assertEquals(styled.border.left, BorderSide.none)
    // un-bold is expressible (Option fields), which the CLI's flag form never could say
    val unbold =
      StyleOverlay(bold = Some(false)).applyTo(CellStyle.default.withFont(Font.default.withBold()))
    assert(!unbold.font.bold)
  }

  test("StyleOverlay.validate refuses values the model's guards would throw on") {
    assert(StyleOverlay(fontSize = Some(0.0)).validate.isLeft)
    assert(StyleOverlay(fontName = Some("")).validate.isLeft)
    assert(StyleOverlay(indent = Some(-1)).validate.isLeft)
    assert(StyleOverlay(textRotation = Some(181)).validate.isLeft)
    assert(StyleOverlay(textRotation = Some(255)).validate.isRight)
    assert(StyleOverlay(bold = Some(true)).validate.isRight)
  }

  // ========== FormulaSupport.textOnly ==========

  test("FormulaSupport.textOnly validates non-empty text and refuses everything else") {
    val fs = FormulaSupport.textOnly
    val wb = Workbook(Sheet(sn("S")))
    assertEquals(fs.validate("=A1+1"), Right(()))
    assertEquals(fs.validate("SUM(A1:A3)"), Right(()))
    assert(fs.validate("").isLeft)
    assert(fs.validate("=  ").isLeft)
    def refused(r: Either[XLError, ?], op: String): Unit = r match
      case Left(XLError.UnsupportedCapability(o, capability, hint)) =>
        assertEquals(o, op)
        assertEquals(capability, "formula support")
        assert(hint.contains("EvalFormulaSupport"), hint)
      case other => fail(s"expected an UnsupportedCapability refusal for $op, got $other")
    refused(fs.shift("=A1", 1, 1), "shift")
    refused(fs.insertRows(wb, sn("S"), Row.from0(0), 1), "insert-rows")
    refused(fs.deleteRows(wb, sn("S"), Row.from0(0), 1), "delete-rows")
    refused(fs.insertCols(wb, sn("S"), Column.from0(0), 1), "insert-cols")
    refused(fs.deleteCols(wb, sn("S"), Column.from0(0), 1), "delete-cols")
    refused(fs.renameSheet(wb, sn("S"), sn("T")), "rename-sheet")
  }

  test("ClearWhat.fromFlags follows the CLI rule: all wins, no flag means contents") {
    assertEquals(
      ClearWhat.fromFlags(all = false, styles = false, comments = false),
      ClearWhat.contents
    )
    assertEquals(ClearWhat.fromFlags(all = true, styles = false, comments = false), ClearWhat.all)
    assertEquals(
      ClearWhat.fromFlags(all = false, styles = true, comments = true),
      ClearWhat(contents = false, styles = true, comments = true)
    )
    assertEquals(
      ClearWhat.fromFlags(all = false, styles = true, comments = false),
      ClearWhat.styles
    )
  }
