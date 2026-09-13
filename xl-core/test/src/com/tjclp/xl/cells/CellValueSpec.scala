package com.tjclp.xl.cells

import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.macros.ref
import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen}
import org.scalacheck.Prop.forAll

/**
 * GH-479: the model's canonical formula text. `CellValue.canonicalFormulaText` is the ONE strip
 * every canonical entry shares — exactly one leading '=' removed, nothing else — and
 * `CellValue.formula` / `FormulaParser.parse` are the validated entries that apply it.
 */
class CellValueSpec extends ScalaCheckSuite:

  // Any text that does not itself start with '=' is already canonical
  private val genBare: Gen[String] = Arbitrary.arbitrary[String].filterNot(_.startsWith("="))

  property("canonicalFormulaText strips exactly the display form's leading '='") {
    forAll(genBare) { (s: String) =>
      assertEquals(CellValue.canonicalFormulaText("=" + s), s)
      assertEquals(CellValue.canonicalFormulaText(s), s)
    }
  }

  property("canonicalFormulaText is idempotent on the display and bare forms") {
    // A formula has two shapes, `=A1` and `A1`; each is a fixed point once canonical. A doubled '='
    // is deliberately NOT one (pinned below), so this law is stated over the two shapes, not all
    // strings.
    forAll(genBare) { (s: String) =>
      val once = CellValue.canonicalFormulaText("=" + s)
      assertEquals(once, s)
      assertEquals(CellValue.canonicalFormulaText(once), once)
      assertEquals(CellValue.canonicalFormulaText(s), s)
    }
  }

  property("canonicalFormulaText removes at most one character, always a leading '='") {
    forAll { (s: String) =>
      val out = CellValue.canonicalFormulaText(s)
      if s.startsWith("=") then assertEquals(out, s.drop(1)) else assertEquals(out, s)
    }
  }

  test("canonicalFormulaText: a leading '+', interior '=', doubled '=' and whitespace survive") {
    assertEquals(CellValue.canonicalFormulaText("=+SUM(A1:B2)"), "+SUM(A1:B2)")
    assertEquals(CellValue.canonicalFormulaText("=IF(A1=1,2,3)"), "IF(A1=1,2,3)")
    assertEquals(CellValue.canonicalFormulaText("==A1"), "=A1")
    // A doubled '=' is not a fixed point BY DESIGN: one strip per entry; every writer heals the rest
    // at the `<f>` boundary (FormulaLeadingEqualsSpec pins that read-back as "A1").
    assertNotEquals(
      CellValue.canonicalFormulaText(CellValue.canonicalFormulaText("==A1")),
      CellValue.canonicalFormulaText("==A1")
    )
    assertEquals(CellValue.canonicalFormulaText(" =A1 "), " =A1 ")
    assertEquals(CellValue.canonicalFormulaText("="), "")
    assertEquals(CellValue.canonicalFormulaText(""), "")
  }

  test("CellValue.formula stores the canonical text for either input shape") {
    val a1: XLResult[CellValue.Formula] = Right(CellValue.Formula("A1"))
    assertEquals(CellValue.formula("=A1"), a1)
    assertEquals(CellValue.formula("A1"), a1)
    val cached = Some(CellValue.Number(BigDecimal(2)))
    assertEquals(
      CellValue.formula("=A1*2", cached, FormulaKind.Normal(ca = true)),
      Right(CellValue.Formula("A1*2", cached, FormulaKind.Normal(ca = true))): XLResult[
        CellValue.Formula
      ]
    )
  }

  test("CellValue.formula is total: an empty expression (also a lone '=') is a Left naming it") {
    assertEquals(
      CellValue.formula(""),
      Left(XLError.FormulaError("", "Formula expression cannot be empty"))
    )
    assertEquals(
      CellValue.formula("="),
      Left(XLError.FormulaError("=", "Formula expression cannot be empty"))
    )
  }

  test("CellValue.formula refuses a Formula as the cached value") {
    assertEquals(
      CellValue.formula("=A1", Some(CellValue.Formula("B1"))),
      Left(XLError.FormulaError("=A1", "Cached value cannot be a Formula"))
    )
  }

  private val table = FormulaKind.DataTable(
    ref = ref"B2:C3",
    dt2D = true,
    dtr = true,
    r1 = Some(ref"A1"),
    r2 = Some(ref"A2"),
    ca = true
  )

  test("CellValue.formula on a DataTable kind requires the derived TABLE(...) text") {
    val derived = FormulaKind.displayExpression(table)
    assertEquals(derived, "TABLE(A1,A2)")
    val expected: XLResult[CellValue.Formula] = Right(CellValue.Formula(derived, None, table))
    assertEquals(CellValue.formula(derived, None, table), expected)
    // the display form of the derived text is accepted and canonicalized too
    assertEquals(CellValue.formula("=" + derived, None, table), expected)
    assert(CellValue.formula("A1*2", None, table).isLeft)
    assertEquals(CellValue.dataTable(table), CellValue.Formula(derived, None, table))
  }

  test("FormulaParser.parse stores the canonical text; both shapes yield the same value") {
    assertEquals(FormulaParser.parse("=X"), Right(CellValue.Formula("X")))
    assertEquals(FormulaParser.parse("X"), FormulaParser.parse("=X"))
    assertEquals(FormulaParser.parse("==A1"), Right(CellValue.Formula("=A1")))
    assertEquals(FormulaParser.parse("=+SUM(A1:B2)"), Right(CellValue.Formula("+SUM(A1:B2)")))
  }

  test("FormulaParser.parse rejects the empty formula and a lone '=', reporting the ORIGINAL") {
    assertEquals(
      FormulaParser.parse(""),
      Left(XLError.FormulaError("", "Formula cannot be empty"))
    )
    assertEquals(
      FormulaParser.parse("="),
      Left(XLError.FormulaError("=", "Formula cannot be empty"))
    )
    assertEquals(
      FormulaParser.parse("=SUM(A1"),
      Left(XLError.FormulaError("=SUM(A1", "Unbalanced parentheses"))
    )
  }

  property("FormulaParser.parse: a parsed formula's text is already canonical") {
    forAll(genBare) { (s: String) =>
      FormulaParser.parse("=" + s) match
        case Right(CellValue.Formula(expr, _, _)) =>
          assertEquals(expr, s)
          assertEquals(CellValue.canonicalFormulaText(expr), expr)
        case Right(other) => fail(s"expected a Formula, got $other")
        case Left(_) => () // unbalanced parentheses / empty / over-limit inputs are refused
    }
  }
