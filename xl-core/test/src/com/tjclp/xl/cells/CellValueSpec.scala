package com.tjclp.xl.cells

import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.macros.ref
import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen}
import org.scalacheck.Prop.forAll

/**
 * GH-479: the model's canonical formula text. `CellValue.canonicalFormulaText` is the ONE rule
 * every canonical entry shares — surrounding whitespace trimmed, exactly one leading '=' removed,
 * trimmed again; nothing interior changes — and `CellValue.formula` / `FormulaParser.parse` are the
 * validated entries that apply it.
 */
class CellValueSpec extends ScalaCheckSuite:

  // Any trimmed text that does not itself start with '=' is already canonical
  private val genBare: Gen[String] =
    Arbitrary.arbitrary[String].map(_.trim).filterNot(_.startsWith("="))

  // Whitespace a shell quote, a JSON string or a script literal may carry around a formula
  private val genPad: Gen[String] = Gen.stringOf(Gen.oneOf(' ', '\t', '\n', '\r'))

  property("canonicalFormulaText strips exactly the display form's leading '='") {
    forAll(genBare) { (s: String) =>
      assertEquals(CellValue.canonicalFormulaText("=" + s), s)
      assertEquals(CellValue.canonicalFormulaText(s), s)
    }
  }

  property("canonicalFormulaText: surrounding whitespace never reaches the model") {
    // The one rule at every entry (fx, FormulaParser.parse, formula, putFormulaInheriting, the
    // edit interpreter, the CLI): `" = SUM(A1) "` and `"=SUM(A1)"` are the same formula.
    forAll(genBare, genPad, genPad, genPad) { (s: String, a: String, b: String, c: String) =>
      assertEquals(CellValue.canonicalFormulaText(a + "=" + b + s + c), s)
      assertEquals(CellValue.canonicalFormulaText(a + s + c), s)
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

  property("canonicalFormulaText is trim, one leading '=' off, trim — and nothing interior") {
    forAll { (s: String) =>
      val out = CellValue.canonicalFormulaText(s)
      val trimmed = s.trim
      assertEquals(out, out.trim)
      if trimmed.startsWith("=") then assertEquals(out, trimmed.drop(1).trim)
      else assertEquals(out, trimmed)
    }
  }

  test(
    "canonicalFormulaText: a leading '+', interior '=', doubled '=' and interior spaces survive"
  ) {
    assertEquals(CellValue.canonicalFormulaText("=+SUM(A1:B2)"), "+SUM(A1:B2)")
    assertEquals(CellValue.canonicalFormulaText("=IF(A1=1,2,3)"), "IF(A1=1,2,3)")
    assertEquals(CellValue.canonicalFormulaText("==A1"), "=A1")
    // A doubled '=' is not a fixed point BY DESIGN: one strip per entry; every writer heals the rest
    // at the `<f>` boundary (FormulaLeadingEqualsSpec pins that read-back as "A1").
    assertNotEquals(
      CellValue.canonicalFormulaText(CellValue.canonicalFormulaText("==A1")),
      CellValue.canonicalFormulaText("==A1")
    )
    assertEquals(CellValue.canonicalFormulaText(" =A1 "), "A1")
    assertEquals(CellValue.canonicalFormulaText(" = SUM( A1 , B1 ) "), "SUM( A1 , B1 )")
    assertEquals(CellValue.canonicalFormulaText("\t=A1\n"), "A1")
    assertEquals(CellValue.canonicalFormulaText("="), "")
    assertEquals(CellValue.canonicalFormulaText(" = "), "")
    assertEquals(CellValue.canonicalFormulaText("   "), "")
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
    // whitespace around a lone '=' (or nothing) is still empty, and the Left names the ORIGINAL
    assertEquals(
      CellValue.formula(" = "),
      Left(XLError.FormulaError(" = ", "Formula expression cannot be empty"))
    )
  }

  test("CellValue.formula trims surrounding whitespace before and after the strip") {
    val a1: XLResult[CellValue.Formula] = Right(CellValue.Formula("SUM( A1:A10 )"))
    assertEquals(CellValue.formula(" =SUM( A1:A10 ) "), a1)
    assertEquals(CellValue.formula(" = SUM( A1:A10 ) "), a1)
    assertEquals(CellValue.formula("SUM( A1:A10 )\n"), a1)
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

  test("CellValue.dataTable never carries a Formula as its cache (the record caches scalars)") {
    // The total constructor keeps the invariant `formula` refuses with a Left: a Formula offered as
    // the cache is dropped, so the record is the same cache-less record every writer would emit.
    val derived = FormulaKind.displayExpression(table)
    val scalar = Some(CellValue.Number(BigDecimal(42)))
    assertEquals(CellValue.dataTable(table, scalar), CellValue.Formula(derived, scalar, table))
    assertEquals(
      CellValue.dataTable(table, Some(CellValue.Formula("A1"))),
      CellValue.Formula(derived, None, table)
    )
    assertEquals(
      CellValue.dataTable(table, Some(CellValue.dataTable(table, scalar))),
      CellValue.Formula(derived, None, table)
    )
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
