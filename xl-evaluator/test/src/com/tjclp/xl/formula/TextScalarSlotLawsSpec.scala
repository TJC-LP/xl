package com.tjclp.xl.formula

import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.Generators.genWideBigDecimal
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.formula.ast.{BindingCoercion, TExpr}
import com.tjclp.xl.formula.eval.{EvalError, Evaluator, ScalarCoercion}
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.syntax.*

/**
 * #709: one law for text in a numeric slot. A text cell read through a single reference
 * ([[TExpr.decodeNumericScalar]] at the evaluator boundary) and the same text as a runtime value
 * (`ScalarCoercion.coerce(…, Numeric)`, the literal, call-result and LET path) read the same
 * number, or are both `#VALUE!`. Plain and rich text alike: rich text reads through its plain text
 * on both sides. No class of text is excluded — both sides parse with
 * `ScalarCoercion.parseNumericText`.
 */
class TextScalarSlotLawsSpec extends ScalaCheckSuite:

  /** Keep extreme exponents covered without constructing and parsing million-digit integers. */
  private def boundedPlain(n: BigDecimal): String =
    if n.scale >= -256 && n.scale <= 256 then n.bigDecimal.toPlainString else n.toString

  /** Text Excel would and would not read as a number, plus whatever ScalaCheck invents. */
  private val genText: Gen[String] =
    val pad = Gen.oneOf("", " ", "  ", "\t")
    val numeric = for
      n <- genWideBigDecimal
      neg <- Gen.oneOf(true, false)
      form <- Gen.oneOf[BigDecimal => String](_.toString, boundedPlain)
      before <- pad
      after <- pad
    yield before + form(if neg then -n else n) + after
    Gen.frequency(
      4 -> numeric,
      2 -> Gen.choose(-100000, 100000).map(_.toString),
      3 -> Arbitrary.arbitrary[String],
      2 -> Gen.oneOf(
        "",
        "abc",
        "TRUE",
        "false",
        "1e400",
        "1E-400",
        "1e2147483648",
        "NaN",
        "Infinity",
        "0x10",
        "1,000",
        "$5",
        "50%",
        "2024-01-15",
        "５",
        "+5",
        ".5",
        "5.",
        "- 5"
      )
    )

  private val evaluator = Evaluator.instance

  private def throughReference(cell: CellValue): Either[EvalError, BigDecimal] =
    val sheet = Sheet(SheetName.unsafe("Sheet1")).put(ref"A1", cell)
    evaluator.eval(TExpr.ref(ref"A1", TExpr.decodeNumericScalar), sheet)

  private def isValueError(r: Either[EvalError, Any]): Boolean = r match
    case Left(EvalError.ErrorValue(CellError.Value, _)) => true
    case _ => false

  property("#709: a text cell in a numeric slot reads as the same text as a runtime value") {
    forAll(genText, Gen.oneOf(true, false)) { (s, rich) =>
      val cell = if rich then CellValue.RichText(RichText.plain(s)) else CellValue.Text(s)
      val byRef = throughReference(cell)
      val byValue = ScalarCoercion.coerce("x", cell, BindingCoercion.Numeric)
      val agree = (byRef, byValue) match
        case (Right(a), Right(b)) => b == a
        case (l, r) => isValueError(l) && isValueError(r)
      assert(agree, s"text ${s.toList.map(_.toInt)}: reference $byRef, value $byValue")
    }
  }
