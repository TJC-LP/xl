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
 * #709: one law for text in a typed scalar slot. A text cell read through a single reference (the
 * slot's decoder at the evaluator boundary: [[TExpr.decodeNumericScalar]],
 * [[TExpr.decodeAsDateScalar]], [[TExpr.decodeBool]]) and the same text as a runtime value
 * (`ScalarCoercion.coerce`, the literal, call-result and LET path) read the same value, or are both
 * `#VALUE!`. Plain text, rich text and a formula's cached text alike. No class of text is excluded:
 * both sides read text through `ScalarCoercion.cellText` / `parseNumericText` / `boolTextValue`,
 * and an out-of-range date serial is the same error value on both.
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
        "\u0665",
        "\u00A05",
        "+5",
        ".5",
        "5.",
        "- 5"
      )
    )

  /** Text a date slot reads at its edges: negative, 0, the last serial, one past it, fractions. */
  private val genDateText: Gen[String] = Gen.frequency(
    3 -> genText,
    2 -> Gen.choose(-10L, 3000000L).map(_.toString),
    1 -> Gen.oneOf("-5", "0", "1", "2958465", "2958466", "45000.5", " 45000 ", "-0.5", "1e-400")
  )

  /** Text a logical slot reads: TRUE/FALSE in any case, padded (not trimmed), and the rest. */
  private val genBoolText: Gen[String] = Gen.frequency(
    3 -> genText,
    2 -> Gen.oneOf(
      "TRUE",
      "FALSE",
      "true",
      "False",
      "tRuE",
      " TRUE",
      "FALSE ",
      "T",
      "1",
      "0",
      "yes"
    )
  )

  /** The three shapes a text value takes in a cell: plain, rich, and a formula's cached text. */
  private def genTextCell(gen: Gen[String]): Gen[(String, CellValue)] =
    for
      s <- gen
      shape <- Gen.oneOf(0, 1, 2)
    yield
      val cell = shape match
        case 0 => CellValue.Text(s)
        case 1 => CellValue.RichText(RichText.plain(s))
        case _ => CellValue.Formula("T(\"x\")", Some(CellValue.Text(s)))
      (s, cell)

  private val evaluator = Evaluator.instance

  private def throughReference[A](
    cell: CellValue,
    decode: com.tjclp.xl.cells.Cell => Either[com.tjclp.xl.codec.CodecError, A]
  ): Either[EvalError, A] =
    val sheet = Sheet(SheetName.unsafe("Sheet1")).put(ref"A1", cell)
    evaluator.eval(TExpr.ref(ref"A1", decode), sheet)

  private def isValueError(r: Either[EvalError, Any]): Boolean = r match
    case Left(EvalError.ErrorValue(CellError.Value, _)) => true
    case _ => false

  /** The law: the same value, or `#VALUE!` on both sides. */
  private def agree(byRef: Either[EvalError, Any], byValue: Either[EvalError, Any]): Boolean =
    (byRef, byValue) match
      case (Right(a), Right(b)) => b == a
      case (l, r) => isValueError(l) && isValueError(r)

  private def shown(s: String, cell: CellValue): String =
    s"text ${s.toList.map(_.toInt)} as $cell"

  property("#709: a text cell in a numeric slot reads as the same text as a runtime value") {
    forAll(genTextCell(genText)) { (s, cell) =>
      val byRef = throughReference(cell, TExpr.decodeNumericScalar)
      val byValue = ScalarCoercion.coerce("x", cell, BindingCoercion.Numeric)
      assert(agree(byRef, byValue), s"${shown(s, cell)}: reference $byRef, value $byValue")
    }
  }

  property("#709: a text cell in a date slot reads as the same text as a runtime value") {
    // out-of-range serials included: both paths give the same error value (#726 makes it #NUM!)
    forAll(genTextCell(genDateText)) { (s, cell) =>
      val byRef = throughReference(cell, TExpr.decodeAsDateScalar)
      val byValue = ScalarCoercion.coerce("x", cell, BindingCoercion.Date)
      assert(agree(byRef, byValue), s"${shown(s, cell)}: reference $byRef, value $byValue")
    }
  }

  property("#709: a text cell in a logical slot reads as the same text as a runtime value") {
    forAll(genTextCell(genBoolText)) { (s, cell) =>
      val byRef = throughReference(cell, TExpr.decodeBool)
      val byValue = ScalarCoercion.coerce("x", cell, BindingCoercion.Bool)
      assert(agree(byRef, byValue), s"${shown(s, cell)}: reference $byRef, value $byValue")
    }
  }

  property("#709: the numeric-text pre-check answers exactly as the bare BigDecimal parse") {
    forAll(genText) { s =>
      val bare = scala.util.Try(BigDecimal(s.trim)).toOption
      assertEquals(ScalarCoercion.parseNumericText(s), bare, s"text ${s.toList.map(_.toInt)}")
    }
  }
