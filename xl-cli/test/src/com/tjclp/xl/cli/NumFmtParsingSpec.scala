package com.tjclp.xl.cli

import munit.CatsEffectSuite

import com.tjclp.xl.cli.helpers.{BatchParser, StyleBuilder}
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * GH-475 (second half): numFmt strings must never be silently mishandled.
 *
 *   - a code made only of quoted literals and semicolons (`"Yes ";;"No "`, the 1/0 toggle-flag
 *     idiom) is a real Excel format code — the put/putf `format` field dropped it on the floor
 *   - a string that is neither a known name nor code-shaped (`curency`) still ships as a custom
 *     code (Excel is the authority on codes, not us), but it now warns so a typo is visible
 */
class NumFmtParsingSpec extends CatsEffectSuite:

  // ========== looksLikeFormatCode ==========

  test("GH-475: quoted-literal/semicolon codes are recognized as format codes") {
    assert(StyleBuilder.looksLikeFormatCode("\"Yes \";;\"No \""))
    assert(StyleBuilder.looksLikeFormatCode("_(* #,##0_);_(* (#,##0);_(* \"-\"_)"))
    assert(StyleBuilder.looksLikeFormatCode("0.0x"))
    assert(StyleBuilder.looksLikeFormatCode("yyyy-mm-dd"))
    assert(StyleBuilder.looksLikeFormatCode("hh:mm:ss"))
    assert(StyleBuilder.looksLikeFormatCode("#,##0"))
    assert(StyleBuilder.looksLikeFormatCode("@"))
  }

  test("GH-475: fat-fingered semantic names are NOT format-code shaped") {
    assert(!StyleBuilder.looksLikeFormatCode("curency"))
    assert(!StyleBuilder.looksLikeFormatCode("nonsense-xyz"))
    assert(!StyleBuilder.looksLikeFormatCode("percent"))
  }

  // ========== numFmtWarning ==========

  test("GH-475: a suspect numFmt string warns and names the string") {
    val warning = StyleBuilder.numFmtWarning("curency")
    assert(warning.isDefined, "expected a warning for 'curency'")
    assert(warning.exists(_.contains("curency")), s"warning must name the string: $warning")
    assert(warning.exists(_.contains("currency")), s"warning should suggest names: $warning")
  }

  test("GH-475: known names and real codes never warn") {
    List("currency", "percent", "General", "datetime", "#,##0", "\"Yes \";;\"No \"").foreach { s =>
      assertEquals(StyleBuilder.numFmtWarning(s), None, s"'$s' must not warn")
    }
  }

  test("GH-475: a suspect string still parses to a custom code (Excel is the authority)") {
    assertEquals(StyleBuilder.parseNumFmt("curency"), Right(NumFmt.Custom("curency")))
  }

  // ========== batch: the put/putf format field ==========

  test("GH-475: put 'format' keeps a quoted-literal/semicolon code") {
    val json = """[{"op":"put","ref":"A1","value":1,"format":"\"Yes \";;\"No \""}]"""
    BatchParser.parseBatchOperations(json).map { result =>
      result.ops.headOption match
        case Some(BatchParser.BatchOp.Put(_, _, format)) =>
          assertEquals(format, Some(NumFmt.Custom("\"Yes \";;\"No \"")))
        case other => fail(s"expected a Put op, got $other")
    }
  }

  test("GH-475: put 'format' with a suspect string warns instead of silently dropping") {
    val json = """[{"op":"put","ref":"A1","value":1,"format":"curency"}]"""
    BatchParser.parseBatchOperations(json).map { result =>
      assert(
        result.warnings.exists(_.contains("curency")),
        s"expected a numFmt warning, got ${result.warnings}"
      )
    }
  }

  test("GH-475: style op with a suspect numFormat warns") {
    val json = """[{"op":"style","range":"A1","numFormat":"curency"}]"""
    BatchParser.parseBatchOperations(json).map { result =>
      assert(
        result.warnings.exists(_.contains("curency")),
        s"expected a numFmt warning, got ${result.warnings}"
      )
    }
  }

  test("GH-475: style op with a real code stays silent") {
    val json = """[{"op":"style","range":"A1","numFormat":"#,##0.0"}]"""
    BatchParser.parseBatchOperations(json).map { result =>
      assertEquals(result.warnings.filter(_.contains("numFmt")), Vector.empty)
    }
  }
