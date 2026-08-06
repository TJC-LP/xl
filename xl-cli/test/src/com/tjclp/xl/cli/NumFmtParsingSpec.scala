package com.tjclp.xl.cli

import java.io.{ByteArrayOutputStream, PrintStream}
import java.nio.charset.StandardCharsets
import java.nio.file.Files

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.cli.helpers.{BatchParser, StyleBuilder}
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig
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

  // ========== the direct CLI `style --format` surface (NOT batch) ==========
  //
  // This is where a human actually fat-fingers a semantic name; the batch-JSON warning does not
  // cover it.

  private def captureStderr[A](io: IO[A]): IO[(A, String)] =
    IO.blocking {
      val baos = new ByteArrayOutputStream()
      val prevErr = System.err
      System.setErr(new PrintStream(baos, true, "UTF-8"))
      try
        import cats.effect.unsafe.implicits.global
        val result = io.unsafeRunSync()
        (result, new String(baos.toByteArray, StandardCharsets.UTF_8))
      finally System.setErr(prevErr)
    }

  /** Run the `style` command with a `--format` string; return (stdout message, stderr). */
  private def runStyleCommand(numFormat: String): IO[(String, String)] =
    IO.blocking {
      val out = Files.createTempFile("xl-numfmt-style-", ".xlsx")
      out.toFile.deleteOnExit()
      out
    }.flatMap { out =>
      val sheet = Sheet("S").put(ref"A1", CellValue.Number(BigDecimal(1)))
      val book = Workbook(Vector(sheet))
      captureStderr(
        WriteCommands.style(
          book,
          book.sheets.headOption,
          "A1",
          bold = false,
          italic = false,
          underline = false,
          bg = None,
          fg = None,
          fontSize = None,
          fontName = None,
          align = None,
          valign = None,
          wrap = false,
          numFormat = Some(numFormat),
          border = None,
          borderTop = None,
          borderRight = None,
          borderBottom = None,
          borderLeft = None,
          borderColor = None,
          replace = false,
          outputPath = out,
          config = WriterConfig.default
        )
      )
    }

  test("GH-475: `style --format curency` warns on stderr instead of shipping the typo silently") {
    runStyleCommand("curency").map { case (stdout, stderr) =>
      assert(stdout.contains("Styled: A1"), s"style command must still succeed: $stdout")
      assert(
        stderr.contains("curency"),
        s"expected a stderr warning naming the typo, got: '$stderr'"
      )
      assert(
        stderr.contains("currency"),
        s"warning should list the known names, got: '$stderr'"
      )
    }
  }

  test("GH-475: `style --format` with a known name or a real code stays silent") {
    List("currency", "percent", "#,##0.00", "\"Yes \";;\"No \"").foldLeft(IO.unit) { (acc, fmt) =>
      acc *> runStyleCommand(fmt).map { case (_, stderr) =>
        assertEquals(stderr, "", s"'$fmt' must not warn")
      }
    }
  }
