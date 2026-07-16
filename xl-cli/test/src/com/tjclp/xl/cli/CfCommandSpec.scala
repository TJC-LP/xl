package com.tjclp.xl.cli

import munit.FunSuite

import java.nio.file.{Files, Path}

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Workbook, Sheet}
import com.tjclp.xl.cf.{CfOperator, CfRule, CfTextOp, Cfvo}
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill

/**
 * E2E tests for the conditional-formatting authoring command (GH-324): `cf add`, `cf list`, and the
 * batch op `cf` — each rule family write → read → typedConditionalFormats assertions, malformed
 * rule strings, and auto-priority stamping.
 */
@SuppressWarnings(
  Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps")
)
class CfCommandSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  val config: WriterConfig = WriterConfig.default

  private def withTempOutput[A](f: Path => A): A =
    val out = Files.createTempFile("cf-test", ".xlsx")
    try f(out)
    finally Files.deleteIfExists(out)

  private def readBackRules(path: Path): Vector[CfRule] =
    ExcelIO
      .instance[IO]
      .read(path)
      .unsafeRunSync()
      .sheets
      .head
      .typedConditionalFormats
      .flatMap(_.rules)

  private def addRule(
    out: Path,
    rule: String,
    bold: Boolean = false,
    italic: Boolean = false,
    underline: Boolean = false,
    strike: Boolean = false,
    bg: Option[String] = None,
    fg: Option[String] = None,
    wb: Workbook = Workbook(Sheet("Test"))
  ): String =
    WriteCommands
      .cfAdd(
        wb,
        Some(wb.sheets.head),
        "A1:A10",
        rule,
        bold,
        italic,
        underline,
        strike,
        bg,
        fg,
        out,
        config
      )
      .unsafeRunSync()

  private def addRuleError(
    rule: String,
    bold: Boolean = false,
    bg: Option[String] = None
  ): String =
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      WriteCommands
        .cfAdd(
          wb,
          Some(wb.sheets.head),
          "A1:A10",
          rule,
          bold,
          false,
          false,
          false,
          bg,
          None,
          out,
          config
        )
        .attempt
        .unsafeRunSync()
        .swap
        .toOption
        .get
        .getMessage
    }

  // ========== rule families E2E (write → read → typed assert) ==========

  test("cf add: cellIs greaterThan with bold + bg round-trips with dxf") {
    withTempOutput { out =>
      val result = addRule(out, "cellIs:greaterThan:100", bold = true, bg = Some("#FFC7CE"))
      assert(result.contains("Added conditional format"), result)
      assert(result.contains("priority 1"), result)

      readBackRules(out) match
        case Vector(CfRule.CellIs(CfOperator.GreaterThan, "100", None, Some(dxf), priority, _)) =>
          assertEquals(priority, 1)
          assertEquals(dxf.fill, Some(Fill.Solid(Color.Rgb(0xffffc7ce))))
          assertEquals(dxf.font.flatMap(_.bold), Some(true))
        case other => fail(s"Expected one CellIs rule, got $other")
    }
  }

  test("cf add: cellIs operator aliases (gt) map to the same operator") {
    withTempOutput { out =>
      addRule(out, "cellIs:gt:50", bold = true)
      readBackRules(out) match
        case Vector(CfRule.CellIs(CfOperator.GreaterThan, "50", None, _, _, _)) => ()
        case other => fail(s"Expected GreaterThan from 'gt', got $other")
    }
  }

  test("cf add: between round-trips both bounds") {
    withTempOutput { out =>
      addRule(out, "between:10:100", bg = Some("yellow"))
      readBackRules(out) match
        case Vector(CfRule.CellIs(CfOperator.Between, "10", Some("100"), Some(_), _, _)) => ()
        case other => fail(s"Expected Between rule, got $other")
    }
  }

  test("cf add: expression keeps colons inside the formula") {
    withTempOutput { out =>
      addRule(out, "expression:SUM(A1:A10)>100", fg = Some("#9C0006"))
      readBackRules(out) match
        case Vector(CfRule.Expression(formula, Some(dxf), _, _)) =>
          assertEquals(formula, "SUM(A1:A10)>100")
          assertEquals(dxf.font.flatMap(_.color), Some(Color.Rgb(0xff9c0006)))
        case other => fail(s"Expected Expression rule, got $other")
    }
  }

  test("cf add: 3-point colorScale round-trips colors with percentile mid") {
    withTempOutput { out =>
      addRule(out, "colorScale:red:white:green")
      readBackRules(out) match
        case Vector(CfRule.ColorScale(min, Some(mid), max, _)) =>
          assertEquals(min.cfvo, Cfvo.Min)
          assertEquals(min.color, Color.Rgb(0xffff0000))
          assertEquals(mid.cfvo, Cfvo.Percentile(BigDecimal(50)))
          assertEquals(mid.color, Color.Rgb(0xffffffff))
          assertEquals(max.cfvo, Cfvo.Max)
          assertEquals(max.color, Color.Rgb(0xff008000))
        case other => fail(s"Expected 3-point ColorScale, got $other")
    }
  }

  test("cf add: 2-point colorScale") {
    withTempOutput { out =>
      addRule(out, "colorScale:#FFFFFF:#4472C4")
      readBackRules(out) match
        case Vector(CfRule.ColorScale(_, None, _, _)) => ()
        case other => fail(s"Expected 2-point ColorScale, got $other")
    }
  }

  test("cf add: dataBar round-trips color") {
    withTempOutput { out =>
      addRule(out, "dataBar:#638EC6")
      readBackRules(out) match
        case Vector(CfRule.DataBar(Cfvo.Min, Cfvo.Max, color, _, _)) =>
          assertEquals(color, Color.Rgb(0xff638ec6))
        case other => fail(s"Expected DataBar, got $other")
    }
  }

  test("cf add: top10 with percent and bottom10") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      WriteCommands
        .cfAdd(
          wb,
          Some(wb.sheets.head),
          "A1:A10",
          "top10:5:percent",
          true,
          false,
          false,
          false,
          None,
          None,
          out,
          config
        )
        .unsafeRunSync()
      val wb2 = ExcelIO.instance[IO].read(out).unsafeRunSync()
      WriteCommands
        .cfAdd(
          wb2,
          Some(wb2.sheets.head),
          "B1:B10",
          "bottom10:3",
          true,
          false,
          false,
          false,
          None,
          None,
          out,
          config
        )
        .unsafeRunSync()

      readBackRules(out) match
        case Vector(
              CfRule.Top10(5, true, false, Some(_), p1, _),
              CfRule.Top10(3, false, true, Some(_), p2, _)
            ) =>
          assertEquals(p1, 1)
          assertEquals(p2, 2)
        case other => fail(s"Expected top10 + bottom10 rules, got $other")
    }
  }

  test("cf add: text contains keeps colons in the needle") {
    withTempOutput { out =>
      addRule(out, "text:contains:status: overdue", bold = true)
      readBackRules(out) match
        case Vector(CfRule.Text(CfTextOp.Contains, needle, Some(_), _, _)) =>
          assertEquals(needle, "status: overdue")
        case other => fail(s"Expected Text contains rule, got $other")
    }
  }

  // ========== auto-priority ==========

  test("cf add: two adds stamp priorities 1 then 2 (auto-priority)") {
    withTempOutput { out =>
      val wb = Workbook(Sheet("Test"))
      addRule(out, "cellIs:greaterThan:100", bold = true, wb = wb)
      val wb2 = ExcelIO.instance[IO].read(out).unsafeRunSync()
      val result2 =
        WriteCommands
          .cfAdd(
            wb2,
            Some(wb2.sheets.head),
            "A1:A10",
            "cellIs:lessThan:0",
            false,
            false,
            false,
            false,
            Some("#FFC7CE"),
            None,
            out,
            config
          )
          .unsafeRunSync()
      assert(result2.contains("priority 2"), result2)

      val priorities = readBackRules(out).flatMap(CfRule.priorityOf)
      assertEquals(priorities, Vector(1, 2))
    }
  }

  // ========== malformed input → clean errors ==========

  test("cf add: unknown rule family yields clean error listing families") {
    val msg = addRuleError("blink:fast", bold = true)
    assert(msg.contains("Unknown rule family"), msg)
    assert(msg.contains("cellIs"), msg)
  }

  test("cf add: unknown cellIs operator yields clean error") {
    val msg = addRuleError("cellIs:sortaBigger:100", bold = true)
    assert(msg.contains("sortaBigger"), msg)
    assert(msg.contains("greaterThan"), msg)
  }

  test("cf add: highlight rule without any format flag yields clean error") {
    val msg = addRuleError("cellIs:greaterThan:100")
    assert(msg.contains("--bold"), msg)
  }

  test("cf add: colorScale with format flags yields clean error") {
    val msg = addRuleError("colorScale:red:green", bold = true)
    assert(msg.contains("inline colors"), msg)
  }

  test("cf add: colorScale with bad color yields clean error") {
    val msg = addRuleError("colorScale:red:notacolor")
    assert(msg.contains("notacolor"), msg)
  }

  test("cf add: top10 with non-numeric rank yields clean error") {
    val msg = addRuleError("top10:many", bold = true)
    assert(msg.contains("positive integer"), msg)
  }

  test("cf add: malformed cellIs (missing value) yields usage error") {
    val msg = addRuleError("cellIs:greaterThan", bold = true)
    assert(msg.contains("cellIs:<operator>:<value>"), msg)
  }

  // ========== cf list ==========

  test("cf list: renders typed rules with priorities") {
    withTempOutput { out =>
      addRule(out, "cellIs:greaterThan:100", bold = true)
      val wb = ExcelIO.instance[IO].read(out).unsafeRunSync()
      val listing = WriteCommands.cfList(wb, Some(wb.sheets.head)).unsafeRunSync()
      assert(listing.contains("A1:A10"), listing)
      assert(listing.contains("cellIs greaterThan 100"), listing)
      assert(listing.contains("priority 1"), listing)
    }
  }

  test("cf list: empty sheet reports no conditional formatting") {
    val wb = Workbook(Sheet("Test"))
    val listing = WriteCommands.cfList(wb, Some(wb.sheets.head)).unsafeRunSync()
    assert(listing.contains("No conditional formatting"), listing)
  }

  // ========== batch op ==========

  test("batch: cf op parses (GH-324)") {
    val json =
      """[{"op":"cf","range":"A1:A10","rule":"cellIs:greaterThan:100","bold":true,"bg":"#FFC7CE"}]"""
    val result = BatchParser.parseBatchJson(json)
    assert(result.isRight, s"Should parse: $result")
    assertEquals(result.toOption.get.warnings, Vector.empty[String])
  }

  test("batch: cf op applies rule with dxf and auto-priority") {
    withTempOutput { out =>
      val json =
        """[
          {"op":"cf","range":"A1:A10","rule":"cellIs:greaterThan:100","bold":true,"bg":"#FFC7CE"},
          {"op":"cf","range":"B1:B10","rule":"dataBar:#638EC6"}
        ]"""
      val jsonPath = Files.createTempFile("cf-batch", ".json")
      Files.writeString(jsonPath, json)
      try
        val wb = Workbook(Sheet("Test"))
        val result =
          WriteCommands
            .batch(wb, Some(wb.sheets.head), jsonPath.toString, out, config)
            .unsafeRunSync()
        assert(result.contains("Applied 2 operations"), result)

        readBackRules(out) match
          case Vector(
                CfRule.CellIs(CfOperator.GreaterThan, "100", None, Some(dxf), 1, _),
                CfRule.DataBar(_, _, _, _, 2)
              ) =>
            assertEquals(dxf.fill, Some(Fill.Solid(Color.Rgb(0xffffc7ce))))
          case other => fail(s"Expected CellIs + DataBar with priorities 1,2, got $other")
      finally Files.deleteIfExists(jsonPath)
    }
  }

  test("batch: cf with malformed rule fails cleanly") {
    val ops = BatchParser
      .parseBatchJson("""[{"op":"cf","range":"A1:A10","rule":"nope:1","bold":true}]""")
      .toOption
      .get
      .ops
    val sheet = Sheet("Test")
    val wb = Workbook(sheet)
    val result = BatchParser.applyBatchOperations(wb, Some(sheet), ops).attempt.unsafeRunSync()
    assert(result.isLeft)
    assert(
      result.swap.toOption.get.getMessage.contains("Unknown rule family"),
      result.swap.toOption.get.getMessage
    )
  }
