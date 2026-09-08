package com.tjclp.xl.cli.contract

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import fs2.Stream
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.output.{CsvRenderer, JsonRenderer, Markdown}
import com.tjclp.xl.cli.read.{CellRecord, ColumnFacts, RecordGrid, RecordWindow}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * GH-635's law: a table written as it streams is byte for byte the table written whole. For every
 * generated window — text, numbers of many digits, booleans, cached and uncached formulas, empties,
 * hidden lines, `--skip-empty`, `--skip-hidden`, `--header-row`, truncation fields, warnings on the
 * outcome — and both output modes, the lines [[Render.lines]] produces for a [[Payload.Streamed]],
 * joined, equal the stdout [[Render]] gives the same rows gathered into a [[Payload.Text]] or
 * [[Payload.Raw]]; the stderr agrees too. Under `--json` that is the envelope with the document
 * spliced element by element in `ujson`'s layout, which is the part that has to be reproduced
 * piecewise. The renderers' `render(grid, …)` (what every golden pins) are the same composition, so
 * they are held to the streamed lines as well.
 */
class StreamedPayloadSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(60)

  private val version = "9.9.9"
  private val sheet = SheetName.unsafe("Data")

  // --- Generators --------------------------------------------------------------------------------

  private val genText: Gen[String] =
    Gen.frequency(
      5 -> Gen.identifier.map(_.take(10)),
      1 -> Gen.const("a,b \"c\" | d"),
      1 -> Gen.const("two\nlines\ttab"),
      1 -> Gen.const("ünïcødé — ✓"),
      1 -> Gen.const("\\backslash and  control"),
      1 -> Gen.const("")
    )

  private val genValue: Gen[CellValue] =
    Gen.frequency(
      4 -> genText.map(CellValue.Text(_)),
      3 -> Gen.choose(-100000, 100000).map(n => CellValue.Number(BigDecimal(n) / 100)),
      1 -> Gen.const(CellValue.Number(BigDecimal("12345678901234567"))),
      1 -> Gen.oneOf(true, false).map(CellValue.Bool(_)),
      1 -> Gen.const(CellValue.Formula("SUM(A1:A2)", Some(CellValue.Number(BigDecimal("12.5"))))),
      1 -> Gen.const(CellValue.Formula("B2*2", None)),
      3 -> Gen.const(CellValue.Empty)
    )

  private val genStyle: Gen[Option[CellStyle]] =
    Gen.frequency(
      4 -> Gen.const(None),
      1 -> Gen.const(Some(CellStyle.default.withNumFmt(NumFmt.Percent))),
      1 -> Gen.const(Some(CellStyle.default.withNumFmt(NumFmt.Currency)))
    )

  /** A window's worth of dense rows plus the flags a `view` can carry. */
  private final case class Table(
    grid: RecordGrid,
    skipEmpty: Boolean,
    skipHidden: Boolean,
    showFormulas: Boolean,
    showLabels: Boolean,
    headerRow: Option[Int],
    truncatedRows: Option[Int],
    truncatedCols: Option[Int],
    warnings: Vector[Warning]
  )

  private val genTable: Gen[Table] =
    for
      rows <- Gen.choose(0, 6)
      cols <- Gen.choose(1, 4)
      startRow <- Gen.choose(0, 2)
      startCol <- Gen.choose(0, 1)
      cells <- Gen.listOfN(rows * cols, Gen.zip(genValue, genStyle))
      hiddenRows <- Gen.someOf(startRow until startRow + rows)
      hiddenCols <- Gen.someOf(startCol until startCol + cols)
      skipEmpty <- Gen.oneOf(true, false)
      skipHidden <- Gen.oneOf(true, false)
      showFormulas <- Gen.oneOf(true, false)
      showLabels <- Gen.oneOf(true, false)
      header <- Gen.option(Gen.choose(startRow, startRow + math.max(rows, 1)))
      truncatedRows <- Gen.option(Gen.choose(rows + 1, 1000))
      truncatedCols <- Gen.option(Gen.choose(cols + 1, 50))
      warnings <- Gen.listOf(
        Gen.oneOf(
          Warning(WarningCode.TRUNCATED, "… showing 2 of 4 rows"),
          Warning(WarningCode.FLAG_IGNORED, "note: --skip-hidden is \"ignored\"", None)
        )
      )
    yield
      val range = CellRange(
        ARef.from0(startCol, startRow),
        ARef.from0(startCol + cols - 1, startRow + math.max(rows, 1) - 1)
      )
      val dense = (0 until rows).toVector.map { r =>
        (0 until cols).toVector.map { c =>
          val (value, style) = cells(r * cols + c)
          CellRecord.of(
            sheet,
            ARef.from0(startCol + c, startRow + r),
            value,
            style,
            hidden = Some(hiddenRows.contains(startRow + r) || hiddenCols.contains(startCol + c)),
            mergedInto = None
          )
        }
      }
      // A window is never empty: with no rows the range still addresses one row of empties
      val padded =
        if dense.nonEmpty then dense
        else
          Vector(
            (0 until cols).toVector.map(c =>
              CellRecord.empty(sheet, ARef.from0(startCol + c, startRow), Some(false), None)
            )
          )
      Table(
        RecordGrid(sheet, range, padded, hiddenRows.toSet, hiddenCols.toSet),
        skipEmpty,
        skipHidden,
        showFormulas,
        showLabels,
        header.map(_ + 1),
        truncatedRows,
        truncatedCols,
        warnings.toVector
      )

  // --- The bodies a `view` builds, and the text they compose -----------------------------------

  private def rows(t: Table): Stream[IO, Vector[CellRecord]] = Stream.emits(t.grid.rows)

  private def facts(t: Table): ColumnFacts =
    ColumnFacts
      .of(t.grid.window, rows(t), Markdown.cellText(t.showFormulas), t.skipEmpty, t.skipHidden)
      .compile
      .lastOrError
      .unsafeRunSync()

  private def markdown(t: Table): (Payload, String) =
    val window = t.grid.window
    val f = facts(t)
    val cols = Markdown.columns(window, f, t.skipEmpty, t.skipHidden)
    val widths = Markdown.columnWidths(cols, f)
    val body =
      Markdown.lines(window, rows(t), cols, widths, t.showFormulas, t.skipEmpty, t.skipHidden)
    val after = "" +: t.warnings.map(_.message)
    val streamed = Payload.Streamed(StreamedBody.Lines(Markdown.header(cols, widths), body, after))
    val whole = Markdown.render(t.grid, t.showFormulas, t.skipEmpty, t.skipHidden)
    (streamed, (whole +: t.warnings.map(_.message)).mkString("\n"))

  private def csv(t: Table): (Payload, String) =
    val window = t.grid.window
    val f = if CsvRenderer.needsFacts(t.skipEmpty) then facts(t) else ColumnFacts.empty
    val cols = CsvRenderer.columns(window, f, t.skipEmpty, t.skipHidden)
    val body = CsvRenderer
      .lines(window, rows(t), cols, t.showFormulas, t.showLabels, t.skipEmpty, t.skipHidden)
    val streamed =
      Payload.Streamed(
        StreamedBody.Lines(CsvRenderer.header(cols, t.showLabels), body, Vector.empty)
      )
    (streamed, CsvRenderer.render(t.grid, t.showFormulas, t.showLabels, t.skipEmpty, t.skipHidden))

  private def json(t: Table): (Payload, String) =
    val window = t.grid.window
    val header = t.headerRow.map { rowNum =>
      val idx = rowNum - 1
      (idx, t.grid.rows.lift(idx - window.range.start.row.index0).getOrElse(Vector.empty))
    }
    val streamed = Payload.Streamed(
      StreamedBody.JsonArray(
        JsonRenderer.head(window, header.isDefined, t.truncatedRows, t.truncatedCols),
        JsonRenderer.elements(window, rows(t), header, t.skipEmpty, t.skipHidden),
        JsonRenderer.tail
      )
    )
    val whole = JsonRenderer
      .render(t.grid, header, t.skipEmpty, t.truncatedRows, t.truncatedCols, t.skipHidden)
    (streamed, whole)

  private def streamedLines(payload: Payload, warnings: Vector[Warning], mode: OutputMode): String =
    Render
      .lines(mode)(Outcome.ok("view", payload, warnings), version)
      .compile
      .toVector
      .unsafeRunSync()
      .mkString("\n")

  private def wholeStdout(payload: Payload, warnings: Vector[Warning], mode: OutputMode): String =
    Render(mode)(Outcome.ok("view", payload, warnings), version).stdout

  // --- The law -----------------------------------------------------------------------------------

  property("markdown: the streamed lines are the whole table, in both modes, stderr agreeing") {
    forAll(genTable) { (t: Table) =>
      val (streamed, text) = markdown(t)
      val gathered = Payload.materialise(streamed).unsafeRunSync()
      assertEquals(gathered, Payload.text(text))
      Vector(OutputMode.Text, OutputMode.Json).foreach { mode =>
        assertEquals(
          streamedLines(streamed, t.warnings, mode),
          wholeStdout(gathered, t.warnings, mode)
        )
        assertEquals(
          Render.stderr(mode)(Outcome.ok("view", streamed, t.warnings)),
          Render(mode)(Outcome.ok("view", gathered, t.warnings), version).stderr
        )
      }
      Prop.passed
    }
  }

  property("csv: the streamed lines are the whole table, in both modes") {
    forAll(genTable) { (t: Table) =>
      val (streamed, text) = csv(t)
      val gathered = Payload.materialise(streamed).unsafeRunSync()
      assertEquals(gathered, Payload.text(text))
      Vector(OutputMode.Text, OutputMode.Json).foreach { mode =>
        assertEquals(
          streamedLines(streamed, t.warnings, mode),
          wholeStdout(gathered, t.warnings, mode)
        )
      }
      Prop.passed
    }
  }

  property(
    "json: the streamed elements are the whole document; under --json the spliced envelope"
  ) {
    forAll(genTable) { (t: Table) =>
      val (streamed, text) = json(t)
      val gathered = Payload.materialise(streamed).unsafeRunSync()
      assertEquals(gathered, Payload.Raw(text))
      // the document is valid JSON whatever the rows held
      assert(ujson.read(text).obj.contains("sheet"), text)
      Vector(OutputMode.Text, OutputMode.Json).foreach { mode =>
        val lines = streamedLines(streamed, t.warnings, mode)
        assertEquals(lines, wholeStdout(gathered, t.warnings, mode), s"$mode")
        if mode == OutputMode.Json then EnvelopeSchema.assertValid(ujson.read(lines))
      }
      Prop.passed
    }
  }

  test("a materialised payload is one element of lines, and an empty stdout is one empty line") {
    val text = Outcome.ok("put", Payload.text("Put: A1 = 1\nSaved: out.xlsx"))
    assertEquals(
      Render.lines(OutputMode.Text)(text, version).compile.toVector.unsafeRunSync(),
      Vector("Put: A1 = 1\nSaved: out.xlsx")
    )
    val empty = Payload.Streamed(StreamedBody.Lines(Vector.empty, Stream.empty, Vector.empty))
    assertEquals(
      Render
        .lines(OutputMode.Text)(Outcome.ok("view", empty), version)
        .compile
        .toVector
        .unsafeRunSync(),
      Vector("")
    )
    assertEquals(Payload.materialise(empty).unsafeRunSync(), Payload.text(""))
    val nothing = Outcome.failed("view", CliError("SHEET_NOT_FOUND", "Sheet not found: X"))
    assertEquals(
      Render.lines(OutputMode.Text)(nothing, version).compile.toVector.unsafeRunSync(),
      Vector("")
    )
  }

  test("a streamed table past the text budget is RESOURCE_LIMIT, never a heap that ran out") {
    val rows = Stream.emits(Vector.fill(100)("x" * 100)).covary[IO]
    val body = Payload.Streamed(StreamedBody.Lines(Vector.empty, rows, Vector.empty))
    val refused = Payload.materialise(body, budget = 1000L).attempt.unsafeRunSync()
    refused match
      case Left(e: CliException) =>
        assertEquals(e.error.code, ErrorCode.RESOURCE_LIMIT)
        assert(e.error.message.contains("cannot be held as one string"), e.error.message)
        assert(e.error.hint.exists(_.contains("--format json")), e.error.hint.toString)
      case other => fail(s"expected RESOURCE_LIMIT, got $other")
    val fits = Payload.materialise(body, budget = 100L * 101L + 1L).unsafeRunSync()
    assertEquals(fits, Payload.text(Vector.fill(100)("x" * 100).mkString("\n")))
    // the JSON document streams inside the envelope and is never budgeted
    val doc = Payload.Streamed(
      StreamedBody.JsonArray(
        "{\n  \"rows\": [",
        rows.map(r => s"""    {"row": 1, "cells": ["$r"]}"""),
        "]\n}"
      )
    )
    val lines = Render
      .lines(OutputMode.Json)(Outcome.ok("view", doc), version)
      .compile
      .toVector
      .unsafeRunSync()
    assert(lines.size > 100, lines.size.toString)
  }
