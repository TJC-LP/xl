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
    val label = f.labelWidth
    val body = Markdown
      .lines(window, rows(t), cols, widths, label, t.showFormulas, t.skipEmpty, t.skipHidden)
    val after = "" +: t.warnings.map(_.message)
    val streamed =
      Payload.Streamed(StreamedBody.Lines(Markdown.header(cols, widths, label), body, after))
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

  /** Every byte [[Render.stream]] produces, joined: what `emit` writes for the outcome. */
  private def streamedBytes(payload: Payload, warnings: Vector[Warning], mode: OutputMode): String =
    Render
      .stream(mode)(Outcome.ok("view", payload, warnings), version)
      .compile
      .toVector
      .unsafeRunSync()
      .mkString

  /** What `emit` prints for a gathered payload: its stdout and the newline, or nothing at all. */
  private def wholeStdout(payload: Payload, warnings: Vector[Warning], mode: OutputMode): String =
    val stdout = Render(mode)(Outcome.ok("view", payload, warnings), version).stdout
    if stdout.isEmpty then "" else stdout + "\n"

  // --- The law -----------------------------------------------------------------------------------

  property("markdown: the streamed bytes are the whole table, in both modes, stderr agreeing") {
    forAll(genTable) { (t: Table) =>
      val (streamed, text) = markdown(t)
      val gathered = Payload.materialise(streamed).unsafeRunSync()
      assertEquals(gathered, Payload.text(text))
      Vector(OutputMode.Text, OutputMode.Json).foreach { mode =>
        assertEquals(
          streamedBytes(streamed, t.warnings, mode),
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

  property("csv: the streamed bytes are the whole table, in both modes (data.text streams)") {
    forAll(genTable) { (t: Table) =>
      val (streamed, text) = csv(t)
      val gathered = Payload.materialise(streamed).unsafeRunSync()
      assertEquals(gathered, Payload.text(text))
      Vector(OutputMode.Text, OutputMode.Json).foreach { mode =>
        val bytes = streamedBytes(streamed, t.warnings, mode)
        assertEquals(bytes, wholeStdout(gathered, t.warnings, mode), s"$mode")
        // data.text, escaped line by line as it streams, reads back as the table
        if mode == OutputMode.Json then
          assertEquals(ujson.read(bytes)("data")("text"), ujson.Str(text))
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
        val lines = streamedBytes(streamed, t.warnings, mode)
        assertEquals(lines, wholeStdout(gathered, t.warnings, mode), s"$mode")
        if mode == OutputMode.Json then EnvelopeSchema.assertValid(ujson.read(lines))
      }
      Prop.passed
    }
  }

  test("a gathered payload is one fragment with its newline; an empty stdout is no bytes at all") {
    val text = Outcome.ok("put", Payload.text("Put: A1 = 1\nSaved: out.xlsx"))
    assertEquals(
      Render.stream(OutputMode.Text)(text, version).compile.toVector.unsafeRunSync(),
      Vector("Put: A1 = 1\nSaved: out.xlsx\n")
    )
    // no lines: nothing, as 0.21.0 printed nothing for the empty text
    val empty = Payload.Streamed(StreamedBody.Lines(Vector.empty, Stream.empty, Vector.empty))
    assertEquals(
      Render
        .stream(OutputMode.Text)(Outcome.ok("view", empty), version)
        .compile
        .string
        .unsafeRunSync(),
      ""
    )
    assertEquals(Payload.materialise(empty).unsafeRunSync(), Payload.text(""))
    // one empty line is the empty text too; two empty lines are a newline
    val oneEmpty = Payload.Streamed(StreamedBody.Lines(Vector.empty, Stream.emit(""), Vector.empty))
    assertEquals(
      Render
        .stream(OutputMode.Text)(Outcome.ok("view", oneEmpty), version)
        .compile
        .string
        .unsafeRunSync(),
      ""
    )
    val twoEmpty =
      Payload.Streamed(StreamedBody.Lines(Vector.empty, Stream.emits(Vector("", "")), Vector.empty))
    assertEquals(
      Render
        .stream(OutputMode.Text)(Outcome.ok("view", twoEmpty), version)
        .compile
        .string
        .unsafeRunSync(),
      "\n\n"
    )
    // under --json the empty table is still an envelope with data.text ""
    val envelope = Render
      .stream(OutputMode.Json)(Outcome.ok("view", empty), version)
      .compile
      .string
      .unsafeRunSync()
    assertEquals(ujson.read(envelope)("data")("text"), ujson.Str(""))
    val nothing = Outcome.failed("view", CliError("SHEET_NOT_FOUND", "Sheet not found: X"))
    assertEquals(
      Render.stream(OutputMode.Text)(nothing, version).compile.toVector.unsafeRunSync(),
      Vector.empty
    )
  }

  test("nothing is produced before the first row: a source that fails at once yields no fragment") {
    val failing = Stream.raiseError[IO](new IllegalStateException("no table"))
    val csvBody = Payload.Streamed(StreamedBody.Lines(Vector(",A,B"), failing, Vector.empty))
    val jsonBody = Payload.Streamed(StreamedBody.JsonArray("{\n  \"rows\": [", failing, "]\n}"))
    Vector(csvBody, jsonBody).foreach { body =>
      Vector(OutputMode.Text, OutputMode.Json).foreach { mode =>
        val produced = Render
          .stream(mode)(Outcome.ok("view", body), version)
          .attempt
          .compile
          .toVector
          .unsafeRunSync()
        assertEquals(produced.collect { case Right(fragment) => fragment }, Vector.empty, s"$mode")
        assert(produced.exists(_.isLeft), s"the failure surfaces in $mode")
      }
    }
  }

  test("a 600-row table streams as the whole, in both modes (past one write of 256 fragments)") {
    val rows = (1 to 600).toVector.map(i => s"row $i,\"q\"\"uote\",$i.5")
    val body = Payload.Streamed(
      StreamedBody.Lines(Vector(",A,B,C"), Stream.emits(rows).covary[IO], Vector.empty)
    )
    val gathered = Payload.text((",A,B,C" +: rows).mkString("\n"))
    Vector(OutputMode.Text, OutputMode.Json).foreach { mode =>
      assertEquals(
        streamedBytes(body, Vector.empty, mode),
        wholeStdout(gathered, Vector.empty, mode)
      )
    }
  }
