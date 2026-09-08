package com.tjclp.xl.cli.read

import java.nio.file.{Files, Path}

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import cats.syntax.all.*
import munit.FunSuite
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue, Comment}
import com.tjclp.xl.cli.{FilterFormat, ViewFormat}
import com.tjclp.xl.cli.contract.{Outcome, OutputMode, Render, Rendered}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * W2.4's acceptance law: for every read verb, the loaded workbook and the streaming reader produce
 * equal payloads when the book uses only the capabilities both sources share
 * ([[Capability.streaming]]: values, styles, formulas, comments). The generated books exercise
 * every one of those: text, numbers, booleans, errors and empties under a few number formats,
 * cached and uncached formulas, comments (authored and not, on occupied and on empty cells) and
 * styled-but-empty cells — no hidden lines, merges or hyperlinks — written to a file and read back
 * through both strategies; every query runs in both output modes and the rendered stdout, stderr
 * and exit code must agree. The file is written as other producers write it: openpyxl's
 * package-absolute worksheet Targets on half the books, and Excel's `<dimension ref="A1"/>` on
 * every empty sheet (the library's writer records none there).
 *
 * The payload fields that BELONG to a capability the streaming reader lacks are the only allowed
 * difference, and they must say so rather than guess: `cell`'s `Dependencies`/`Dependents` (the
 * `graph` capability) print `(not available in streaming mode)` / `null`, and the typed record's
 * `hidden` (the `hidden` capability) is `null` where the loaded workbook says `false`. The law
 * projects exactly those fields out and compares the rest byte for byte.
 */
class SourceParitySpec extends FunSuite with ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(12).withMaxDiscardRatio(10)

  // --- Generators --------------------------------------------------------------------------------

  private val genText: Gen[String] =
    Gen.frequency(
      6 -> Gen.identifier.map(_.take(8)),
      1 -> Gen.const("a,b \"c\" | d"),
      1 -> Gen.const("Total"),
      1 -> Gen.const("x y"),
      1 -> Gen.const("")
    )

  private val genNumber: Gen[BigDecimal] =
    Gen.frequency(
      5 -> Gen.choose(-1000, 1000).map(BigDecimal(_)),
      3 -> Gen.choose(-100000, 100000).map(n => BigDecimal(n) / 100),
      1 -> Gen.const(BigDecimal("12345678901234567")),
      1 -> Gen.const(BigDecimal("0.125"))
    )

  /** Formula cells as a file carries them: with a cached number, text or boolean, or uncached. */
  private val genFormula: Gen[CellValue] =
    Gen.frequency(
      3 -> Gen.const(CellValue.Formula("SUM(A1:A2)", Some(CellValue.Number(BigDecimal("12.5"))))),
      2 -> Gen.const(CellValue.Formula("A1&\"x\"", Some(CellValue.Text("Totalx")))),
      1 -> Gen.const(CellValue.Formula("A1>0", Some(CellValue.Bool(true)))),
      2 -> Gen.const(CellValue.Formula("B2*2", None))
    )

  private val genValue: Gen[Option[CellValue]] =
    Gen.frequency(
      4 -> genText.map(t => Some(CellValue.Text(t))),
      4 -> genNumber.map(n => Some(CellValue.Number(n))),
      1 -> Gen.oneOf(true, false).map(b => Some(CellValue.Bool(b))),
      1 -> Gen
        .oneOf(CellError.Div0, CellError.NA, CellError.Ref)
        .map(e => Some(CellValue.Error(e))),
      2 -> genFormula.map(Some(_)),
      3 -> Gen.const(None)
    )

  private val genNumFmt: Gen[Option[NumFmt]] =
    Gen.frequency(
      5 -> Gen.const(None),
      1 -> Gen.const(Some(NumFmt.Percent)),
      1 -> Gen.const(Some(NumFmt.Decimal)),
      1 -> Gen.const(Some(NumFmt.Currency)),
      1 -> Gen.const(Some(NumFmt.Custom("0.0")))
    )

  private val genComment: Gen[Comment] =
    for
      text <- Gen.oneOf("input", "check this", "two\nlines")
      author <- Gen.option(Gen.oneOf("qa", "Ana Lee"))
    yield Comment.plainText(text, author)

  /**
   * A sheet: up to 6 rows by 4 columns, a header row of words first. A position with no value but a
   * number format becomes a styled-but-empty cell, and half the sheets carry one more such cell
   * just past the grid's last row and column — outside the non-empty bounding box, so the
   * stored-cell box (the `<dimension>`) is wider than any scan of the values could recover; up to
   * two comments land anywhere in the grid or on the row below it (an empty cell).
   */
  private def genSheet(name: String): Gen[Sheet] =
    for
      rows <- Gen.choose(1, 6)
      cols <- Gen.choose(1, 4)
      headers <- Gen.listOfN(cols, Gen.identifier.map(_.take(6)).suchThat(_.nonEmpty))
      cells <- Gen.listOfN(rows * cols, Gen.zip(genValue, genNumFmt))
      styledCorner <- Gen.oneOf(true, false)
      commentCount <- Gen.choose(0, 2)
      comments <- Gen.listOfN(
        commentCount,
        Gen.zip(Gen.choose(0, cols - 1), Gen.choose(0, rows + 1), genComment)
      )
    yield
      val blank: Sheet = Sheet(SheetName.unsafe(name))
      val withHeaders = headers.zipWithIndex.foldLeft(blank) { case (s, (h, col)) =>
        s.put(Cell(ARef.from0(col, 0), CellValue.Text(h)))
      }
      val withCells = cells.zipWithIndex.foldLeft(withHeaders) { case (s, ((value, numFmt), i)) =>
        val ref = ARef.from0(i % cols, 1 + i / cols)
        val put: Sheet = value.fold(s)(v => s.put(Cell(ref, v)))
        numFmt.fold(put)(fmt =>
          put.styleAt(ref.toA1, CellStyle.default.withNumFmt(fmt)).getOrElse(put)
        )
      }
      val withCorner =
        if styledCorner then
          val corner = ARef.from0(cols, rows + 1).toA1
          withCells
            .styleAt(corner, CellStyle.default.withNumFmt(NumFmt.Decimal))
            .getOrElse(withCells)
        else withCells
      comments.foldLeft(withCorner) { case (s, (col, row, comment)) =>
        s.comment(ARef.from0(col, row), comment)
      }

  /**
   * A book as another producer would have written it: with `openpyxlTargets` the workbook rels name
   * the worksheets by package-absolute Target (`/xl/worksheets/sheet1.xml`); the `Blank` sheet,
   * when present, is an empty sheet that will carry Excel's `<dimension ref="A1"/>`.
   */
  private final case class ParityBook(wb: Workbook, openpyxlTargets: Boolean)

  private val genBook: Gen[ParityBook] =
    for
      two <- Gen.oneOf(true, false)
      blank <- Gen.oneOf(true, false)
      openpyxl <- Gen.oneOf(true, false)
      data <- genSheet("Data")
      other <- genSheet("Other")
    yield
      val sheets = Vector(data) ++ Option.when(blank)(Sheet("Blank")) ++ Option.when(two)(other)
      ParityBook(Workbook(sheets), openpyxl)

  // --- The law -----------------------------------------------------------------------------------

  private val excel = ReadTestKit.excel

  private def withFile[A](book: ParityBook)(f: (Workbook, Path) => IO[A]): IO[A] =
    IO.blocking(Files.createTempFile("xl-parity-", ".xlsx")).flatMap { path =>
      val asProduced: (String, Array[Byte]) => Array[Byte] = (name, bytes) =>
        val excelShaped = ReadTestKit.excelEmptySheetDimension(name, bytes)
        if book.openpyxlTargets then ReadTestKit.openpyxlTargets(name, excelShaped) else excelShaped
      (excel.write(book.wb, path) *> ReadTestKit.rewriteZip(path)(asProduced) *>
        excel.read(path).flatMap(loaded => f(loaded, path)))
        .guarantee(IO.blocking(Files.deleteIfExists(path)).void)
    }

  private def queries(wb: Workbook, sheet: Sheet): Vector[(Option[String], ReadQuery)] =
    val name = Some(sheet.name.value)
    val someText = sheet.cells.values.map(_.value).collectFirst { case CellValue.Text(s) => s }
    val refs = sheet.cells.keys.toVector.sortBy(r => (r.row.index0, r.col.index0)).take(4)
    val commented = sheet.comments.keys.toVector.sortBy(r => (r.row.index0, r.col.index0))
    val views = Vector(
      ReadTestKit.view(None),
      ReadTestKit.view(None, ViewFormat.Csv, showLabels = true),
      ReadTestKit.view(None, ViewFormat.Json),
      ReadTestKit.view(None, ViewFormat.Json, skipEmpty = true),
      ReadTestKit.view(None, skipEmpty = true),
      ReadTestKit.view(None, ViewFormat.Json, headerRow = Some(1)),
      ReadTestKit.view(Some("A1:E8"), ViewFormat.Json, limit = 3),
      ReadTestKit.view(Some("A1:E8"), ViewFormat.Csv, limit = 3),
      ReadTestKit.view(None, ViewFormat.Json, offset = 1, maxCols = 2),
      ReadTestKit.view(None, offset = 1, maxCols = 2),
      ReadTestKit.view(Some("B2"), ViewFormat.Json),
      ReadTestKit.view(None, showFormulas = true),
      ReadTestKit.view(None, ViewFormat.Csv, showFormulas = true)
    ).map(q => (name, q: ReadQuery))
    val searches = Vector(
      (None, ReadQuery.Search("\\d", 50, None)),
      (None, ReadQuery.Search("\\d", 2, None)),
      (name, ReadQuery.Search("TRUE|FALSE", 0, None)),
      (None, ReadQuery.Search("^$", 50, None)),
      (None, ReadQuery.Search("SUM|\\*", 50, None)),
      (None, ReadQuery.Search(java.util.regex.Pattern.quote(someText.getOrElse("Total")), 50, None))
    )
    val stats = Vector((name, ReadQuery.Stats("A1:E8")), (name, ReadQuery.Stats("B2:B3")))
    val filters = Vector(
      (name, ReadTestKit.filter("B > 0")),
      (name, ReadTestKit.filter("A IS NOT EMPTY", format = FilterFormat.Csv, header = true)),
      (name, ReadTestKit.filter("B >= 0 OR C = TRUE", format = FilterFormat.Json, limit = 2)),
      (name, ReadTestKit.filter("A IS EMPTY", columns = Some("A:B"))),
      // A --columns token outside the used range: blank cells, the row number kept, from both
      (name, ReadTestKit.filter("B > 0", columns = Some("Z"), format = FilterFormat.Json)),
      (
        name,
        ReadTestKit.filter("A IS NOT EMPTY", columns = Some("Z,A"), format = FilterFormat.Csv)
      ),
      (name, ReadTestKit.filter("B > 0", limit = 0))
    )
    val cells = (refs ++ commented :+ ARef.from0(7, 7)).distinct.map { r =>
      (name, ReadQuery.Cell(r.toA1, noStyle = false))
    }
    views ++ searches ++ stats ++ filters ++ cells

  private def rendered(outcome: Outcome, mode: OutputMode): (Int, Rendered) =
    (outcome.exitCode.code, Render(mode)(outcome, "test"))

  /** The fields of the capabilities the streaming reader lacks, projected out of a payload. */
  private def sharedOnly(text: String, mode: OutputMode, query: ReadQuery): String =
    (query, mode) match
      case (_: ReadQuery.Cell, OutputMode.Text) =>
        text.linesIterator
          .filterNot(l => l.startsWith("Dependencies:") || l.startsWith("Dependents:"))
          .mkString("\n")
      case (_: ReadQuery.Cell, OutputMode.Json) =>
        val parsed = ujson.read(text)
        parsed("data") match
          case obj: ujson.Obj =>
            obj.value.remove("dependencies")
            obj.value.remove("dependents")
            obj.value.remove("hidden")
          case _ => ()
        ujson.write(parsed)
      case (_: ReadQuery.Search, OutputMode.Json) =>
        val parsed = ujson.read(text)
        parsed("data")("matches").arr.foreach {
          case obj: ujson.Obj => obj.value.remove("hidden")
          case _ => ()
        }
        ujson.write(parsed)
      case _ => text

  property("in-memory and streaming sources produce equal payloads for every read verb") {
    forAll(genBook) { (book: ParityBook) =>
      withFile(book) { (loaded, path) =>
        val wb = book.wb
        val checks = wb.sheets.flatMap(sheet => queries(wb, sheet)).flatMap { case (flag, query) =>
          Vector(OutputMode.Text, OutputMode.Json).map(mode => (flag, query, mode))
        }
        checks.traverse_ { case (flag, query, mode) =>
          (
            Reads.outcome(query, SheetSource.inMemory(loaded), flag, mode),
            Reads.outcome(query, SheetSource.streaming(path, excel), flag, mode)
          ).mapN { (memory, streaming) =>
            val (memoryExit, memoryOut) = rendered(memory, mode)
            val (streamExit, streamOut) = rendered(streaming, mode)
            val label = s"$query with -s $flag in $mode"
            assertEquals(streamExit, memoryExit, s"exit codes differ for $label")
            assertEquals(streamOut.stderr, memoryOut.stderr, s"stderr differs for $label")
            assertEquals(
              sharedOnly(streamOut.stdout, mode, query),
              sharedOnly(memoryOut.stdout, mode, query),
              s"stdout differs beyond the graph and hidden fields for $label"
            )
            // The projected fields must be honest: the streaming reader says it cannot know
            (query, mode) match
              case (_: ReadQuery.Cell, OutputMode.Text) if memory.ok =>
                assert(streamOut.stdout.contains("Dependencies: (not available in streaming mode)"))
                assert(streamOut.stdout.contains("Dependents: (not available in streaming mode)"))
                assert(!memoryOut.stdout.contains("not available"), memoryOut.stdout)
              case (_: ReadQuery.Cell, OutputMode.Json) if memory.ok =>
                val stream = ujson.read(streamOut.stdout)("data")
                val loadedData = ujson.read(memoryOut.stdout)("data")
                assertEquals(stream("dependencies"), ujson.Null)
                assertEquals(stream("dependents"), ujson.Null)
                assertEquals(stream("hidden"), ujson.Null)
                assertEquals(loadedData("hidden"), ujson.Bool(false))
                assert(loadedData("dependents").arrOpt.isDefined, loadedData("dependents"))
                assert(loadedData("dependencies").arrOpt.isDefined, loadedData("dependencies"))
              case (_: ReadQuery.Search, OutputMode.Json) if memory.ok =>
                val stream = ujson.read(streamOut.stdout)("data")("matches").arr
                val loadedData = ujson.read(memoryOut.stdout)("data")("matches").arr
                stream.foreach(m => assertEquals(m("hidden"), ujson.Null))
                loadedData.foreach(m => assertEquals(m("hidden"), ujson.Bool(false)))
              // filter's JSON rides the envelope as the array it is, never as a string of text
              case (f: ReadQuery.Filter, OutputMode.Json)
                  if memory.ok && f.format == FilterFormat.Json =>
                Vector(memoryOut.stdout, streamOut.stdout).foreach { out =>
                  val data = ujson.read(out)("data")
                  assert(data.arrOpt.isDefined, s"filter data is not an array for $label: $data")
                  data.arr.foreach(row => assert(row("row").numOpt.isDefined, row.toString))
                }
              case _ => ()
          }
        }
      }.unsafeRunSync()
      Prop.passed
    }
  }

  test("a query needing a capability the streaming reader lacks is refused in band") {
    val wb = Workbook(Vector(Sheet("Data").put(ARef.from0(0, 0), CellValue.Number(BigDecimal(1)))))
    ReadTestKit
      .withTempWorkbook(wb) { path =>
        for
          eval <- ReadTestKit.streaming(
            path,
            Some("Data"),
            ReadTestKit.view(None, evalFormulas = true)
          )
          html <- ReadTestKit.streaming(path, Some("Data"), ReadTestKit.view(None, ViewFormat.Html))
          fine <- ReadTestKit.streaming(path, Some("Data"), ReadTestKit.view(None))
        yield
          assertEquals(eval.error.map(_.code), Some("UNSUPPORTED_IN_STREAM"))
          assertEquals(eval.exitCode.code, 2)
          assert(eval.error.exists(_.message.contains("--eval is not supported with --stream")))
          assertEquals(html.error.map(_.code), Some("UNSUPPORTED_IN_STREAM"))
          assert(html.error.exists(_.message.contains("html")), html.error.toString)
          assert(fine.ok, fine.error.toString)
      }
      .unsafeRunSync()
  }

  test("the capability table names what each source answers") {
    assertEquals(Capability.inMemory, Capability.all.toSet)
    assert(Capability.streaming.subsetOf(Capability.inMemory))
    assertEquals(
      Capability.streaming,
      Set(Capability.Values, Capability.Styles, Capability.Formulas, Capability.Comments)
    )
    assertEquals(SheetSource.inMemory(Workbook.empty).capabilities, Capability.inMemory)
    val rows = Capability.json.arr.toVector
    assertEquals(rows.map(_("name").str), Capability.all.map(_.name))
    rows.foreach(r => assertEquals(r.obj.keys.toList, List("name", "doc", "inMemory", "streaming")))
  }
