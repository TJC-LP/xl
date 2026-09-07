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
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.cli.{FilterFormat, ViewFormat}
import com.tjclp.xl.cli.contract.{Outcome, OutputMode, Render, Rendered}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * W2.4's acceptance law: for every read verb, the loaded workbook and the streaming reader produce
 * equal payloads when the book uses only the capabilities both sources share
 * ([[Capability.streaming]]: values, styles, cached formulas, comments). The generated books hold
 * text, numbers, booleans, errors and empties under a few number formats — no formulas (the `graph`
 * capability), no hidden lines, merges or hyperlinks — written to a file and read back through both
 * strategies; every query runs in both output modes and the rendered stdout, stderr and exit code
 * must agree.
 *
 * `cell` is the one verb whose payload names a capability the streaming reader lacks: `dependents`
 * (`Dependents: (not available in streaming mode)` / `"dependents": null`). The law compares the
 * rest of the payload and asserts that the difference is exactly that field.
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
      1 -> Gen.const("x y")
    )

  private val genNumber: Gen[BigDecimal] =
    Gen.frequency(
      5 -> Gen.choose(-1000, 1000).map(BigDecimal(_)),
      3 -> Gen.choose(-100000, 100000).map(n => BigDecimal(n) / 100),
      1 -> Gen.const(BigDecimal("12345678901234567")),
      1 -> Gen.const(BigDecimal("0.125"))
    )

  private val genValue: Gen[Option[CellValue]] =
    Gen.frequency(
      4 -> genText.map(t => Some(CellValue.Text(t))),
      4 -> genNumber.map(n => Some(CellValue.Number(n))),
      1 -> Gen.oneOf(true, false).map(b => Some(CellValue.Bool(b))),
      1 -> Gen
        .oneOf(CellError.Div0, CellError.NA, CellError.Ref)
        .map(e => Some(CellValue.Error(e))),
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

  /** A sheet: up to 6 rows by 4 columns, a header row of words first. */
  private def genSheet(name: String): Gen[Sheet] =
    for
      rows <- Gen.choose(1, 6)
      cols <- Gen.choose(1, 4)
      headers <- Gen.listOfN(cols, Gen.identifier.map(_.take(6)).suchThat(_.nonEmpty))
      cells <- Gen.listOfN(rows * cols, Gen.zip(genValue, genNumFmt))
    yield
      val blank: Sheet = Sheet(SheetName.unsafe(name))
      val withHeaders = headers.zipWithIndex.foldLeft(blank) { case (s, (h, col)) =>
        s.put(Cell(ARef.from0(col, 0), CellValue.Text(h)))
      }
      cells.zipWithIndex.foldLeft(withHeaders) { case (s, ((value, numFmt), i)) =>
        val ref = ARef.from0(i % cols, 1 + i / cols)
        value.fold(s) { v =>
          val put: Sheet = s.put(Cell(ref, v))
          numFmt.fold(put)(fmt =>
            put.styleAt(ref.toA1, CellStyle.default.withNumFmt(fmt)).getOrElse(put)
          )
        }
      }

  private val genBook: Gen[Workbook] =
    for
      two <- Gen.oneOf(true, false)
      data <- genSheet("Data")
      other <- genSheet("Other")
    yield Workbook(if two then Vector(data, other) else Vector(data))

  // --- The law -----------------------------------------------------------------------------------

  private val excel = ReadTestKit.excel

  private def withFile[A](wb: Workbook)(f: (Workbook, Path) => IO[A]): IO[A] =
    IO.blocking(Files.createTempFile("xl-parity-", ".xlsx")).flatMap { path =>
      (excel.write(wb, path) *> excel.read(path).flatMap(loaded => f(loaded, path)))
        .guarantee(IO.blocking(Files.deleteIfExists(path)).void)
    }

  private def queries(wb: Workbook, sheet: Sheet): Vector[(Option[String], ReadQuery)] =
    val name = Some(sheet.name.value)
    val someText = sheet.cells.values.map(_.value).collectFirst { case CellValue.Text(s) => s }
    val refs = sheet.cells.keys.toVector.sortBy(r => (r.row.index0, r.col.index0)).take(4)
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
      ReadTestKit.view(None, showFormulas = true)
    ).map(q => (name, q: ReadQuery))
    val searches = Vector(
      (None, ReadQuery.Search("\\d", 50, None)),
      (None, ReadQuery.Search("\\d", 2, None)),
      (name, ReadQuery.Search("TRUE|FALSE", 0, None)),
      (None, ReadQuery.Search(java.util.regex.Pattern.quote(someText.getOrElse("Total")), 50, None))
    )
    val stats = Vector((name, ReadQuery.Stats("A1:E8")), (name, ReadQuery.Stats("B2:B3")))
    val filters = Vector(
      (name, ReadTestKit.filter("B > 0")),
      (name, ReadTestKit.filter("A IS NOT EMPTY", format = FilterFormat.Csv, header = true)),
      (name, ReadTestKit.filter("B >= 0 OR C = TRUE", format = FilterFormat.Json, limit = 2)),
      (name, ReadTestKit.filter("A IS EMPTY", columns = Some("A:B")))
    )
    val cells = (refs :+ ARef.from0(7, 7)).map(r => (name, ReadQuery.Cell(r.toA1, noStyle = false)))
    views ++ searches ++ stats ++ filters ++ cells

  private def rendered(outcome: Outcome, mode: OutputMode): (Int, Rendered) =
    (outcome.exitCode.code, Render(mode)(outcome, "test"))

  /** `cell`'s payload with the `dependents` capability projected out, in either mode. */
  private def withoutDependents(text: String, mode: OutputMode): String = mode match
    case OutputMode.Text => text.linesIterator.filterNot(_.startsWith("Dependents:")).mkString("\n")
    case OutputMode.Json =>
      val parsed = ujson.read(text)
      parsed("data") match
        case obj: ujson.Obj => obj.value.remove("dependents")
        case _ => ()
      ujson.write(parsed)

  property("in-memory and streaming sources produce equal payloads for every read verb") {
    forAll(genBook) { (wb: Workbook) =>
      withFile(wb) { (loaded, path) =>
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
            query match
              case _: ReadQuery.Cell if memory.ok =>
                // the one declared difference: the streaming reader cannot know the dependents
                assertEquals(
                  withoutDependents(streamOut.stdout, mode),
                  withoutDependents(memoryOut.stdout, mode),
                  s"cell payload differs beyond dependents for $label"
                )
                mode match
                  case OutputMode.Text =>
                    assert(
                      streamOut.stdout.contains("Dependents: (not available in streaming mode)")
                    )
                    assert(memoryOut.stdout.contains("Dependents: (none)"), memoryOut.stdout)
                  case OutputMode.Json =>
                    assertEquals(ujson.read(streamOut.stdout)("data")("dependents"), ujson.Null)
              case _ =>
                assertEquals(streamOut.stdout, memoryOut.stdout, s"stdout differs for $label")
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
    assert(!Capability.streaming.contains(Capability.Eval))
    assert(!Capability.streaming.contains(Capability.Render))
    assert(!Capability.streaming.contains(Capability.Hidden))
    assert(!Capability.streaming.contains(Capability.Graph))
    assertEquals(SheetSource.inMemory(Workbook.empty).capabilities, Capability.inMemory)
    val rows = Capability.json.arr.toVector
    assertEquals(rows.map(_("name").str), Capability.all.map(_.name))
    rows.foreach(r => assertEquals(r.obj.keys.toList, List("name", "doc", "inMemory", "streaming")))
  }
