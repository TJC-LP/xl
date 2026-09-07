package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path}
import java.time.LocalDate

import cats.effect.IO
import cats.syntax.all.*

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * Workbooks the golden corpus runs against. Built in-process on every run (never checked in) so a
 * golden pins the CLI's rendering of known content, not a binary blob; the library's writer is
 * deterministic, so the files are byte-identical from run to run.
 */
object TestFixtures:

  /**
   * Sheet `Data`: text, integers, a decimal, ISO dates and ONE formula (`B4 = SUM(B1:B3)`) whose
   * cached value comes from the library's own recalculation before the write, so cached-value
   * readers (`view`, `--stream view`, `cell`) see `42.5`. Sheet `Summary` is empty.
   */
  def simpleBook(): Workbook = book(firstQuantity = 10)

  /** `simpleBook` with `B1` bumped to 11 (so `B4` recalculates to 43.5): the `diff` counterpart. */
  def changedBook(): Workbook = book(firstQuantity = 11)

  /** One sheet only, so unqualified single-cell reads resolve without `-s`. */
  def singleSheetBook(): Workbook =
    Workbook(Vector(Sheet("Sheet1").put(ref"A1", "solo").put(ref"B1", 7)))

  /**
   * `A1 = A1+1`: a circular reference, so `--eval` cannot evaluate the sheet and `--strict` (on
   * `view` or a write) gates deterministically.
   */
  def circularBook(): Workbook =
    Workbook(Vector(Sheet("Data").put(ref"A1", CellValue.Formula("A1+1", None))))

  /** `simpleBook` plus one workbook-scoped defined name, so `names` has something to list. */
  def namedBook(): Workbook = simpleBook().withDefinedName("Total", "Data!$B$4")

  private def book(firstQuantity: Int): Workbook =
    val data = Sheet("Data")
      .put(ref"A1", "Hello")
      .put(ref"B1", firstQuantity)
      .put(ref"C1", LocalDate.of(2024, 1, 15))
      .put(ref"A2", "World")
      .put(ref"B2", 20)
      .put(ref"C2", LocalDate.of(2024, 6, 30))
      .put(ref"A3", "Total")
      .put(ref"B3", BigDecimal("12.5"))
      .put(ref"B4", CellValue.Formula("SUM(B1:B3)", None))
    Workbook(Vector(data, Sheet("Summary"))).recalculate().workbook

  /** File names the golden args refer to, relative to `<DIR>`. */
  val files: Map[String, () => Workbook] = Map(
    "simple.xlsx" -> (() => simpleBook()),
    "simple-copy.xlsx" -> (() => simpleBook()),
    "changed.xlsx" -> (() => changedBook()),
    "single.xlsx" -> (() => singleSheetBook()),
    "inplace.xlsx" -> (() => simpleBook()),
    "circular.xlsx" -> (() => circularBook()),
    "named.xlsx" -> (() => namedBook())
  )

  /** Write every fixture into a fresh temp directory and return it. */
  def materialize: IO[Path] =
    IO.blocking(Files.createTempDirectory("xl-golden-")).flatTap { dir =>
      files.toList.traverse_ { (name, build) =>
        ExcelIO.instance[IO].write(build(), dir.resolve(name))
      }
    }

  /** Remove the directory and everything the cases wrote into it. */
  def delete(dir: Path): IO[Unit] =
    IO.blocking {
      val entries = Files.list(dir)
      try entries.forEach(p => Files.deleteIfExists(p))
      finally entries.close()
      Files.deleteIfExists(dir)
    }.void
