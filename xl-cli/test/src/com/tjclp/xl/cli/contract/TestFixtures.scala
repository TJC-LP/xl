package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path}
import java.time.LocalDate

import cats.effect.IO
import cats.syntax.all.*

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{Row, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, Comment}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.RowProperties
import com.tjclp.xl.workbooks.{CalcPr, DefinedName}

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

  /**
   * The `describe`/`deps` fixture (ADR-017 §2.10). `Sheet1` holds two constants and their sum plus
   * a merge, a comment and a freeze pane; `Sheet2!A1` reads `Sheet1!A1` and `Sheet2!B1` chains once
   * more; `Hidden` is a hidden sheet with a hidden row; one visible and one hidden defined name.
   * Recalculated before the write, so every formula is cached.
   */
  def linkedBook(): Workbook =
    val sheet1 = Sheet("Sheet1")
      .put(ref"A1", 5)
      .put(ref"A2", 7)
      .put(ref"A3", CellValue.Formula("SUM(A1:A2)", None))
      .merge(ref"B1:C1")
      .comment(ref"A1", Comment.plainText("input", Some("qa")))
      .freezeAt(ref"B2")
    val sheet2 = Sheet("Sheet2")
      .put(ref"A1", CellValue.Formula("Sheet1!A1*2", None))
      .put(ref"B1", CellValue.Formula("A1+1", None))
    val hidden = Sheet("Hidden")
      .put(ref"A1", "x")
      .setRowProperties(Row.from1(2), RowProperties(hidden = true))
    val recalculated = Workbook(Vector(sheet1, sheet2, hidden)).recalculate().workbook
    val withState = recalculated
      .setSheetState(SheetName.unsafe("Hidden"), Some("hidden"))
      .getOrElse(recalculated)
    withState.copy(metadata =
      withState.metadata.copy(definedNames =
        Vector(
          DefinedName("Total", "Sheet1!$A$3"),
          DefinedName("Secret", "Sheet1!$A$1", hidden = true)
        )
      )
    )

  /**
   * The `audit` fixture: on `Calc` one cell per bucket — a cached `#DIV/0!`, an uncached formula, a
   * `TODAY()`, an `INDIRECT`, a two-cycle, an external-workbook reference, an unparseable function
   * and a reader of an undefined name — plus a second uncached formula on `Notes` (so `-s Calc` has
   * something to exclude) and an iterative `calcPr`. Written WITHOUT recalculation: the caches are
   * the fixture.
   */
  def dirtyBook(): Workbook =
    def cached(expr: String, value: Int): CellValue =
      CellValue.Formula(expr, Some(CellValue.Number(BigDecimal(value))))
    val calc = Sheet("Calc")
      .put(ref"A1", CellValue.Formula("1/0", Some(CellValue.Error(CellError.Div0))))
      .put(ref"B1", CellValue.Formula("A1+1", None))
      .put(ref"C1", cached("TODAY()", 45000))
      .put(ref"D1", cached("INDIRECT(\"A2\")", 0))
      .put(ref"E1", cached("F1+1", 0))
      .put(ref"F1", cached("E1+1", 0))
      .put(ref"G1", cached("'[1]Ext'!A1", 3))
      .put(ref"H1", cached("UNSUPPORTED(1)", 0))
      .put(ref"I1", cached("NoSuchName*2", 0))
    val notes = Sheet("Notes").put(ref"A1", CellValue.Formula("1+1", None))
    Workbook(Vector(calc, notes)).withCalcPr(
      CalcPr(
        iterativeCalculation = true,
        maxIterations = Some(100),
        maxChange = Some(BigDecimal("0.001"))
      )
    )

  /**
   * The `cell` range-gap fixture (ADR-017 §2.10): `B1 = SUM(A1:A5)` over a column with only A1 and
   * A3 occupied, and `C1 = SUM(Missing!A1:A3)` over a sheet the workbook does not have. Pins that
   * `Dependencies` lists a range's OCCUPIED cells — empty cells and absent-sheet ranges are not
   * listed. Caches authored explicitly (no recalculation): the graph, not the values, is the point.
   */
  def gapsBook(): Workbook =
    val data = Sheet("Data")
      .put(ref"A1", 1)
      .put(ref"A3", 3)
      .put(ref"B1", CellValue.Formula("SUM(A1:A5)", Some(CellValue.Number(BigDecimal(4)))))
      .put(ref"C1", CellValue.Formula("SUM(Missing!A1:A3)", Some(CellValue.Error(CellError.Ref))))
    Workbook(Vector(data))

  /**
   * One sheet, one cell: `A1 = 12345678901234567`, an integer a `Double` cannot hold (it rounds to
   * …568). Pins that every JSON rendering — bare `--format json` and the `--json` envelope's `data`
   * for `view`, `eval` and `evala` — prints the number lexeme exactly.
   */
  def preciseBook(): Workbook =
    Workbook(Vector(Sheet("Data").put(ref"A1", BigDecimal("12345678901234567"))))

  /**
   * Sheet names a formula must quote (GH-608, GH-609): `On-Premise` (hyphen) and `M&A` (ampersand).
   * `Summary!G9` reads `'On-Premise'!G9`, so `cell` and `deps` must both print the quoted
   * qualifier; `Summary!I23` mentions `'M&A'` inside a call to a function the parser does not know
   * (`ZZZNOTAFUNC`, a stable parse failure), so `rename-sheet M&A …` must refuse with a typed code,
   * the cell's location and the parser's own diagnostic. Caches authored explicitly.
   */
  def qualifiedBook(): Workbook =
    def cached(expr: String, value: Int): CellValue =
      CellValue.Formula(expr, Some(CellValue.Number(BigDecimal(value))))
    val onPremise = Sheet("On-Premise").put(ref"G9", 5)
    val deals = Sheet("M&A").put(ref"I12", 7)
    val summary = Sheet("Summary")
      .put(ref"G9", cached("'On-Premise'!G9*2", 10))
      .put(ref"I23", cached("IF(ZZZNOTAFUNC(1)=1,'M&A'!I12,0)", 0))
    Workbook(Vector(onPremise, deals, summary))

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
    "named.xlsx" -> (() => namedBook()),
    "linked.xlsx" -> (() => linkedBook()),
    "dirty.xlsx" -> (() => dirtyBook()),
    "gaps.xlsx" -> (() => gapsBook()),
    "precise.xlsx" -> (() => preciseBook()),
    "qualified.xlsx" -> (() => qualifiedBook())
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
