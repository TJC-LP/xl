package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.addressing.{ARef, CellRange, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.formula.eval.{
  ArrayArithmetic,
  ArrayResult,
  EvalError,
  Evaluator,
  SheetEvaluator
}
import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.{CellStore, Sheet}
import com.tjclp.xl.styles.units.StyleId
import com.tjclp.xl.workbooks.Workbook

/**
 * #891: which way array mode reads a range ([[ArrayArithmetic.readsPositionally]]), at each edge of
 * the decision, and that the decision changes no value.
 *
 * The rule had one condition (#750): position by position when the range has no more positions than
 * the sheet has cells. It has a second now: and at least half of the range lies inside the sheet's
 * extent, unless the range is tiny (eight positions or fewer). So every range #750 sent to its
 * populated box still goes there (the cases named "as before" below), and a range that mostly lies
 * past the data joins them (`A1:A1000` over 100 rows, a whole column of a sheet of more than
 * 1,048,576 cells). For those, and for ranges on either side of each edge, every formula here must
 * evaluate to the same thing under the rule, under an all-positional read and under an all-box
 * read.
 */
class RangeReadDecisionSpec extends FunSuite:
  import RangeReadDecisionSpec.*

  // the 1.2M-cell sheet takes a few seconds to build and read position by position
  override val munitTimeout = scala.concurrent.duration.Duration(180, "s")

  private def positional(sheet: Sheet, range: String): Boolean =
    ArrayArithmetic.readsPositionally(rangeOf(range), sheet)

  test("as before: a huge range on a tiny sheet is read through its box") {
    val corners = sheetOf(List(number(0, 0), number(1048575, 16)))
    assert(!positional(corners, "A:XFD"))
    assert(!positional(corners, "A:Q"))
    assert(!positional(corners, "A1:B2"), "four positions, two cells")
    assert(!positional(Sheet(SheetName.unsafe("S")), "A1"), "an empty sheet has no extent")
  }

  test("as before: a range inside the data of a larger sheet is read position by position") {
    assert(positional(table, "A2:A50"))
    assert(positional(table, "B10:D20"))
    assert(positional(table, "T100"))
    assert(positional(table, "A1:T100"), "exactly as many positions as the sheet has cells")
  }

  test("as before: one position more than the sheet has cells goes to the box") {
    // the same extent with one cell fewer: 2,000 positions against 1,999 cells
    val oneFewer = table.copy(cells = table.cells.removed(ARef.from0(5, 50)))
    assertEquals(oneFewer.cells.extent, table.cells.extent)
    assert(!positional(oneFewer, "A1:T100"))
    assert(positional(oneFewer, "A1:T99"))
    assert(!positional(table, "A:A"), "a whole column against 2,000 cells")
  }

  test("new: a range no larger than the sheet goes to its box when half of it is past the data") {
    // the data ends at row 100: A1:A199 has 100 of 199 positions inside it, A1:A200 100 of 200
    assert(positional(table, "A1:A199"))
    assert(!positional(table, "A1:A200"))
    assert(!positional(table, "A1:A1000"))
    assert(!positional(table, "V1:V10"), "wholly outside the extent")
    assert(positional(table, "S90:U100"), "two of its three columns inside")
    assert(!positional(table, "T92:U100"), "nine of eighteen positions inside")
    // a formatted blank far away stretches the extent, and the rule is the old one again
    val formatted = table.put(Cell(ARef.from0(30, 4999), CellValue.Empty, Some(StyleId(1))))
    assert(positional(formatted, "A1:A1000"))
  }

  test("as before: a tiny range is read position by position wherever it lies") {
    // up to TinyRange positions the extent is not consulted: a box has next to nothing to save
    assertEquals(ArrayArithmetic.TinyRange, 8L)
    assert(positional(table, "T100:U100"), "one of two positions inside")
    assert(positional(table, "V1:V8"), "eight positions, none inside")
    assert(!positional(table, "V1:V9"), "nine positions, none inside")
    assert(positional(table, "X500"), "one cell, far past the data")
    assert(!positional(sheetOf(List(number(0, 0))), "V1:V8"), "eight positions, one cell")
  }

  test("each side of every edge evaluates as both reads do") {
    val ranges = List(
      "A1:A199",
      "A1:A200",
      "A1:A1000",
      "V1:V10",
      "S90:U100",
      "T100:U100",
      "T92:U100",
      "V1:V8",
      "V1:V9",
      "A1:T100",
      "A2:A50",
      "E1:H300",
      "A90:V140",
      "A:A",
      "100:100"
    )
    val sheets = List("table" -> table, "ragged" -> ragged)
    for
      (label, sheet) <- sheets
      range <- ranges
      template <- templates
    do
      val formula = template.replace("$R", range)
      val auto = evalAuto(sheet, formula)
      assertEquals(auto, evalWith(Dense, sheet, formula), s"$label: $formula against positional")
      assertEquals(auto, evalWith(Sparse, sheet, formula), s"$label: $formula against the box")
  }

  test("SUMPRODUCT beside a sparse argument: a raw range past the data is read through its box") {
    // (A1:A1000>0)*1 is sparse under the rule; B1:B1000 (1,000 positions, 2,000 cells, 100 rows of
    // data) was walked position by position and is read through its box now
    val formulas = List(
      "=SUMPRODUCT((A1:A1000>0)*1,B1:B1000)",
      "=SUMPRODUCT((A1:A1000>0)*1,B1:B1000,C1:C1000)",
      "=SUMPRODUCT((A1:A199>0)*1,B1:B199)",
      "=SUMPRODUCT((A1:A200>0)*1,B1:B200)",
      "=SUMPRODUCT((A:A>0)*1,B:B)",
      "=SUMPRODUCT((A1:C1000<>\"\")*1,D1:F1000)"
    )
    for
      (label, sheet) <- List("table" -> table, "ragged" -> ragged)
      formula <- formulas
    do
      val auto = evalAuto(sheet, formula)
      assertEquals(auto, evalWith(Dense, sheet, formula), s"$label: $formula against positional")
      assertEquals(auto, evalWith(Sparse, sheet, formula), s"$label: $formula against the box")
    assertEquals(evalAuto(table, "=SUMPRODUCT((A1:A1000>0)*1,B1:B1000)"), Right(BigDecimal(99099)))
  }

  test("a positioned evaluation: the placement inside a range is read whichever read is chosen") {
    // `xl eval --at`: the formula sits, uncached, at its placement, and a read of that cell is the
    // one-cell cycle (GH-715). The placed record stretches the extent to its row, so A1:A1000 goes
    // to its box with the placement at A50 or A500 (100 and 500 of its 1,000 positions inside the
    // extent) and stays positional with it at A600. The record holds a value, so it lies in the
    // box and is read either way.
    def placedAt(at: ARef, formula: String) =
      val sheet = SheetEvaluator.placed(table, at, formula)
      val book = Workbook(sheet)
      val placement = SheetEvaluator.placement(formula, book, sheet.name, at)
      val result: Any = placement.evaluator.withArrayResults
        .eval(parsed(formula), sheet, Clock.system, Some(book), Some(at))
      (result, placement.wasRead, ArrayArithmetic.readsPositionally(rangeOf("A1:A1000"), sheet))
    val formulas = List("=SUM(A1:A1000*1)", "=SUMPRODUCT((A1:A1000>0)*1,B1:B1000)")
    for
      formula <- formulas
      (at, expectPositional) <- List(
        ARef.from0(0, 499) -> false,
        ARef.from0(0, 599) -> true,
        ARef.from0(0, 49) -> false
      )
    do
      val (result, wasRead, positional) = placedAt(at, formula)
      assert(wasRead, s"$formula at ${at.toA1}: the placement was not read")
      assertEquals(result, Left(EvalError.CircularRef(List(at))), s"$formula at ${at.toA1}")
      assertEquals(positional, expectPositional, s"$formula at ${at.toA1}: how A1:A1000 is read")
    // a placement the formula does not read leaves the value as it is unplaced
    val (sum, sumRead, _) = placedAt(ARef.from0(25, 4), "=SUM(A1:A1000*1)")
    assertEquals(sum, evalAuto(table, "=SUM(A1:A1000*1)"))
    assert(!sumRead)
    val (product, productRead, _) =
      placedAt(ARef.from0(25, 4), "=SUMPRODUCT((A1:A1000>0)*1,B1:B1000)")
    assertEquals(product, Right(BigDecimal(99099)))
    assert(!productRead)
  }

  test("an aggregate folds the cells stored in a range, or every position when positioned") {
    // SUM, COUNT and the rest over one range argument: no rule by size any more. The fold walks
    // the store's cells inside the range once, whatever the range and the sheet are, so there is
    // no box to find first and no edge between two reads; only a positioned evaluation (its
    // placement can be an absent cell) and the laws' forced read go position by position
    val Auto = ArrayArithmetic.RangeRead.Auto
    assert(!FunctionSpecs.aggregateReadsPositions(positioned = false, Auto))
    assert(!FunctionSpecs.aggregateReadsPositions(positioned = false, Sparse))
    assert(FunctionSpecs.aggregateReadsPositions(positioned = false, Dense))
    assert(FunctionSpecs.aggregateReadsPositions(positioned = true, Auto))
    assert(FunctionSpecs.aggregateReadsPositions(positioned = true, Sparse))
    assert(FunctionSpecs.aggregateReadsPositions(positioned = true, Dense))
  }

  test("a bounded range of billions of positions over a few cells costs the cells") {
    // B1:XFC1048576 over cells at its opposite corners: the used range is as large, so clipping
    // to it saves nothing, and the populated box the fold used to read is the range. 17 billion
    // positions, four records. The fold meets the four (the timeout is the backstop: position
    // by position this does not end)
    val sheet = sheetOf(
      List(
        number(0, 1),
        number(1048575, 16382),
        Cell(ARef.from0(700, 500000), CellValue.Text("t")),
        Cell(ARef.from0(9, 9), CellValue.Empty, Some(StyleId(1))),
        number(5, 0) // outside the range
      )
    )
    val range = "B1:XFC1048576"
    val positions = BigDecimal(16382L * 1048576L)
    assertEquals(evalAuto(sheet, s"=SUM($range)"), Right(BigDecimal(1L + 1048575L * 20 + 16382)))
    assertEquals(evalAuto(sheet, s"=COUNT($range)"), Right(BigDecimal(2)))
    assertEquals(evalAuto(sheet, s"=COUNTA($range)"), Right(BigDecimal(3)))
    // every position but the three values is blank, the formatted one included
    assertEquals(evalAuto(sheet, s"=COUNTBLANK($range)"), Right(positions - 3))
    assertEquals(evalAuto(sheet, s"=MAX($range)"), Right(BigDecimal(1048575L * 20 + 16382)))
  }

  test("an aggregate over the stored cells is the aggregate over every position") {
    val ranges = List(
      "A1:A199",
      "A1:A200",
      "A1:A1000",
      "V1:V8",
      "V1:V9",
      "V1:V300",
      "V1:V400",
      "A1:A400",
      "S90:U100",
      "T92:U100",
      "A1:T100",
      "A1:U100",
      "A2:A50",
      "B60:B70",
      "E1:H300",
      "A90:V140",
      "A:A",
      "100:100"
    )
    for
      (label, sheet) <- List("table" -> table, "ragged" -> ragged)
      range <- ranges
      template <- aggregates
    do
      val formula = template.replace("$R", range)
      val auto = evalAuto(sheet, formula)
      // the fold over the stored cells against the fold over every position
      assertEquals(auto, evalWith(Dense, sheet, formula), s"$label: $formula against positional")
      assertEquals(auto, evalWith(Sparse, sheet, formula), s"$label: $formula, forced sparse")
    // the values themselves, in and past the data: rows 1..100 hold r*20
    assertEquals(evalAuto(table, "=SUM(A1:A199)"), Right(BigDecimal(99000)))
    assertEquals(evalAuto(table, "=SUM(A1:A200)"), Right(BigDecimal(99000)))
    assertEquals(evalAuto(table, "=COUNTBLANK(A1:A199)"), Right(BigDecimal(99)))
    assertEquals(evalAuto(table, "=COUNTBLANK(A1:A200)"), Right(BigDecimal(100)))
    assertEquals(evalAuto(table, "=COUNTBLANK(V1:V9)"), Right(BigDecimal(9)))
    assertEquals(evalAuto(table, "=COUNTA(A1:A200)"), Right(BigDecimal(100)))
  }

  test("a positioned aggregate reads its placement inside a range") {
    // a positioned fold reads every position, so its guard sees the placement wherever it is
    def placedAt(at: ARef, formula: String) =
      val sheet = SheetEvaluator.placed(table, at, formula)
      val book = Workbook(sheet)
      val placement = SheetEvaluator.placement(formula, book, sheet.name, at)
      val result: Any = placement.evaluator.withArrayResults
        .eval(parsed(formula), sheet, Clock.system, Some(book), Some(at))
      (result, placement.wasRead)
    for
      formula <- List("=SUM(A1:A1000)", "=COUNT(A1:A1000)", "=COUNTBLANK(A1:A1000)")
      at <- List(ARef.from0(0, 499), ARef.from0(0, 599), ARef.from0(0, 49), ARef.from0(0, 999))
    do
      val (result, wasRead) = placedAt(at, formula)
      assert(wasRead, s"$formula at ${at.toA1}: the placement was not read")
      assertEquals(result, Left(EvalError.CircularRef(List(at))), s"$formula at ${at.toA1}")
    // a placement outside the range leaves the values as they are unplaced
    val (sum, sumRead) = placedAt(ARef.from0(25, 4), "=SUM(A1:A1000)")
    assertEquals(sum, Right(BigDecimal(99000)))
    assert(!sumRead)
    val (blank, blankRead) = placedAt(ARef.from0(25, 4), "=COUNTBLANK(A1:A1000)")
    assertEquals(blank, Right(BigDecimal(900)))
    assert(!blankRead)
  }

  test("the read chosen shows in the array's box, not in its values") {
    def box(sheet: Sheet, formula: String): (Int, Int, Int, Int) =
      evalAuto(sheet, formula) match
        case Right(ar: ArrayResult) => (ar.boxTop, ar.boxLeft, ar.boxRows, ar.boxCols)
        case other => fail(s"expected an array, got $other")
    assertEquals(box(table, "=A1:A199*1"), (0, 0, 199, 1), "positional: the whole range")
    assertEquals(box(table, "=A1:A200*1"), (0, 0, 100, 1), "the box: the rows that hold data")
    assertEquals(box(table, "=A1:A1000*1"), (0, 0, 100, 1))
  }

  test("a whole column of a sheet with more cells than a column has rows") {
    // 60,000 rows of 20 columns: A:A has 1,048,576 positions against 1,200,000 cells, so #750's
    // condition alone read all of them; 60,000 lie inside the extent
    val rows = 60000
    val big = sheetOf(Iterator.range(0, rows * 20).map(i => number(i / 20, i % 20)))
    val column = rangeOf("A:A")
    assert(column.cellCount <= big.cells.size.toLong, "the count alone keeps it positional")
    assert(!ArrayArithmetic.readsPositionally(column, big))
    assert(ArrayArithmetic.readsPositionally(rangeOf("A1:A60000"), big))
    assert(ArrayArithmetic.readsPositionally(rangeOf("A1:T60000"), big))
    evalAuto(big, "=A:A*1") match
      case Right(ar: ArrayResult) =>
        assertEquals((ar.rows, ar.cols), (1048576, 1))
        assertEquals((ar.boxTop, ar.boxLeft, ar.boxRows, ar.boxCols), (0, 0, rows, 1))
      case other => fail(s"expected an array, got $other")
    List(
      "=SUM(A:A*1)",
      "=SUM((A:A=\"\")*1)",
      "=COUNT(A:A*1)",
      "=SUMPRODUCT((A:A>0)*1,B:B)",
      "=ROWS(A:A*1)"
    ).foreach { formula =>
      val auto = evalAuto(big, formula)
      assert(auto.isInstanceOf[Right[?, ?]], s"$formula: $auto")
      assertEquals(auto, evalWith(Dense, big, formula), formula)
    }
    // sum of column A: row r holds r*20
    assertEquals(evalAuto(big, "=SUM(A:A*1)"), Right(BigDecimal(20L * rows * (rows - 1) / 2)))
    assertEquals(evalAuto(big, "=SUM((A:A=\"\")*1)"), Right(BigDecimal(1048576 - rows)))
  }

object RangeReadDecisionSpec:
  private val Dense = ArrayArithmetic.RangeRead.Dense
  private val Sparse = ArrayArithmetic.RangeRead.Sparse

  private def number(row: Int, col: Int): Cell =
    Cell(ARef.from0(col, row), CellValue.Number(BigDecimal(row.toLong * 20 + col)))

  private def sheetOf(cells: IterableOnce[Cell]): Sheet =
    Sheet(SheetName.unsafe("S")).copy(cells = CellStore.from(cells))

  private def rangeOf(a1: String): CellRange =
    CellRange.parse(a1).fold(err => throw new IllegalArgumentException(s"$a1: $err"), identity)

  /** A1:T100, every cell a number: 2,000 cells, extent and used range A1:T100. */
  private val table: Sheet = sheetOf(Iterator.range(0, 2000).map(i => number(i / 20, i % 20)))

  /**
   * The table with what a real sheet has: holes, text, booleans, error values, cached and uncached
   * formulas (one of them failing), and formatted blanks inside and just outside the data.
   */
  private val ragged: Sheet =
    val holes = (0 until 2000 by 7).map(i => ARef.from0(i % 20, i / 20))
    val base = holes.foldLeft(table)((s, at) => s.copy(cells = s.cells.removed(at)))
    val extras = List(
      Cell(ARef.from0(0, 2), CellValue.Text("a")),
      Cell(ARef.from0(1, 3), CellValue.Bool(true)),
      Cell(ARef.from0(0, 40), CellValue.Error(CellError.Div0)),
      Cell(ARef.from0(4, 98), CellValue.Error(CellError.NA)),
      Cell(ARef.from0(0, 60), CellValue.Formula("B61*2", Some(CellValue.Number(BigDecimal(7))))),
      Cell(ARef.from0(1, 61), CellValue.Formula("C62+1", None)),
      Cell(ARef.from0(18, 95), CellValue.Formula("1/0", None)),
      Cell(ARef.from0(0, 99), CellValue.Empty, Some(StyleId(1))),
      Cell(ARef.from0(21, 120), CellValue.Empty, Some(StyleId(1))),
      Cell(ARef.from0(0, 150), CellValue.Number(BigDecimal(5)))
    )
    base.copy(cells = base.cells.putAll(extras))

  private val templates: List[String] = List(
    "=$R*2",
    "=$R=\"\"",
    "=$R&\"x\"",
    "=IF($R=\"\",1,0)",
    "=SUM($R*1)",
    "=SUM(($R=\"\")*1)",
    "=COUNT($R*1)",
    "=COUNTA($R&\"\")",
    "=MAX($R+$R)",
    "=SUMPRODUCT(($R>0)*1)",
    "=SUMPRODUCT(($R=\"\")*($R=\"\"))",
    "=IFERROR($R/$R,\"e\")",
    "=ISBLANK($R)",
    "=TRANSPOSE($R)",
    "=INDEX($R*1,1,1)",
    "=ROWS($R*1)"
  )

  /** Each aggregate over one raw range: the numeric folds, the three counts, and two ranges. */
  private val aggregates: List[String] = List(
    "=SUM($R)",
    "=COUNT($R)",
    "=COUNTA($R)",
    "=COUNTBLANK($R)",
    "=AVERAGE($R)",
    "=MAX($R)",
    "=MIN($R)",
    "=SUM($R,$R)",
    "=COUNTBLANK($R)+COUNTA($R)",
    "=IFERROR(SUM($R),\"e\")"
  )

  private def parsed(formula: String) =
    FormulaParser
      .parse(formula)
      .fold(err => throw new IllegalArgumentException(s"$formula: $err"), identity)

  private def evalAuto(sheet: Sheet, formula: String): Any =
    Evaluator.arrayInstance(Rng.seeded(0)).eval(parsed(formula), sheet, Clock.system, None, None)

  private def evalWith(mode: ArrayArithmetic.RangeRead, sheet: Sheet, formula: String): Any =
    Evaluator
      .arrayInstanceReading(Rng.seeded(0), mode)
      .eval(parsed(formula), sheet, Clock.system, None, None)
