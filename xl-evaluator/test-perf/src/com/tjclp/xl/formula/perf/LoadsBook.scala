package com.tjclp.xl.formula.perf

import scala.util.Random

import com.tjclp.xl.addressing.{ARef, Column, SheetName}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * Synthetic "freight loads" workbook for the recalculation benchmark (#767).
 *
 *   - `Inputs`: four scalars behind single-cell defined names (`year_from`, `year_to`, `own_max`,
 *     `zero_max`).
 *   - `Loads`: `rows` data rows of ten keyed columns and `fillerCols` numeric filler columns; with
 *     `helpers`, three helper formula columns per row (`win`, `spend`, `cpm`). Every keyed and
 *     helper column has a whole-column defined name (`ld_lane`, `ld_spend`, ...) over rows
 *     2..rows+1.
 *   - `Summary`: `aggs` COUNTIFS/SUMIFS/MINIFS/MAXIFS formulas, eight per row, keyed on (`$A` lane,
 *     `$B` carrier) with two to nine criteria pairs.
 *   - `Rate History`: `rh` rows keyed on (`$B` lane, `$E` carrier) with `rhCols` aggregate columns
 *     (MIN/MAX cost per mile, then per-year COUNTIFS and SUMIFS of cost and miles).
 *
 * `Crit.Helper` writes the criteria against the helper columns (`ld_spend,1`, MIN/MAX over
 * `ld_cpm`); `Crit.Raw` expands them into the equivalent raw-column criteria and takes MIN/MAX over
 * `ld_cost`. Lanes and carriers scale with `rows`, so the size of an equality bucket stays flat as
 * the book grows while a scan grows with it.
 *
 * Pure and deterministic: a `Shape` always builds the same workbook (one seeded `Random`, no IO).
 * Ported from the profiling generator of the #767 investigation.
 */
object LoadsBook:

  enum Crit derives CanEqual:
    case Helper, Raw

  /**
   * One benchmark shape. `lanes` and `carriers` of 0 scale with `rows`: `max(20, rows / 50)` lanes
   * and `max(120, rows / 40)` carriers. `wholeColumns` binds the `ld_*` names to whole columns
   * (`Loads!$F:$F`, header row included) instead of rows 2..rows+1, which sends every criteria call
   * through the used-range bound.
   */
  final case class Shape(
    name: String,
    rows: Int,
    aggs: Int,
    rh: Int,
    helpers: Boolean,
    crit: Crit,
    rhCols: Int = 14,
    lanes: Int = 0,
    carriers: Int = 0,
    carriersPerLane: Int = 4,
    fillerCols: Int = 20,
    wholeColumns: Boolean = false,
    seed: Long = 0x5eedL
  ):
    def nLanes: Int = if lanes > 0 then lanes else math.max(20, rows / 50)
    def nCarriers: Int = if carriers > 0 then carriers else math.max(120, rows / 40)
    def useHelperCrit: Boolean = helpers && crit == Crit.Helper

  /** Two thousand formulas, helper criteria: the CI shape (counters, oracle, heap). */
  val Tiny: Shape = Shape("Tiny", rows = 400, aggs = 64, rh = 24, helpers = true, Crit.Helper)

  /** 10k rows, helper criteria: today's graph-bound shape (38k formulas, ~92M legacy edges). */
  val H10: Shape = Shape("H10", rows = 10000, aggs = 1000, rh = 500, helpers = true, Crit.Helper)

  /** 20k rows, raw criteria and no helper columns: today's eval-bound shape. */
  val R20: Shape = Shape("R20", rows = 20000, aggs = 1000, rh = 500, helpers = false, Crit.Raw)

  /** The real workbook's shape: 33,233 rows, ~4k summary and 6k x 14 Rate History formulas. */
  val Real: Shape = Shape("Real", rows = 33233, aggs = 4000, rh = 6000, helpers = true, Crit.Helper)

  val presets: Vector[Shape] = Vector(Tiny, H10, R20, Real)

  def preset(name: String): Option[Shape] = presets.find(_.name.equalsIgnoreCase(name))

  /** A built book plus the formula census and the oracle sample cells. */
  final case class Built(
    shape: Shape,
    workbook: Workbook,
    dataFormulas: Int,
    summaryFormulas: Int,
    rhFormulas: Int,
    sampleRefs: Vector[String]
  ):
    def totalFormulas: Int = dataFormulas + summaryFormulas + rhFormulas

  // Loads column layout (0-based)
  private val ColFlagA = 0
  private val ColFlagB = 1
  private val ColLane = 2
  private val ColCarrier = 3
  private val ColYear = 4
  private val ColCost = 5
  private val ColMiles = 6
  private val ColHasCost = 7
  private val ColOwn = 8
  private val ColZero = 9
  private val ColFiller0 = 10

  private val FirstRow = 2
  private val PerSummaryRow = 8
  private val Years = Vector(2021, 2022, 2023, 2024, 2025, 2020, 2019)
  private val Inputs =
    Vector("year_from" -> 2021, "year_to" -> 2024, "own_max" -> 6, "zero_max" -> 2)

  def colLetter(c: Int): String = Column.from0(c).toLetter
  def laneKey(i: Int): String = f"LN$i%05d"
  def carrierKey(i: Int): String = f"CR$i%04d"

  private def num(d: Double): CellValue = CellValue.Number(BigDecimal(d))
  private def int(i: Int): CellValue = CellValue.Number(BigDecimal(i))
  private def text(s: String): CellValue = CellValue.Text(s)
  private def fx(s: String): CellValue = CellValue.Formula(s)

  /** (0-based column, 1-based row, value) */
  private type Put = (Int, Int, CellValue)

  private def sheet(name: String, puts: IterableOnce[Put]): Sheet =
    val cells = puts.iterator.map { case (c, r, v) =>
      val ref = ARef.from0(c, r - 1)
      ref -> Cell(ref, v)
    }.toMap
    Sheet(SheetName.unsafe(name), cells = cells)

  /** One Loads data row's key, drawn from the seeded generator. */
  private final case class LoadRow(lane: Int, carrier: Int)

  def build(shape: Shape): Built =
    val rng = new Random(shape.seed)
    val last = shape.rows + 1
    val nCols = ColFiller0 + shape.fillerCols
    val helperWin = nCols
    val helperSpend = nCols + 1
    val helperCpm = nCols + 2

    // ---------- Inputs ----------
    val inputsSheet = sheet(
      "Inputs",
      Inputs.zipWithIndex.flatMap { case ((n, v), i) =>
        Vector((0, i + 1, text(n)), (1, i + 1, int(v)))
      }
    )

    // ---------- Loads ----------
    val headers =
      Vector("flag_a", "flag_b", "lane", "carrier", "year", "cost", "miles", "has_cost") ++
        Vector("own", "zero") ++ (0 until shape.fillerCols).map(i => s"v$i")
    val helperHeaders =
      if shape.helpers then
        Vector(
          (helperWin, 1, text("win")),
          (helperSpend, 1, text("spend")),
          (helperCpm, 1, text("cpm"))
        )
      else Vector.empty
    val loadRows = Vector.newBuilder[LoadRow]
    val loadPuts = Vector.newBuilder[Put]
    loadPuts ++= headers.zipWithIndex.map((h, c) => (c, 1, text(h)))
    loadPuts ++= helperHeaders
    (0 until shape.rows).foreach { i =>
      val r = FirstRow + i
      val u = rng.nextDouble()
      // skewed: low lane ids are hot, as in real loads
      val lane = math.min(shape.nLanes - 1, (shape.nLanes * u * u).toInt)
      // each lane is served by ~carriersPerLane carriers, so (lane, carrier) keys repeat
      val carrier = (lane * 7 + rng.nextInt(shape.carriersPerLane)) % shape.nCarriers
      loadRows += LoadRow(lane, carrier)
      val flagA = if rng.nextDouble() < 0.9 then 1 else 0
      val flagB = if rng.nextDouble() < 0.85 then 0 else 1
      val year = 2019 + rng.nextInt(7)
      val miles = 50 + rng.nextInt(2500)
      val hasCost = rng.nextDouble() < 0.92
      val cost =
        if hasCost then num(math.rint(miles * (1.5 + rng.nextDouble() * 2.0) * 100) / 100)
        else int(0)
      val own = rng.nextInt(10)
      val zero = rng.nextInt(5)
      loadPuts ++= Vector(
        (ColFlagA, r, int(flagA)),
        (ColFlagB, r, int(flagB)),
        (ColLane, r, text(laneKey(lane))),
        (ColCarrier, r, text(carrierKey(carrier))),
        (ColYear, r, int(year)),
        (ColCost, r, cost),
        (ColMiles, r, int(miles)),
        (ColHasCost, r, int(if hasCost then 1 else 0)),
        (ColOwn, r, int(own)),
        (ColZero, r, int(zero))
      )
      loadPuts ++= (0 until shape.fillerCols).map(f =>
        (ColFiller0 + f, r, num(math.rint(rng.nextDouble() * 100000) / 100))
      )
      if shape.helpers then
        val a = colLetter(ColFlagA)
        val b = colLetter(ColFlagB)
        val y = colLetter(ColYear)
        val o = colLetter(ColOwn)
        val h = colLetter(ColHasCost)
        val z = colLetter(ColZero)
        val w = colLetter(helperWin)
        val sp = colLetter(helperSpend)
        val cst = colLetter(ColCost)
        val m = colLetter(ColMiles)
        loadPuts ++= Vector(
          (
            helperWin,
            r,
            fx(s"($a$r=1)*($b$r=0)*($y$r>=year_from)*($y$r<=year_to)*($o$r<=own_max)")
          ),
          (helperSpend, r, fx(s"$w$r*($h$r=1)*($z$r<=zero_max)")),
          (helperCpm, r, fx(s"""IF($sp$r*$m$r>0,$cst$r/$m$r,"")"""))
        )
    }
    val rowKeys = loadRows.result()
    val loadsSheet = sheet("Loads", loadPuts.result())

    def wholeColumn(c: Int): String =
      if shape.wholeColumns then s"Loads!$$${colLetter(c)}:$$${colLetter(c)}"
      else s"Loads!$$${colLetter(c)}$$$FirstRow:$$${colLetter(c)}$$$last"
    val columnNames = Vector(
      "ld_flag_a" -> ColFlagA,
      "ld_flag_b" -> ColFlagB,
      "ld_lane" -> ColLane,
      "ld_carrier" -> ColCarrier,
      "ld_year" -> ColYear,
      "ld_cost" -> ColCost,
      "ld_miles" -> ColMiles,
      "ld_has_cost" -> ColHasCost,
      "ld_own" -> ColOwn,
      "ld_zero" -> ColZero
    ) ++ (if shape.helpers then
            Vector("ld_win" -> helperWin, "ld_spend" -> helperSpend, "ld_cpm" -> helperCpm)
          else Vector.empty)

    // criteria fragments
    val rawWin =
      """ld_flag_a,1,ld_flag_b,0,ld_year,">="&year_from,ld_year,"<="&year_to,ld_own,"<="&own_max"""
    val spendCrit =
      if shape.useHelperCrit then "ld_spend,1"
      else rawWin + """,ld_has_cost,1,ld_zero,"<="&zero_max"""
    val winCrit = if shape.useHelperCrit then "ld_win,1" else rawWin
    val minMaxRange = if shape.useHelperCrit then "ld_cpm" else "ld_cost"
    val minMaxExtra = if shape.useHelperCrit then "" else """,ld_cost,">0""""

    // ---------- Summary ----------
    val summaryRows = (shape.aggs + PerSummaryRow - 1) / PerSummaryRow
    def summaryWidth(sr: Int): Int = math.min(PerSummaryRow, shape.aggs - sr * PerSummaryRow)
    val summaryPuts = Vector.newBuilder[Put]
    summaryPuts ++= Vector((0, 1, text("lane")), (1, 1, text("carrier")))
    (0 until summaryRows).foreach { sr =>
      val r = 2 + sr
      val k = rowKeys(rng.nextInt(shape.rows))
      summaryPuts ++= Vector((0, r, text(laneKey(k.lane))), (1, r, text(carrierKey(k.carrier))))
      val key = s"ld_lane,$$A$r,ld_carrier,$$B$r"
      val templates = Vector(
        s"COUNTIFS($key,$spendCrit)",
        s"SUMIFS(ld_cost,$key,$spendCrit)",
        s"SUMIFS(ld_miles,$key,$spendCrit)",
        s"MINIFS($minMaxRange,$key,$spendCrit$minMaxExtra)",
        s"MAXIFS($minMaxRange,$key,$spendCrit)",
        s"COUNTIFS(ld_lane,$$A$r,$winCrit)",
        s"""SUMIFS(ld_cost,ld_carrier,$$B$r,$winCrit,ld_year,">="&year_from,ld_year,"<="&year_to)""",
        s"""COUNTIFS($key,ld_year,">="&year_from,ld_year,"<="&year_to,ld_own,"<="&own_max,ld_has_cost,1,ld_zero,"<="&zero_max,$spendCrit)"""
      )
      (0 until summaryWidth(sr)).foreach(t => summaryPuts += ((2 + t, r, fx(templates(t)))))
    }
    val summarySheet = sheet("Summary", summaryPuts.result())

    // ---------- Rate History ----------
    val yearTemplates: Vector[(Int, Int) => String] = Years.indices.toVector.flatMap { _ =>
      Vector[(Int, Int) => String](
        (r, c) => s"COUNTIFS(ld_lane,$$B$r,ld_carrier,$$E$r,ld_year,${colLetter(c)}$$1,$spendCrit)",
        (r, c) =>
          s"SUMIFS(ld_cost,ld_lane,$$B$r,ld_carrier,$$E$r,ld_year,${colLetter(c)}$$1,$spendCrit)",
        (r, c) =>
          s"SUMIFS(ld_miles,ld_lane,$$B$r,ld_carrier,$$E$r,ld_year,${colLetter(c)}$$1,$spendCrit)"
      )
    }
    val rhTemplates: Vector[(Int, Int) => String] = Vector[(Int, Int) => String](
      (r, _) => s"MINIFS($minMaxRange,ld_lane,$$B$r,ld_carrier,$$E$r,$spendCrit$minMaxExtra)",
      (r, _) => s"MAXIFS($minMaxRange,ld_lane,$$B$r,ld_carrier,$$E$r,$spendCrit)"
    ) ++ yearTemplates
    val rhCols = math.min(shape.rhCols, rhTemplates.size)
    val rhPuts = Vector.newBuilder[Put]
    rhPuts ++= Vector((1, 1, text("lane")), (4, 1, text("carrier")))
    rhPuts ++= (0 until rhCols).map { j =>
      val header =
        if j == 0 then text("min_cpm")
        else if j == 1 then text("max_cpm")
        else int(Years((j - 2) / 3))
      (5 + j, 1, header)
    }
    (0 until shape.rh).foreach { h =>
      val r = 2 + h
      val k = rowKeys(rng.nextInt(shape.rows))
      rhPuts ++= Vector(
        (1, r, text(laneKey(k.lane))),
        (2, r, text("origin")),
        (3, r, text("dest")),
        (4, r, text(carrierKey(k.carrier)))
      )
      (0 until rhCols).foreach(j => rhPuts += ((5 + j, r, fx(rhTemplates(j)(r, 5 + j)))))
    }
    val rhSheet = sheet("Rate History", rhPuts.result())

    val base = Workbook(Vector(inputsSheet, loadsSheet, summarySheet, rhSheet))
    val withScalars = Inputs.zipWithIndex.foldLeft(base) { case (wb, ((n, _), i)) =>
      wb.withDefinedName(n, s"Inputs!$$B$$${i + 1}")
    }
    val workbook = columnNames.foldLeft(withScalars) { case (wb, (n, c)) =>
      wb.withDefinedName(n, wholeColumn(c))
    }
    // oracle samples: every template of the first and middle Summary rows, the helper cells of
    // three Loads rows, the first five Rate History columns of four rows, then the last Summary row
    def summarySamples(sr: Int): Vector[String] =
      (0 until summaryWidth(sr)).toVector.map(t => s"Summary!${colLetter(2 + t)}${2 + sr}")
    def rhSamples(h: Int): Vector[String] =
      (0 until math.min(5, rhCols)).toVector.map(j => s"'Rate History'!${colLetter(5 + j)}${2 + h}")
    val helperSamples =
      if shape.helpers then
        Vector(2, 3, 50)
          .filter(_ <= last)
          .flatMap(r =>
            Vector(helperWin, helperSpend, helperCpm).map(c => s"Loads!${colLetter(c)}$r")
          )
      else Vector.empty
    val samples =
      Vector(0, summaryRows / 2).distinct.filter(_ < summaryRows).flatMap(summarySamples) ++
        helperSamples ++
        Vector(0, shape.rh / 3, 2 * shape.rh / 3, shape.rh - 1).distinct
          .filter(_ >= 0)
          .flatMap(rhSamples) ++
        Vector(summaryRows - 1).filter(_ > summaryRows / 2).flatMap(summarySamples)
    val summaryFormulas = math.max(0, shape.aggs)
    Built(
      shape,
      workbook,
      dataFormulas = if shape.helpers then 3 * shape.rows else 0,
      summaryFormulas = summaryFormulas,
      rhFormulas = shape.rh * rhCols,
      sampleRefs = samples.distinct.take(40)
    )

  /**
   * Running totals over a formula column: `B` constants, `C_n = B_n*2` and `D_n = SUM($C$2:C_n)` on
   * sheet `Run` for rows 2..rows+1. The legacy graph gives `D_n` one edge per formula above it
   * (quadratic); `D_n` is `2 * (sum of B_2..B_n)` with `B_n = n mod 7 + 1`.
   */
  def runTotal(rows: Int): Workbook =
    val puts = (0 until rows).iterator.flatMap { i =>
      val r = FirstRow + i
      Vector((1, r, int(r % 7 + 1)), (2, r, fx(s"B$r*2")), (3, r, fx(s"SUM($$C$$2:C$r)")))
    }
    Workbook(Vector(sheet("Run", puts)))

  /** The value `runTotal(rows)` computes for `D_r` (1-based row `r`). */
  def runTotalAt(r: Int): BigDecimal =
    BigDecimal((FirstRow to r).map(n => 2L * (n % 7 + 1)).sum)
