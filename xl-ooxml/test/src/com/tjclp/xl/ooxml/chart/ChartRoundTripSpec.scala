package com.tjclp.xl.ooxml.chart

import munit.ScalaCheckSuite
import org.scalacheck.Prop
import org.scalacheck.Prop.forAll

import com.tjclp.xl.Generators
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.charts.Chart
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XmlUtil
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook

/**
 * GH-222 KEYSTONE self-coherence law, GH-407 form: `emit(parse(emit(chart))) == emit(chart)` for
 * arbitrary generated charts — the writer MATERIALIZES defaults (accent-cycle series fills, "Series
 * N" literal names) that the reader captures back into the model, so re-emission is a fixpoint.
 * This still proves the emitted dialect is a subset of the read fence mechanically (parse must
 * succeed for the law to hold), with caches/idx/order/axes/pie-dPts dropped losslessly for the XML.
 *
 * PLUS the exact identity `ChartReader.parse(emit(chart)) == Some(chart)` on the fully-explicit
 * subspace — every series carries an explicit fill AND name, so the writer materializes nothing.
 * (Pie slices: explicit `pointFills` prefixes emit as dPt fills with the accent cycle beyond them —
 * GH-418; exactness additionally needs pointFills CANONICAL, i.e. sized within the values vector
 * with no trailing accent-coincident entry, since trailing accents re-derive and strip on parse.)
 *
 * Exercised under BOTH cache regimes: bare `c:f` (referenced sheet absent from the workbook) and
 * fully-resolved caches over a sheet mixing numbers, text, booleans, blanks, dates, and
 * cached-formula values.
 *
 * Determinism: pinned ScalaCheck seed, the OoxmlGenerativeRoundTripSpec convention.
 */
class ChartRoundTripSpec extends ScalaCheckSuite:

  // The OoxmlGenerativeRoundTripSpec convention: crank locally via
  // XL_ROUNDTRIP_MIN_SUCCESS=1000 (or -Dxl.roundtrip.minSuccess=... where the runner
  // forwards JVM props).
  private val minSuccess: Int =
    sys.props
      .get("xl.roundtrip.minSuccess")
      .orElse(sys.env.get("XL_ROUNDTRIP_MIN_SUCCESS"))
      .flatMap(_.toIntOption)
      .getOrElse(200)

  override def scalaCheckTestParameters: org.scalacheck.Test.Parameters =
    super.scalaCheckTestParameters
      .withMinSuccessfulTests(minSuccess)
      .withInitialSeed(org.scalacheck.rng.Seed(20260611L))

  private def emit(chart: Chart, workbook: Workbook): String =
    XmlUtil.compact(OoxmlChart(chart, ChartCaches.resolve(workbook, chart)).toXml)

  /** GH-407 law: emit∘parse∘emit == emit (the materialized model is an emission fixpoint). */
  private def lawProp(chart: Chart, workbook: Workbook): Prop =
    val first = emit(chart, workbook)
    ChartReader.parse(first) match
      case None =>
        Prop.falsified :| s"parse(emit(chart)) == None\nchart: $chart\nxml: $first"
      case Some(parsed) =>
        val second = emit(parsed, workbook)
        Prop(second == first) :|
          s"emit(parse(emit(chart))) != emit(chart)\nchart:  $chart\nparsed: $parsed\nfirst:  $first\nsecond: $second"

  /** Exact identity on the fully-explicit subspace: nothing left for the writer to materialize. */
  private def exactProp(chart: Chart, workbook: Workbook): Prop =
    val xml = emit(chart, workbook)
    val parsed = ChartReader.parse(xml)
    Prop(parsed == Some(chart)) :|
      s"parse(emit(chart)) != Some(chart)\nchart:  $chart\nparsed: $parsed\nxml: $xml"

  /**
   * Deterministically lift a generated chart into the fully-explicit subspace: every series gets an
   * explicit name and fill (generated values kept where present, writer defaults filled in where
   * absent — so explicit fills also cover arbitrary non-accent colors).
   */
  private def explicit(chart: Chart): Chart =
    import com.tjclp.xl.charts.SeriesName
    import com.tjclp.xl.ooxml.DefaultTheme
    import com.tjclp.xl.styles.color.Color
    chart.copy(series = chart.series.zipWithIndex.map { case (s, i) =>
      s.copy(
        name = s.name.orElse(Some(SeriesName.Literal(s"Series ${i + 1}"))),
        fill = s.fill.orElse(Some(Color.Rgb(DefaultTheme.accentArgb(i))))
      )
    })

  /** Grid sheet covering the generator space with every cache-relevant value class. */
  private def materialized(name: SheetName): Sheet =
    (0 to 11).foldLeft(Sheet(name)) { (sheet, row) =>
      (0 to 7).foldLeft(sheet) { (s, col) =>
        val cellRef = com.tjclp.xl.addressing.ARef.from0(col, row)
        col match
          case 0 | 1 | 2 => s.put(Cell(cellRef, CellValue.Number(BigDecimal(row * 10 + col))))
          case 3 => s.put(Cell(cellRef, CellValue.Number(BigDecimal("1.5") * row - 7)))
          case 4 | 5 => s.put(Cell(cellRef, CellValue.Text(s"v$row$col")))
          case 6 => s // blank column: pts skipped, ptCount still counts
          case _ =>
            row % 3 match
              case 0 => s.put(Cell(cellRef, CellValue.Bool(row % 2 == 0)))
              case 1 =>
                s.put(
                  Cell(
                    cellRef,
                    CellValue.Formula("A1*2", Some(CellValue.Number(BigDecimal(row))))
                  )
                )
              case _ =>
                s.put(
                  Cell(
                    cellRef,
                    CellValue.DateTime(java.time.LocalDateTime.of(2026, 6, 1 + row, 12, 0))
                  )
                )
      }
    }

  /**
   * GH-418: stamp generated per-slice fills onto pie charts (0..cellCount arbitrary opaque colors;
   * non-pie charts pass through untouched — the writer has no dPt slot for them). Kept OUT of the
   * shared [[Generators.genSeries]] so the pinned-seed value streams of every other property using
   * genChart/genChartFrame stay undisturbed.
   */
  private def withPointFills(chart: Chart): org.scalacheck.Gen[Chart] =
    import com.tjclp.xl.charts.{ChartType, Series}
    import com.tjclp.xl.styles.color.Color
    chart.series match
      case Vector(s) if chart.chartType == ChartType.Pie =>
        for
          n <- org.scalacheck.Gen.choose(0, s.values.cellCount)
          colors <- org.scalacheck.Gen.listOfN[Color.Rgb](
            n,
            org.scalacheck.Gen.choose(0, 0xffffff).map(rgb => Color.Rgb(0xff000000 | rgb))
          )
        yield chart.copy(series = Vector(s.copy(pointFills = colors.toVector)))
      case _ => org.scalacheck.Gen.const(chart)

  /**
   * Canonicalize pie pointFills the way the reader does (GH-418): trailing entries that coincide
   * with the accent cycle re-derive on emission and strip on parse.
   */
  private def canonicalPointFills(chart: Chart): Chart =
    import com.tjclp.xl.ooxml.DefaultTheme
    import com.tjclp.xl.styles.color.Color
    chart.copy(series = chart.series.map { s =>
      val trailing = s.pointFills.zipWithIndex.reverse.takeWhile { case (c, k) =>
        c == Color.Rgb(DefaultTheme.accentArgb(k))
      }.size
      s.copy(pointFills = s.pointFills.dropRight(trailing))
    })

  private val plain = SheetName.unsafe("Data")
  private val quoted = SheetName.unsafe("Q1 'Report") // needsQuoting: space + quote + cell-shaped

  property("LAW: emit∘parse∘emit == emit — bare c:f (referenced sheet absent)") {
    forAll(Generators.genChart(plain)) { chart =>
      lawProp(chart, Workbook(Vector(Sheet(SheetName.unsafe("Other")))))
    }
  }

  property("LAW: emit∘parse∘emit == emit — resolved caches over a mixed-value sheet") {
    forAll(Generators.genChart(plain)) { chart =>
      lawProp(chart, Workbook(Vector(materialized(plain))))
    }
  }

  property("LAW: emit∘parse∘emit == emit — quoting-positive sheet name, resolved caches") {
    forAll(Generators.genChart(quoted)) { chart =>
      lawProp(chart, Workbook(Vector(materialized(quoted))))
    }
  }

  property("LAW: parse(emit(chart)) == chart on the fully-explicit subspace — bare c:f") {
    forAll(Generators.genChart(plain)) { chart =>
      exactProp(explicit(chart), Workbook(Vector(Sheet(SheetName.unsafe("Other")))))
    }
  }

  property("LAW: parse(emit(chart)) == chart on the fully-explicit subspace — resolved caches") {
    forAll(Generators.genChart(plain)) { chart =>
      exactProp(explicit(chart), Workbook(Vector(materialized(plain))))
    }
  }

  property("GH-418 LAW: emit∘parse∘emit == emit — pie pointFills, resolved caches") {
    forAll(Generators.genChart(plain).flatMap(withPointFills)) { chart =>
      lawProp(chart, Workbook(Vector(materialized(plain))))
    }
  }

  property("GH-418 LAW: parse(emit(chart)) == chart — explicit + canonical pointFills") {
    forAll(Generators.genChart(plain).flatMap(withPointFills)) { chart =>
      exactProp(canonicalPointFills(explicit(chart)), Workbook(Vector(materialized(plain))))
    }
  }

  test("emission is write-twice stable and backend-agnostic at the model level") {
    val chart = Chart(
      com.tjclp.xl.charts.ChartType.Bar(),
      Vector(
        com.tjclp.xl.charts.Series(
          com.tjclp.xl.charts.DataRef(plain, ref"B2:B5"),
          Some(com.tjclp.xl.charts.DataRef(plain, ref"A2:A5")),
          Some(com.tjclp.xl.charts.SeriesName.FromCell(plain, ref"B1"))
        )
      ),
      Some("T"),
      Some(com.tjclp.xl.charts.Legend())
    )
    val wb = Workbook(Vector(materialized(plain)))
    assertEquals(emit(chart, wb), emit(chart, wb))
  }
