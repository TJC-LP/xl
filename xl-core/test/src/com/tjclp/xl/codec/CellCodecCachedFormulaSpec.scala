package com.tjclp.xl.codec

import java.time.{LocalDate, LocalDateTime}

import com.tjclp.xl.Generators
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellError, CellValue, FormulaKind}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.codec.syntax.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.richtext.RichText
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.*

/**
 * GH-477: typed reads see a formula cell's cached value.
 *
 * A recalculated or Excel-authored book stores `Formula(expr, Some(v), kind)`; every `CellCodec`
 * read (and therefore `readTyped` / `readTypedOr` / `readTypedOpt`) must decode `v` exactly as it
 * would decode a plain cell holding `v`. An UNCACHED formula (`Formula(expr, None, kind)`) has
 * nothing to read and stays a `TypeMismatch` whose `actual` is the formula itself. The strict
 * escape hatch (`readStrict` / `readTypedStrict`) keeps the pre-0.20 "any formula is a mismatch"
 * semantics.
 */
class CellCodecCachedFormulaSpec extends ScalaCheckSuite:

  private val at = ref"B1"

  private def cached(v: CellValue, expr: String = "A1*3"): Cell =
    Cell(at, CellValue.Formula(expr, Some(v), FormulaKind.Normal()))

  private val uncached: Cell = Cell(at, CellValue.Formula("A1*3", None, FormulaKind.Normal()))

  private val dt = LocalDateTime.of(2025, 1, 15, 10, 30)
  private val rich = RichText.plain("styled")

  // ========== Cached formula reads as the cached value, one test per codec ==========

  test("GH-477: String codec reads a cached Text / Number / Bool result"):
    assertEquals(CellCodec[String].read(cached(CellValue.Text("hello"))), Right(Some("hello")))
    assertEquals(CellCodec[String].read(cached(CellValue.Number(BigDecimal(6)))), Right(Some("6")))
    assertEquals(CellCodec[String].read(cached(CellValue.Bool(true))), Right(Some("true")))

  test("GH-477: Int codec reads a cached Number result"):
    assertEquals(CellCodec[Int].read(cached(CellValue.Number(BigDecimal(6)))), Right(Some(6)))

  test("GH-477: Long codec reads a cached Number result"):
    assertEquals(
      CellCodec[Long].read(cached(CellValue.Number(BigDecimal(9876543210L)))),
      Right(Some(9876543210L))
    )

  test("GH-477: Double codec reads a cached Number result (the issue's exact probe)"):
    assertEquals(
      CellCodec[Double].read(cached(CellValue.Number(BigDecimal("6.0")))),
      Right(Some(6.0))
    )

  test("GH-477: BigDecimal codec reads a cached Number result"):
    assertEquals(
      CellCodec[BigDecimal].read(cached(CellValue.Number(BigDecimal(6)))),
      Right(Some(BigDecimal(6)))
    )

  test("GH-477: Boolean codec reads a cached Bool result and a cached 0/1"):
    assertEquals(CellCodec[Boolean].read(cached(CellValue.Bool(true))), Right(Some(true)))
    assertEquals(
      CellCodec[Boolean].read(cached(CellValue.Number(BigDecimal(0)))),
      Right(Some(false))
    )

  test("GH-477: LocalDate codec reads a cached DateTime result and a cached serial"):
    assertEquals(
      CellCodec[LocalDate].read(cached(CellValue.DateTime(dt))),
      Right(Some(LocalDate.of(2025, 1, 15)))
    )
    // 45971 = 2025-11-10 in the 1900 date system (mirrors CellCodecSpec)
    assertEquals(
      CellCodec[LocalDate].read(cached(CellValue.Number(BigDecimal(45971)))),
      Right(Some(LocalDate.of(2025, 11, 10)))
    )

  test("GH-477: LocalDateTime codec reads a cached DateTime result"):
    assertEquals(CellCodec[LocalDateTime].read(cached(CellValue.DateTime(dt))), Right(Some(dt)))

  test("GH-477: RichText codec reads a cached RichText result and a cached plain Text"):
    assertEquals(CellCodec[RichText].read(cached(CellValue.RichText(rich))), Right(Some(rich)))
    assertEquals(
      CellCodec[RichText].read(cached(CellValue.Text("plain"))),
      Right(Some(RichText.plain("plain")))
    )

  // ========== Cached results that still do not decode ==========

  test("GH-477: a cached result of the wrong shape is a TypeMismatch naming the CACHED value"):
    assertEquals(
      CellCodec[Int].read(cached(CellValue.Text("n/a"))),
      Left(CodecError.TypeMismatch("Int", CellValue.Text("n/a")))
    )
    assertEquals(
      CellCodec[BigDecimal].read(cached(CellValue.Error(CellError.Div0))),
      Left(CodecError.TypeMismatch("BigDecimal", CellValue.Error(CellError.Div0)))
    )

  test("GH-477: a cached Empty result reads as an empty cell (Right(None))"):
    assertEquals(CellCodec[Int].read(cached(CellValue.Empty)), Right(None))
    assertEquals(CellCodec[String].read(cached(CellValue.Empty)), Right(None))

  test("GH-477: a cached Number that does not fit Int is still a ParseError, not a mismatch"):
    CellCodec[Int].read(cached(CellValue.Number(BigDecimal("6.5")))) match
      case Left(CodecError.ParseError(_, "Int", _)) => ()
      case other => fail(s"expected ParseError, got $other")

  // ========== Uncached formula: nothing to read ==========

  test("GH-477: an UNCACHED formula is a TypeMismatch whose actual is the formula, every codec"):
    def mismatch[A: CellCodec](expected: String): Unit =
      assertEquals(
        CellCodec[A].read(uncached),
        Left(CodecError.TypeMismatch(expected, uncached.value)),
        s"codec for $expected"
      )
    mismatch[String]("String")
    mismatch[Int]("Int")
    mismatch[Long]("Long")
    mismatch[Double]("Double")
    mismatch[BigDecimal]("BigDecimal")
    mismatch[Boolean]("Boolean")
    mismatch[LocalDate]("LocalDate")
    mismatch[LocalDateTime]("LocalDateTime")
    mismatch[RichText]("RichText")

  // ========== Sheet-level helpers ==========

  private val sheet: Sheet = Sheet("Calc")
    .put(ref"A1", CellValue.Number(BigDecimal(2)))
    .put(ref"B1", CellValue.Formula("A1*3", Some(CellValue.Number(BigDecimal(6)))))
    .put(ref"C1", CellValue.Formula("A1*4", None))
    .put(ref"D1", CellValue.Formula("\"x\"&\"y\"", Some(CellValue.Text("xy"))))

  test("GH-477: readTyped on a cached formula is Right(Some(cached))"):
    assertEquals(sheet.readTyped[BigDecimal](ref"B1"), Right(Some(BigDecimal(6))))
    assertEquals(sheet.readTyped[Double](ref"B1"), Right(Some(6.0)))
    assertEquals(sheet.readTyped[String](ref"D1"), Right(Some("xy")))

  test("GH-477: readTypedOpt on a cached formula is Some(cached)"):
    assertEquals(sheet.readTypedOpt[BigDecimal](ref"B1"), Some(BigDecimal(6)))
    assertEquals(sheet.readTypedOpt[Int](ref"B1"), Some(6))

  test(
    "GH-477: readTypedOr on a cached formula is the cached value, on an uncached one the default"
  ):
    assertEquals(sheet.readTypedOr[Int](ref"B1", -1), 6)
    assertEquals(sheet.readTypedOr[Int](ref"C1", -1), -1)

  test("GH-477: readTyped / readTypedOpt on an uncached formula stay TypeMismatch / None"):
    assertEquals(
      sheet.readTyped[Int](ref"C1"),
      Left(CodecError.TypeMismatch("Int", CellValue.Formula("A1*4", None)))
    )
    assertEquals(sheet.readTypedOpt[Int](ref"C1"), None)

  // ========== Law: read(Formula(e, Some(v), k)) == read(Cell(v)) for every non-formula v ==========

  /**
   * Every non-formula CellValue shape a cache may hold
   * (Text/Number/Bool/Empty/Error/DateTime/Rich).
   */
  private val genCacheable: Gen[CellValue] = Gen.frequency(
    5 -> Generators.genCellValue,
    1 -> Generators.genExcelDateTime.map(CellValue.DateTime.apply),
    1 -> Generators.genRichTextCellValue
  )

  /** (cached value, formula cell carrying it) across every FormulaKind, incl. data tables. */
  private val genCachedFormula: Gen[(CellValue, CellValue.Formula)] =
    for
      v <- genCacheable
      kind <- Generators.genFormulaKind
      expr <- Generators.genFormulaExpr
    yield kind match
      case dt: FormulaKind.DataTable => (v, CellValue.dataTable(dt, Some(v)))
      case other => (v, CellValue.Formula(expr, Some(v), other))

  private def readsAlike[A: CellCodec](v: CellValue, f: CellValue.Formula): Prop =
    CellCodec[A].read(Cell(at, f)) ?= CellCodec[A].read(Cell(at, v))

  property("GH-477: reading a cached formula equals reading its cached value, every codec/kind"):
    forAll(genCachedFormula) { case (v, f) =>
      readsAlike[String](v, f) && readsAlike[Int](v, f) && readsAlike[Long](v, f) &&
      readsAlike[Double](v, f) && readsAlike[BigDecimal](v, f) && readsAlike[Boolean](v, f) &&
      readsAlike[LocalDate](v, f) && readsAlike[LocalDateTime](v, f) && readsAlike[RichText](v, f)
    }

  // ========== Strict escape hatch: pre-GH-477 semantics, ANY formula is a mismatch ==========

  test(
    "GH-477: readStrict rejects a CACHED formula with TypeMismatch(expected, formula), every codec"
  ):
    def strictMismatch[A: CellCodec](expected: String, cell: Cell): Unit =
      assertEquals(
        CellCodec[A].readStrict(cell),
        Left(CodecError.TypeMismatch(expected, cell.value)),
        s"codec for $expected"
      )
    strictMismatch[String]("String", cached(CellValue.Text("hello")))
    strictMismatch[Int]("Int", cached(CellValue.Number(BigDecimal(6))))
    strictMismatch[Long]("Long", cached(CellValue.Number(BigDecimal(6))))
    strictMismatch[Double]("Double", cached(CellValue.Number(BigDecimal(6))))
    strictMismatch[BigDecimal]("BigDecimal", cached(CellValue.Number(BigDecimal(6))))
    strictMismatch[Boolean]("Boolean", cached(CellValue.Bool(true)))
    strictMismatch[LocalDate]("LocalDate", cached(CellValue.DateTime(dt)))
    strictMismatch[LocalDateTime]("LocalDateTime", cached(CellValue.DateTime(dt)))
    strictMismatch[RichText]("RichText", cached(CellValue.RichText(rich)))
    strictMismatch[Int]("Int", uncached)

  private def strictAlike[A: CellCodec](cell: Cell): Prop =
    CellCodec[A].readStrict(cell) ?= CellCodec[A].read(cell)

  property("GH-477: readStrict equals read on every non-formula cell, every codec"):
    forAll(genCacheable) { v =>
      val cell = Cell(at, v)
      strictAlike[String](cell) && strictAlike[Int](cell) && strictAlike[Long](cell) &&
      strictAlike[Double](cell) && strictAlike[BigDecimal](cell) && strictAlike[Boolean](cell) &&
      strictAlike[LocalDate](cell) && strictAlike[LocalDateTime](cell) && strictAlike[RichText](
        cell
      )
    }

  test("GH-477: readTypedStrict rejects cached and uncached formulas alike, reads plain cells"):
    assertEquals(
      sheet.readTypedStrict[BigDecimal](ref"B1"),
      Left(
        CodecError.TypeMismatch(
          "BigDecimal",
          CellValue.Formula("A1*3", Some(CellValue.Number(BigDecimal(6))))
        )
      )
    )
    assertEquals(
      sheet.readTypedStrict[Int](ref"C1"),
      Left(CodecError.TypeMismatch("Int", CellValue.Formula("A1*4", None)))
    )
    assertEquals(sheet.readTypedStrict[Int](ref"A1"), Right(Some(2)))
    assertEquals(sheet.readTypedStrict[Int](ref"Z9"), Right(None))

  test("GH-477: CellReader.readStrict defaults to read, so external readers keep their semantics"):
    val external = new CellReader[Int]:
      def read(cell: Cell): Either[CodecError, Option[Int]] =
        Right(Some(cell.effectiveValue.toString.length))
    val cell = cached(CellValue.Number(BigDecimal(6)))
    assertEquals(external.readStrict(cell), external.read(cell))

  // ========== Cell.effectiveValue / isUncachedFormula ==========

  test(
    "GH-477: Cell.effectiveValue is the cached value of a cached formula, else the value itself"
  ):
    assertEquals(
      cached(CellValue.Number(BigDecimal(6))).effectiveValue,
      CellValue.Number(BigDecimal(6))
    )
    assertEquals(uncached.effectiveValue, uncached.value)
    assertEquals(Cell(at, CellValue.Text("x")).effectiveValue, CellValue.Text("x"))
    assertEquals(Cell.empty(at).effectiveValue, CellValue.Empty)

  property("GH-477: effectiveValue of a cached formula is its cache (never a Formula), any kind"):
    forAll(genCachedFormula) { case (v, f) =>
      val cell = Cell(at, f)
      (cell.effectiveValue ?= v) && !cell.isUncachedFormula && cell.isFormula
    }

  test("GH-477: isUncachedFormula holds only for Formula(_, None, _)"):
    assert(uncached.isUncachedFormula)
    assert(!cached(CellValue.Number(BigDecimal(6))).isUncachedFormula)
    assert(!Cell(at, CellValue.Number(BigDecimal(6))).isUncachedFormula)
    assert(!Cell.empty(at).isUncachedFormula)
