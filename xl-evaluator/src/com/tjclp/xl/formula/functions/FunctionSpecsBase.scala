package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue, BindingCoercion}
import com.tjclp.xl.formula.eval.{
  EvalError,
  Evaluator,
  ArrayArithmetic,
  ArrayResult,
  RangeOperand,
  ScalarCoercion
}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{CellError, CellValue}
import java.time.LocalDate

trait FunctionSpecsBase:
  protected given numericExpr: ArgSpec[TExpr[BigDecimal]] = ArgSpec.expr[BigDecimal]
  protected given stringExpr: ArgSpec[TExpr[String]] = ArgSpec.expr[String]
  protected given intExpr: ArgSpec[TExpr[Int]] = ArgSpec.expr[Int]
  protected given booleanExpr: ArgSpec[TExpr[Boolean]] = ArgSpec.expr[Boolean]
  protected given cellValueExpr: ArgSpec[TExpr[CellValue]] = ArgSpec.expr[CellValue]
  protected given dateExpr: ArgSpec[TExpr[LocalDate]] = ArgSpec.expr[LocalDate]
  protected given rangeLocation: ArgSpec[TExpr.RangeLocation] = ArgSpec.rangeLocation
  @deprecated(
    "Use rangeLocation (ArgSpec[TExpr.RangeLocation]) instead — see ArgSpec.cellRange",
    "0.18.0"
  )
  @annotation.nowarn("cat=deprecation") // forwarder must keep referencing the deprecated given
  protected given cellRange: ArgSpec[CellRange] = ArgSpec.cellRange

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  protected given anyExpr: ArgSpec[TExpr[Any]] with
    def describeParts: List[String] = List("value")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (TExpr[Any], List[TExpr[?]])] =
      args match
        // GH-603/GH-654: an omitted argument in a value position is Excel's blank, read as 0 on
        // EVERY evaluation route — `IF(TRUE,,5)` is 0 and so is the TRUE branch of the spilled
        // `IF(A1:A2>1,,5)` — so the slot coerces here rather than in one evaluator entry point.
        // The printer sees through Coerced, so the slot still prints empty.
        case TExpr.Missing :: tail =>
          Right(
            (
              TExpr.Coerced[Any](TExpr.Missing.asInstanceOf[TExpr[Any]], BindingCoercion.Numeric),
              tail
            )
          )
        case head :: tail => Right((head.asInstanceOf[TExpr[Any]], tail))
        case Nil =>
          Left(ParseError.InvalidArguments(fnName, pos, describe, "0 arguments"))

    def toValues(args: TExpr[Any]): List[ArgValue] =
      List(ArgValue.Expr(args))

    // a value slot takes any element: offered to the flagged functions (criteria, MATCH/TEXT
    // values); IF branches, CHOOSE and OFFSET's anchor are Any slots too, but those functions
    // carry their own array semantics and are never flagged
    override def scalarSlots(args: TExpr[Any]): List[(TExpr[?], LiftSlot)] =
      List((args, LiftSlot.Value))

    override def replaceScalarSlots(
      args: TExpr[Any],
      replacements: List[TExpr[?]]
    ): (TExpr[Any], List[TExpr[?]]) =
      replacements match
        case head :: rest => (head.asInstanceOf[TExpr[Any]], rest)
        case Nil => (args, Nil)

    def map(
      args: TExpr[Any]
    )(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: CellRange => CellRange
    ): TExpr[Any] =
      mapExpr(args).asInstanceOf[TExpr[Any]]

  type UnaryNumeric = TExpr[BigDecimal]
  type BinaryNumeric = (TExpr[BigDecimal], TExpr[BigDecimal])
  type BinaryNumericOpt = (TExpr[BigDecimal], Option[TExpr[BigDecimal]])
  type UnaryText = TExpr[String]
  type BinaryTextInt = (TExpr[String], TExpr[Int])
  type TextIntInt = (TExpr[String], TExpr[Int], TExpr[Int])
  type FindArgs = (TExpr[String], TExpr[String], Option[TExpr[Int]])
  type SubstituteArgs = (TExpr[String], TExpr[String], TExpr[String], Option[TExpr[Int]])
  type TextArgs = (TExpr[Any], TExpr[String])
  type TextList = List[TExpr[String]]
  type UnaryBoolean = TExpr[Boolean]
  type BooleanList = List[TExpr[Boolean]]
  type UnaryCellValue = TExpr[CellValue]
  type UnaryRange = TExpr.RangeLocation

  // Variadic numeric: Either a range (aggregated) or a single numeric expression
  // Used for Excel-compatible SUM(1,2,3) / SUM(A1:A5,B1:B5) / SUM(A1,5,B1:B3)
  type NumericArg = Either[TExpr.RangeLocation, TExpr[BigDecimal]]
  type VariadicNumeric = List[NumericArg]
  type RangeCriteriaList = List[(TExpr.RangeLocation, TExpr[Any])]
  type SumIfArgs = (TExpr.RangeLocation, TExpr[Any], Option[TExpr.RangeLocation])
  type CountIfArgs = (TExpr.RangeLocation, TExpr[Any])
  type SumIfsArgs = (TExpr.RangeLocation, RangeCriteriaList)
  type CountIfsArgs = RangeCriteriaList
  type AverageIfArgs = (TExpr.RangeLocation, TExpr[Any], Option[TExpr.RangeLocation])
  type AverageIfsArgs = (TExpr.RangeLocation, RangeCriteriaList)
  type DateInt = (TExpr[LocalDate], TExpr[Int])
  type DatePairUnit = (TExpr[LocalDate], TExpr[LocalDate], TExpr[String])
  // GH-394: holiday args are RangeLocations (sheet-qualified ranges and defined names resolve
  // at evaluation through Evaluator.resolveRangeLocation, like every range-typed slot)
  type DatePairOptRange = (TExpr[LocalDate], TExpr[LocalDate], Option[TExpr.RangeLocation])
  type DateIntOptRange = (TExpr[LocalDate], TExpr[Int], Option[TExpr.RangeLocation])
  type DatePairOptBasis = (TExpr[LocalDate], TExpr[LocalDate], Option[TExpr[Int]])
  type IfArgs = (TExpr[Boolean], TExpr[Any], TExpr[Any])
  // GH-120 statistical functions over a range
  type RangeIntArgs = (TExpr.RangeLocation, TExpr[Int])
  type RankArgs = (TExpr[BigDecimal], TExpr.RangeLocation, Option[TExpr[Int]])
  type RangeNumArgs = (TExpr.RangeLocation, TExpr[BigDecimal])
  // GH-76 dynamic arrays (spill engine)
  type SequenceArgs =
    (TExpr[Int], Option[TExpr[Int]], Option[TExpr[BigDecimal]], Option[TExpr[BigDecimal]])
  type SortArgs = (TExpr.RangeLocation, Option[TExpr[Int]], Option[TExpr[Int]])
  type UniqueArgs = (TExpr.RangeLocation, Option[TExpr[Boolean]], Option[TExpr[Boolean]])
  // GH-580: `include` is a range OR an array-valued expression (B1:B3>1, (A1:A3="x")*(B1:B3>0))
  type FilterArgs = (TExpr.RangeLocation, ArgSpec.SumProductArg, Option[TExpr[Any]])
  // GH-122 OFFSET: anchor ref + row/col offsets + optional height/width
  type OffsetArgs = (AnyExpr, TExpr[Int], TExpr[Int], Option[TExpr[Int]], Option[TExpr[Int]])
  // GH-274 INDIRECT: ref_text + optional a1 flag (FALSE = R1C1, documented-unsupported)
  type IndirectArgs = (TExpr[String], Option[TExpr[Boolean]])
  type IfErrorArgs = (TExpr[CellValue], TExpr[CellValue])
  type NoArgs = EmptyTuple
  type DateTripleInt = (TExpr[Int], TExpr[Int], TExpr[Int])
  type UnaryDate = TExpr[LocalDate]
  type AnyExpr = TExpr[Any]
  // GH-394: cashflow/date ranges are RangeLocations — `=XIRR(S!A1:B1, S!A2:B2)` and named
  // ranges resolve at evaluation; the CellRange-taking public builders wrap RangeLocation.Local
  type NpvArgs = (TExpr[BigDecimal], TExpr.RangeLocation)
  type IrrArgs = (TExpr.RangeLocation, Option[TExpr[BigDecimal]])
  type VlookupArgs = (TExpr[CellValue], TExpr.RangeLocation, TExpr[Int], Option[TExpr[Boolean]])
  // GH-197: Changed to accept both ranges AND array expressions
  type SumProductArgs = List[ArgSpec.SumProductArg]
  type XLookupArgs = (
    AnyExpr,
    TExpr.RangeLocation,
    TExpr.RangeLocation,
    Option[AnyExpr],
    Option[TExpr[Int]],
    Option[TExpr[Int]]
  )
  // GH-669: the array is a location, a union/intersection (area_num picks an area) or an array
  // constant; the fourth slot is area_num
  type IndexArgs = (
    ReferenceOperators.Operand,
    TExpr[BigDecimal],
    Option[TExpr[BigDecimal]],
    Option[TExpr[BigDecimal]]
  )
  type MatchArgs = (AnyExpr, TExpr.RangeLocation, Option[TExpr[BigDecimal]])
  type AddressArgs = (
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    Option[TExpr[BigDecimal]],
    Option[TExpr[Boolean]],
    Option[TExpr[String]]
  )
  // GH-424 CELL: info_type + optional positional reference (never evaluated — its address is
  // the datum)
  type CellArgs = (TExpr[String], Option[AnyExpr])
  // GH-476 HYPERLINK: link_location + optional friendly_name (the displayed text)
  type HyperlinkArgs = (TExpr[String], Option[TExpr[String]])
  type XnpvArgs = (TExpr[BigDecimal], TExpr.RangeLocation, TExpr.RangeLocation)
  type XirrArgs = (TExpr.RangeLocation, TExpr.RangeLocation, Option[TExpr[BigDecimal]])
  type TvmArgs = (
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    Option[TExpr[BigDecimal]],
    Option[TExpr[BigDecimal]]
  )
  type RateArgs = (
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    TExpr[BigDecimal],
    Option[TExpr[BigDecimal]],
    Option[TExpr[BigDecimal]],
    Option[TExpr[BigDecimal]]
  )
  // GH-605 RRI: nper, pv, fv — three required numbers
  type RriArgs = (TExpr[BigDecimal], TExpr[BigDecimal], TExpr[BigDecimal])

  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  protected def evalAny(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, Any] =
    // Resolve PolyRef/SheetPolyRef to typed Ref before evaluation.
    // This fixes cell references used as criteria in SUMIFS, COUNTIF, etc.
    // GH-476: asResolvedValueExpr pushes through the transparent UnaryPlus wrapper too, so the
    // house `=+Sheet!Ref` style resolves in Any-typed argument positions (=IF(1=1,+S2!G1,0)).
    val resolved = expr match
      case _: TExpr.PolyRef | _: TExpr.SheetPolyRef | _: TExpr.UnaryPlus[?] =>
        TExpr.asResolvedValueExpr(expr)
      case other => other
    ctx.evalExpr[Any](resolved.asInstanceOf[TExpr[Any]])

  /**
   * GH-333: evaluate an argument array-aware, materializing bare ranges.
   *
   * Unlike evalAny (whose scalar boundary collapses arrays and, in a plain cell, intersects bare
   * ranges), this yields ArrayResults for range-shaped arguments and array-producing expressions —
   * the IF branch/condition convention in array mode, mirroring the evaluator's own evalMaybeArray
   * operand handling.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  protected def evalMaybeArrayArg(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, Any] =
    expr match
      case TExpr.RangeRef(range, _) =>
        extractRangeAsMatrixEval(range, ctx.sheet, ctx).map(ArrayResult(_))
      case TExpr.SheetRange(sheetName, range, _) =>
        Evaluator
          .resolveRangeLocation(
            TExpr.RangeLocation.CrossSheet(sheetName, range),
            ctx.sheet,
            ctx.workbook
          )
          .flatMap { case (targetSheet, _) =>
            extractRangeAsMatrixEval(range, targetSheet, ctx).map(ArrayResult(_))
          }
      // GH-476: the house `=+Sheet!Ref` style resolves here too (asResolvedValueExpr pushes
      // through the transparent UnaryPlus), as it does for evalAny
      case _: TExpr.PolyRef | _: TExpr.SheetPolyRef | _: TExpr.UnaryPlus[?] =>
        ctx.evalArrayExpr(TExpr.asResolvedValueExpr(expr).asInstanceOf[TExpr[Any]])
      case other =>
        ctx.evalArrayExpr(other.asInstanceOf[TExpr[Any]])

  /**
   * The IF / IFS condition and NOT's operand, a value position: array-aware in array mode (it
   * broadcasts), a value in a plain cell — its references implicitly intersected with the formula's
   * cell, an array value read at its top-left, as Excel computes a legacy formula.
   */
  protected def evalCondition(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, Any] =
    if ctx.arrayMode then evalMaybeArrayArg(ctx, expr) else evalAny(ctx, expr)

  /**
   * The value IF, IFS, CHOOSE or SWITCH selects: the reference itself when the call feeds a
   * reference position (Excel's IF and CHOOSE return references, so `SUM(IF(A1:A10>2,A1:A10,0))`
   * sums the whole range, and `SUM(IF(TRUE,C1,0))` skips a text C1 as a range would); array-aware
   * in array mode; otherwise a value.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  protected def evalSelected(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, Any] =
    if ctx.selectsReference then ctx.evalReference(expr.asInstanceOf[TExpr[Any]])
    else if ctx.arrayMode then evalMaybeArrayArg(ctx, expr)
    else evalAny(ctx, expr)

  /**
   * What a reference-returning function (OFFSET, INDIRECT, INDEX) evaluates to for the reference it
   * computed: `whole` — the range's values — in array mode; in a plain cell's value position, the
   * implicitly intersected cell (`=INDIRECT("A1:A10")*2` in row 5 reads A5), `#VALUE!` where the
   * formula's row or column does not cross the range. A reference position asks the function for
   * its reference instead ([[FunctionSpec.reference]]).
   */
  protected def referenceResult(
    range: CellRange,
    target: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext
  )(
    whole: => Either[EvalError, ArrayResult]
  ): Either[EvalError, ArrayResult] =
    if ctx.arrayMode || ctx.selectsReference then whole
    else
      Evaluator
        .implicitIntersection(range, ctx.currentCell)
        .flatMap(at => extractRangeAsMatrixEval(CellRange(at, at), target, ctx).map(ArrayResult(_)))

  /**
   * A reference-returning function's value from what its [[FunctionSpec.reference]] computed: the
   * value it computed instead of a reference (a `#REF!`) as is, a reference by [[referenceResult]]
   * with `whole` materializing it.
   */
  protected def referencedValue(located: Either[ArrayResult, RangeOperand], ctx: EvalContext)(
    whole: RangeOperand => Either[EvalError, ArrayResult]
  ): Either[EvalError, ArrayResult] =
    located match
      case Left(value) => Right(value)
      case Right(ref) => referenceResult(ref.range, ref.sheet, ctx)(whole(ref))

  /**
   * A reference IF, IFS, CHOOSE or SWITCH selected, read as the array a broadcast needs: IFS's
   * remaining pairs under an array condition, when the IFS feeds a reference position in array mode
   * (the `@` operand).
   */
  protected def materializeOperand(ctx: EvalContext, value: Any): Either[EvalError, Any] =
    value match
      case RangeOperand(target, range) =>
        extractRangeAsMatrixEval(range, target, ctx).map(ArrayResult(_))
      case other => Right(other)

  /** A range reference, seen through the coercion a typed slot wraps it in. */
  protected def bareRange(expr: TExpr[?]): Option[TExpr[?]] = expr match
    case r @ (_: TExpr.RangeRef | _: TExpr.SheetRange) => Some(r)
    case TExpr.Coerced(inner, _) => bareRange(inner)
    case _ => None

  /**
   * GH-603/GH-654: whether an argument slot was left EMPTY in the source (`SORT(rng,,-1)`): the
   * parser's `TExpr.Missing`, possibly under the coercion a typed slot wrapped it in.
   */
  protected def isEmptySlot(expr: TExpr[?]): Boolean = expr match
    case TExpr.Missing => true
    case TExpr.Coerced(inner, _) => isEmptySlot(inner)
    case _ => false

  /**
   * GH-654: the dynamic-array and reference functions read an empty optional slot as OMITTED —
   * Excel tests the argument's presence there, not its value — so `OFFSET(A1,0,0,,2)` keeps the
   * anchor's height, `SEQUENCE(3,,5)` is 5,6,7, `SORT(rng,,-1)` sorts by its first column and
   * `XLOOKUP(x,a,b,,0)` is `#N/A` when nothing matches (an explicit 0 in each of those slots is an
   * error). Classic functions read the same slot as a blank VALUE (see `ArgSpec.option`):
   * `VLOOKUP(x,rng,2,)` is an exact match, `MATCH(x,rng,)` too, `LOG(10,)` has base 0. Both rules
   * verified against LibreOffice's recalculation.
   */
  protected def unlessOmitted[A](arg: Option[TExpr[A]]): Option[TExpr[A]] =
    arg.filterNot(isEmptySlot)

  protected def rangeCellReader(
    targetSheet: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext
  ): ARef => Either[EvalError, CellValue] =
    Evaluator.cellValueReader(
      targetSheet,
      ctx.clock,
      ctx.workbook,
      ctx.depth,
      ctx.rng,
      ctx.memo.getOrElse(new Evaluator.EvalMemo),
      ctx.workbookPath,
      ctx.aggregateMemo
    )

  protected def extractRangeAsMatrixEval(
    range: CellRange,
    targetSheet: com.tjclp.xl.sheets.Sheet,
    ctx: EvalContext
  ): Either[EvalError, Vector[Vector[CellValue]]] =
    ArrayArithmetic.rangeToArrayEval(range, rangeCellReader(targetSheet, ctx)).map(_.values)

  /** Normalize an evaluated value to an ArrayResult (scalars become 1x1). */
  protected def toCellArray(value: Any): ArrayResult = value match
    case arr: ArrayResult => arr
    case scalar => ArrayResult.single(ArrayArithmetic.anyToCellValue(scalar))

  protected def evalValue(ctx: EvalContext, expr: TExpr[?]): Either[EvalError, ExprValue] =
    evalAny(ctx, expr).map(ExprValue.from)

  protected def toCellValue(value: ExprValue): CellValue =
    value match
      case ExprValue.Cell(cv) => cv
      case ExprValue.Text(s) => CellValue.Text(s)
      case ExprValue.Number(n) => CellValue.Number(n)
      case ExprValue.Bool(b) => CellValue.Bool(b)
      case ExprValue.Date(d) => CellValue.DateTime(d.atStartOfDay())
      case ExprValue.DateTime(dt) => CellValue.DateTime(dt)
      case ExprValue.Opaque(other) => CellValue.Text(other.toString)

  /**
   * GH-307: total integer-argument extraction for ExprValue-evaluated arguments (EDATE/EOMONTH
   * months, WORKDAY days, YEARFRAC basis, XLOOKUP modes).
   *
   * Replaces the silent `toInt` (non-numeric → 0 → garbage results). Conventions come from the
   * shared ScalarCoercion Integer table: fractionals TRUNCATE toward zero (Excel truncates
   * months/days), numeric text parses ("3" → 3), booleans are TRUE = 1 / FALSE = 0, anything else
   * is a clean per-cell error naming the function (Excel: #VALUE!).
   */
  protected def toIntArg(fnName: String, value: ExprValue): Either[EvalError, Int] =
    val raw = value match
      case ExprValue.Number(n) => n
      case ExprValue.Text(s) => s
      case ExprValue.Bool(b) => b
      case ExprValue.Date(d) => d
      case ExprValue.DateTime(dt) => dt
      case ExprValue.Cell(cv) => cv
      case ExprValue.Opaque(other) => other
    ScalarCoercion.coerce(s"$fnName integer argument", raw, BindingCoercion.Integer).flatMap {
      case i: Int => Right(i)
      case other =>
        Left(EvalError.TypeMismatch(s"$fnName integer argument", "integer", s"$other"))
    }

  protected def coerceToNumeric(value: CellValue): BigDecimal =
    value match
      case CellValue.Number(n) => n
      case CellValue.Bool(b) => ArrayArithmetic.boolToNumeric(b)
      case CellValue.DateTime(dt) => BigDecimal(CellValue.dateTimeToExcelSerial(dt))
      case CellValue.Formula(_, Some(cached), _) => coerceToNumeric(cached)
      case _ => BigDecimal(0)

  /**
   * Extract numeric value from CellValue, handling formulas with cached results.
   *
   * GH-449: a DateTime cell yields its Excel serial — dates ARE numbers in Excel's aggregate
   * positions (=MIN/MAX over a date column are the earliest/latest date, =COUNT counts them,
   * =SUM sums the serials), mirroring TExprDecoders.decodeNumeric. Booleans stay skipped: Excel's
   * range aggregates ignore logicals, which is why this is not simply coerceToNumeric.
   */
  protected def extractNumericValue(value: CellValue): Option[BigDecimal] =
    value match
      case CellValue.Number(n) => Some(n)
      case CellValue.DateTime(dt) => Some(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case CellValue.Formula(_, Some(cached), _) => extractNumericValue(cached)
      case _ => None

  /** Extract the anchor ARef from a reference expression (cell ref, or the start of a range). */
  protected def extractARef(expr: TExpr[?]): Option[ARef] =
    expr match
      case TExpr.PolyRef(ref, _) => Some(ref)
      case TExpr.Ref(ref, _, _) => Some(ref)
      case TExpr.SheetPolyRef(_, ref, _) => Some(ref)
      case TExpr.SheetRef(_, ref, _, _) => Some(ref)
      case TExpr.RangeRef(range, _) => Some(range.start)
      case TExpr.SheetRange(_, range, _) => Some(range.start)
      case _ => None

  /**
   * GH-662: the typed `#N/A` a lookup raises when nothing matches. Left channel per the GH-344
   * charter (RANK's not-found precedent), so IFNA/ISNA/ERROR.TYPE see the code and an unguarded
   * miss promotes to a cached `#N/A` at the CellValue boundary; the context is the human diagnostic
   * and is dropped, by design, at that boundary. LOOKUP/XMATCH must raise through this when added.
   */
  protected def lookupNotFound(context: String): EvalError =
    EvalError.ErrorValue(CellError.NA, Some(context))

  /** One rendering of a lookup value for text matching and diagnostics (VLOOKUP/HLOOKUP/MATCH). */
  protected def renderLookupValue(value: ExprValue): String = value match
    case ExprValue.Text(s) => s
    // #665: the one number → text rule, so a diagnostic never says `999.0` where `&` says `999`
    case ExprValue.Number(n) => com.tjclp.xl.formula.eval.ScalarCoercion.numberText(n)
    case ExprValue.Bool(b) => b.toString
    case ExprValue.Date(d) => d.toString
    case ExprValue.DateTime(dt) => dt.toString
    case ExprValue.Cell(cv) => cv.toString
    case ExprValue.Opaque(other) => other.toString

  /**
   * GH-467: normalize a lookup value for MATCH/XLOOKUP comparison. Cell-ref lookup values arrive as
   * ExprValue.Cell and must be dereferenced to their underlying scalar before comparing (the
   * literal `"k2"` and a ref to a cell holding `"k2"` must behave identically); date/datetime
   * lookup values coerce to their Excel serial — dates ARE numbers in Excel's comparison plane
   * (mirroring GH-449's dateTimeToExcelSerial coercion at numeric boundaries).
   */
  protected def normalizeLookupValue(value: ExprValue): ExprValue =
    value match
      case ExprValue.Cell(CellValue.Number(n)) => ExprValue.Number(n)
      case ExprValue.Cell(CellValue.Text(s)) => ExprValue.Text(s)
      case ExprValue.Cell(CellValue.Bool(b)) => ExprValue.Bool(b)
      case ExprValue.Cell(CellValue.DateTime(dt)) =>
        ExprValue.Number(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case ExprValue.Cell(CellValue.Formula(_, Some(cached), _)) =>
        normalizeLookupValue(ExprValue.Cell(cached))
      case ExprValue.Date(d) =>
        ExprValue.Number(BigDecimal(CellValue.dateTimeToExcelSerial(d.atStartOfDay())))
      case ExprValue.DateTime(dt) =>
        ExprValue.Number(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case other => other

  /**
   * GH-467: exact equality for the MATCH/XLOOKUP lookup plane, over a lookup value already passed
   * through [[normalizeLookupValue]]. Total and blank-safe: a blank cell matches nothing (a
   * default-equal fallback here made blanks "match" every lookup value, shifting MATCH positions
   * over ranges with holes), mismatched types match nothing, and DateTime cells compare by their
   * Excel serial so date lookups find date columns.
   */
  protected def matchesLookupExact(cv: CellValue, lookup: ExprValue): Boolean =
    (cv, lookup) match
      case (CellValue.Number(n), ExprValue.Number(v)) => n == v
      case (CellValue.DateTime(dt), ExprValue.Number(v)) =>
        BigDecimal(CellValue.dateTimeToExcelSerial(dt)) == v
      case (CellValue.Text(s), ExprValue.Text(v)) => s.equalsIgnoreCase(v)
      case (CellValue.Bool(b), ExprValue.Bool(v)) => b == v
      case (CellValue.Error(e1), ExprValue.Cell(CellValue.Error(e2))) => e1 == e2
      case (CellValue.Formula(_, Some(cached), _), v) => matchesLookupExact(cached, v)
      case _ => false

  /** Extract text for matching, coercing numbers and booleans to strings. */
  protected def extractTextForMatch(cv: CellValue): Option[String] =
    cv match
      case CellValue.Text(s) => Some(s)
      case CellValue.Number(n) => Some(n.bigDecimal.stripTrailingZeros().toPlainString)
      case CellValue.Bool(b) => Some(if b then "TRUE" else "FALSE")
      case CellValue.Formula(_, Some(cached), _) => extractTextForMatch(cached)
      case _ => None

  /**
   * Extract numeric for matching, parsing text as numbers.
   *
   * GH-488: a DateTime cell yields its Excel serial, so a date-typed VLOOKUP/HLOOKUP key finds a
   * date key column (dates ARE numbers in Excel's comparison plane — same judgment as
   * [[normalizeLookupValue]] and GH-449's numeric-boundary coercion).
   */
  protected def extractNumericForMatch(cv: CellValue): Option[BigDecimal] =
    cv match
      case CellValue.Number(n) => Some(n)
      case CellValue.Text(s) => scala.util.Try(BigDecimal(s.trim)).toOption
      case CellValue.Bool(b) => Some(ArrayArithmetic.boolToNumeric(b))
      case CellValue.DateTime(dt) => Some(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case CellValue.Formula(_, Some(cached), _) => extractNumericForMatch(cached)
      case _ => None

  /** Validate that two ranges have the same dimensions. */
  protected def validateDimensions(
    range1: CellRange,
    range2: CellRange,
    fnName: String,
    range1Name: String,
    range2Name: String
  ): Either[EvalError, Unit] =
    if range1.width != range2.width || range1.height != range2.height then
      Left(
        EvalError.EvalFailed(
          s"$fnName: $range1Name and $range2Name must have same dimensions " +
            s"(${range1.height}×${range1.width} vs ${range2.height}×${range2.width})",
          Some(s"$fnName(${range1.toA1}, ..., ${range2.toA1})")
        )
      )
    else Right(())
