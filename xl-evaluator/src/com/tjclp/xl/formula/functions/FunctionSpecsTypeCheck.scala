package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.ast.{TExpr, ExprValue}
import com.tjclp.xl.formula.eval.{
  ArrayArithmetic,
  ArrayLifting,
  ArrayResult,
  EvalError,
  Evaluator,
  NameWalk,
  ScalarCoercion,
  WorkbookAudit
}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.{Clock, Arity}

import com.tjclp.xl.cells.{CellError, CellValue}

trait FunctionSpecsTypeCheck extends FunctionSpecsBase:
  val iferror: FunctionSpec[CellValue] { type Args = IfErrorArgs } =
    FunctionSpec.arrayAware[CellValue, IfErrorArgs](
      "IFERROR",
      Arity.two,
      FunctionFlags(lift = ArrayLift.slots(0)),
      Set(1)
    )(evalArrayFallback(catchOnlyNa = false)) { (args, ctx) =>
      val (valueExpr, valueIfErrorExpr) = args
      val result = evalValue(ctx, valueExpr) match
        case Left(_) =>
          evalValue(ctx, valueIfErrorExpr).map(toCellValue)
        case Right(ExprValue.Cell(cv)) =>
          // GH-512: an error CACHED inside a formula cell counts. On any recalculated book every
          // formula cell carries a cache, so matching only a bare CellValue.Error made IFERROR
          // miss the overwhelmingly common shape. `carriedError` is the shared matcher the
          // aggregate guards already use.
          if ArrayArithmetic.carriedError(cv).isDefined then
            evalValue(ctx, valueIfErrorExpr).map(toCellValue)
          else Right(cv)
        case Right(other) =>
          Right(toCellValue(other))
      result.map(unwrapCachedValue)
    }

  /**
   * GH-511: IFNA — IFERROR's discriminating sibling, and the reason banker books prefer it over
   * IFERROR on lookups. It swallows ONLY `#N/A`, so a `#DIV/0!` or `#VALUE!` propagates instead of
   * being masked: a missing lookup key is expected, a division by zero is a bug, and one guard
   * should not hide both.
   *
   * `#N/A` reaches here on either channel — as an error VALUE from a cell or a cached formula, or
   * on the Left channel from a function that failed with it (the GH-344 posture `iserr` takes
   * below) — so both are matched.
   */
  val ifna: FunctionSpec[CellValue] { type Args = IfErrorArgs } =
    FunctionSpec.arrayAware[CellValue, IfErrorArgs](
      "IFNA",
      Arity.two,
      FunctionFlags(lift = ArrayLift.slots(0)),
      Set(1)
    )(evalArrayFallback(catchOnlyNa = true)) { (args, ctx) =>
      val (valueExpr, valueIfNaExpr) = args
      val result = evalValue(ctx, valueExpr) match
        case Left(EvalError.ErrorValue(CellError.NA, _)) =>
          evalValue(ctx, valueIfNaExpr).map(toCellValue)
        case Left(other) => Left(other)
        case Right(ExprValue.Cell(cv)) if ArrayArithmetic.carriedError(cv).contains(CellError.NA) =>
          evalValue(ctx, valueIfNaExpr).map(toCellValue)
        case Right(ExprValue.Cell(cv)) => Right(cv)
        case Right(other) => Right(toCellValue(other))
      result.map(unwrapCachedValue)
    }

  /**
   * GH-754: retain a selected fallback whole; only an array primary broadcasts against it. Evaluate
   * a fallback only when at least one primary element needs it. Scalar random fallbacks retain
   * their per-element draws; array fallbacks evaluate whole and broadcast. IFNA catches only #N/A
   * on either error channel; IFERROR retains its scalar catch-all policy.
   */
  private def evalArrayFallback(catchOnlyNa: Boolean)(
    args: IfErrorArgs,
    ctx: EvalContext
  ): Either[EvalError, Any] =
    val (primary, fallback) = args
    def catches(value: CellValue): Boolean =
      ArrayArithmetic.carriedError(value).exists(err => !catchOnlyNa || err == CellError.NA)
    def selectedFallback: Either[EvalError, Any] =
      evalMaybeArrayArg(ctx, fallback).map {
        case array: ArrayResult => array.mapPure(unwrapCachedValue)
        case scalar => toCellValue(ExprValue.from(scalar))
      }
    evalMaybeArrayArg(ctx, primary) match
      case Left(error) if !catchOnlyNa || EvalError.toCellError(error).contains(CellError.NA) =>
        selectedFallback
      case Left(error) => Left(error)
      case Right(array: ArrayResult) =>
        if !array.tally.exists((value, _) => catches(value)) then
          Right(array.mapPure(unwrapCachedValue))
        else
          // A failing fallback carries its Excel error only at selected positions. Structural
          // failures without an Excel value remain loud, as in the other array conditionals.
          def replacement: Either[EvalError, Any] = selectedFallback match
            case Left(error) =>
              ArrayLifting.elementResult(Left(error))
            case right => right
          def scalarDraw(value: Any): Boolean = value match
            case _: ArrayResult => false
            case _ => fallbackDraws(fallback, ctx)
          replacement.flatMap {
            case value if scalarDraw(value) =>
              // Keep the first draw, then evaluate only for each further selected element.
              // perCell refuses oversized dense work, as generic lifting does (#750).
              val replacements = Iterator.single(Right(value): Either[EvalError, Any]) ++
                Iterator.continually(replacement)
              ArrayArithmetic.broadcastN(Vector(array), perCell = true) { elements =>
                val original = elements.headOption.getOrElse(CellValue.Empty)
                if catches(original) then
                  replacements.nextOption().getOrElse(Right(CellValue.Empty)).map { next =>
                    unwrapCachedValue(toCellArray(next)(0, 0))
                  }
                else Right(unwrapCachedValue(original))
              }
            case value =>
              ArrayArithmetic.broadcastN(Vector(array, toCellArray(value))) { elements =>
                val original = elements.headOption.getOrElse(CellValue.Empty)
                Right(
                  unwrapCachedValue(
                    if catches(original) then elements.lift(1).getOrElse(CellValue.Empty)
                    else original
                  )
                )
              }
          }
      case Right(scalar) =>
        val value = toCellValue(ExprValue.from(scalar))
        if catches(value) then selectedFallback else Right(value)

  /** Share the generation's existing name-to-random-call analysis with generic lifting. */
  private def fallbackDraws(expr: TExpr[?], ctx: EvalContext): Boolean =
    def build(wb: com.tjclp.xl.workbooks.Workbook) =
      WorkbookAudit.namesReaching(wb, WorkbookAudit.callsWhere(_, _, SumProductTail.draws))
    lazy val reach =
      ctx.workbook.map(wb => ctx.aggregateMemo.fold(build(wb))(_.namesReachingDraw(wb, build)))
    WorkbookAudit.callsWhere(
      expr,
      (qualifier, name) =>
        reach.exists(r => r(ctx.sheet.name, qualifier, name) != NameWalk.Reach.Absent),
      SumProductTail.draws
    )

  /**
   * GH-511: NA() — the `#N/A` literal. Excel's idiom for marking a deliberate gap, and the only way
   * to author the value [[ifna]] catches.
   */
  val na: FunctionSpec[CellValue] { type Args = NoArgs } =
    FunctionSpec.simple[CellValue, NoArgs]("NA", Arity.none) { (_, _) =>
      Right(CellValue.Error(CellError.NA))
    }

  /**
   * GH-630: ERROR.TYPE(error_val) — the number Excel assigns each error value (`#NULL!` 1 … `#N/A`
   * 7, `#GETTING_DATA` 8, `#SPILL!` 9 … `#CALC!` 14, [[CellError.errorTypeNumber]]), `#N/A` for
   * anything that is not an error. The error reaches here on either channel — as a VALUE from a
   * cell, a cache or an error literal, or on the Left channel from a computation that failed with
   * one (`ERROR.TYPE(1/0)` is 2); host failures stay loud.
   */
  val errorType: FunctionSpec[CellValue] { type Args = UnaryCellValue } =
    FunctionSpec.simple[CellValue, UnaryCellValue](
      "ERROR.TYPE",
      Arity.one,
      flags = FunctionFlags(lift = ArrayLift.all)
    ) { (expr, ctx) =>
      def number(err: CellError): CellValue = CellValue.Number(BigDecimal(err.errorTypeNumber))
      evalValue(ctx, expr) match
        case Left(failure) =>
          EvalError.toErrorValue(failure).map(number).toRight(failure)
        case Right(ExprValue.Cell(cv)) =>
          Right(ArrayArithmetic.carriedError(cv).fold(CellValue.Error(CellError.NA))(number))
        case Right(_) => Right(CellValue.Error(CellError.NA))
    }

  val iserror: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue](
      "ISERROR",
      Arity.one,
      flags = FunctionFlags(lift = ArrayLift.all)
    ) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(_) => Right(true)
        case Right(ExprValue.Cell(cv)) => Right(ArrayArithmetic.carriedError(cv).isDefined)
        case Right(_) => Right(false)
    }

  val iserr: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue](
      "ISERR",
      Arity.one,
      flags = FunctionFlags(lift = ArrayLift.all)
    ) { (expr, ctx) =>
      evalValue(ctx, expr) match
        // GH-344: Excel's ISERR excludes #N/A on the Left channel too (`=ISERR(1<na-cell)` is
        // FALSE); every other failure — error value or host — stays TRUE like ISERROR.
        case Left(EvalError.ErrorValue(CellError.NA, _)) => Right(false)
        case Left(_) => Right(true)
        case Right(ExprValue.Cell(cv)) =>
          Right(ArrayArithmetic.carriedError(cv).exists(_ != CellError.NA))
        case Right(_) => Right(false)
    }

  /**
   * GH-511: ISNA — the third member of the family, absent while ISERROR and ISERR shipped.
   * `DataTableSeeder.ErrorPredicates` already named it, so the seeder's guard walk carried an
   * unreachable branch for it.
   */
  val isna: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue](
      "ISNA",
      Arity.one,
      flags = FunctionFlags(lift = ArrayLift.all)
    ) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(EvalError.ErrorValue(CellError.NA, _)) => Right(true)
        case Left(_) => Right(false)
        case Right(ExprValue.Cell(cv)) =>
          Right(ArrayArithmetic.carriedError(cv).contains(CellError.NA))
        case Right(_) => Right(false)
    }

  val isnumber: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue](
      "ISNUMBER",
      Arity.one,
      flags = FunctionFlags(lift = ArrayLift.all)
    ) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(_) => Right(false)
        case Right(ExprValue.Cell(CellValue.Number(_))) => Right(true)
        case Right(ExprValue.Cell(CellValue.Formula(_, Some(CellValue.Number(_)), _))) =>
          Right(true)
        case Right(ExprValue.Number(_)) => Right(true)
        case Right(_) => Right(false)
    }

  /** ISTEXT(value): TRUE for text, rich text (#775) and a formula's cached text included. */
  val istext: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue](
      "ISTEXT",
      Arity.one,
      flags = FunctionFlags(lift = ArrayLift.all)
    )((expr, ctx) => Right(isTextValue(evalValue(ctx, expr))))

  /**
   * ISNONTEXT(value): the negation of [[istext]] — TRUE for a number, logical, error or blank cell,
   * FALSE for text, empty text (`""`) and rich text (#775).
   */
  val isnontext: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue](
      "ISNONTEXT",
      Arity.one,
      flags = FunctionFlags(lift = ArrayLift.all)
    )((expr, ctx) => Right(!isTextValue(evalValue(ctx, expr))))

  /** The one "is text" test of ISTEXT and ISNONTEXT ([[ScalarCoercion.cellText]] for a cell). */
  private def isTextValue(value: Either[EvalError, ExprValue]): Boolean =
    value match
      case Right(ExprValue.Cell(cv)) => ScalarCoercion.cellText(cv).isDefined
      case Right(ExprValue.Text(_)) => true
      case _ => false

  val isblank: FunctionSpec[Boolean] { type Args = UnaryCellValue } =
    FunctionSpec.simple[Boolean, UnaryCellValue](
      "ISBLANK",
      Arity.one,
      flags = FunctionFlags(lift = ArrayLift.all)
    ) { (expr, ctx) =>
      evalValue(ctx, expr) match
        case Left(_) => Right(false)
        case Right(ExprValue.Cell(CellValue.Empty)) => Right(true)
        case Right(_) => Right(false)
    }

  /**
   * GH-476: N(value) — Excel's information-family coercion, absent from the roster and therefore an
   * UnknownFunction host error on ordinary banker files.
   *
   * Excel's table: numbers pass through, dates become their serial, TRUE/FALSE become 1/0, text and
   * blanks become 0, and an error value propagates. References and formula caches carry errors on
   * the value channel, so N lifts them onto the absorbing Left channel here.
   */
  val n: FunctionSpec[BigDecimal] { type Args = UnaryCellValue } =
    FunctionSpec.simple[BigDecimal, UnaryCellValue](
      "N",
      Arity.one,
      flags = FunctionFlags(returnsNumeric = true, lift = ArrayLift.all)
    ) { (expr, ctx) =>
      evalValue(ctx, expr).flatMap(numericCoercionN)
    }

  private def numericCoercionN(value: ExprValue): Either[EvalError, BigDecimal] =
    value match
      case ExprValue.Number(x) => Right(x)
      case ExprValue.Bool(b) => Right(if b then BigDecimal(1) else BigDecimal(0))
      case ExprValue.Date(d) =>
        Right(BigDecimal(CellValue.dateTimeToExcelSerial(d.atStartOfDay())))
      case ExprValue.DateTime(dt) => Right(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case ExprValue.Text(_) => Right(BigDecimal(0))
      case ExprValue.Opaque(_) => Right(BigDecimal(0))
      case ExprValue.Cell(cv) => numericCoercionNCell(cv)

  private def numericCoercionNCell(cv: CellValue): Either[EvalError, BigDecimal] =
    cv match
      case CellValue.Number(x) => Right(x)
      case CellValue.Bool(b) => Right(if b then BigDecimal(1) else BigDecimal(0))
      case CellValue.DateTime(dt) => Right(BigDecimal(CellValue.dateTimeToExcelSerial(dt)))
      case CellValue.Formula(_, Some(cached), _) => numericCoercionNCell(cached)
      case CellValue.Error(err) => Left(EvalError.ErrorValue(err, Some("N")))
      case _ => Right(BigDecimal(0))

  /**
   * GH-684: T(value) — N's text counterpart. Microsoft's "T function" page: "If value is or refers
   * to text, T returns value. If value does not refer to text, T returns "" (empty text)." Rich
   * text is text, so it returns its plain text. A formula cell answers T of its cached value.
   *
   * The T page does not mention errors. An error value propagates, as the "N function" page states
   * for N and as LibreOffice caches `T(1/0)` (`#DIV/0!`). A range argument lifts element-wise
   * (`ArrayLift.all`): it spills in an array context and is implicitly intersected in a plain cell.
   *
   * see: https://support.microsoft.com/en-us/office/t-function-fb83aeec-45e7-4924-af95-53e073541228
   * see: https://support.microsoft.com/en-us/office/n-function-a624cad1-3635-4208-b54a-29733d1278c9
   */
  val t: FunctionSpec[String] { type Args = UnaryCellValue } =
    FunctionSpec.simple[String, UnaryCellValue](
      "T",
      Arity.one,
      flags = FunctionFlags(lift = ArrayLift.all)
    ) { (expr, ctx) =>
      evalValue(ctx, expr).flatMap(textCoercionT)
    }

  private def textCoercionT(value: ExprValue): Either[EvalError, String] =
    value match
      case ExprValue.Text(s) => Right(s)
      case ExprValue.Cell(cv) => textCoercionTCell(cv)
      case ExprValue.Number(_) | ExprValue.Bool(_) | ExprValue.Date(_) | ExprValue.DateTime(_) |
          ExprValue.Opaque(_) =>
        Right("")

  @annotation.tailrec
  private def textCoercionTCell(cv: CellValue): Either[EvalError, String] =
    cv match
      case CellValue.Text(s) => Right(s)
      case CellValue.RichText(rt) => Right(rt.toPlainText)
      case CellValue.Formula(_, Some(cached), _) => textCoercionTCell(cached)
      case CellValue.Error(err) => Left(EvalError.ErrorValue(err, Some("T")))
      case _ => Right("")
