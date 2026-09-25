package com.tjclp.xl.formula.functions

import com.tjclp.xl.formula.Arity
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator, RangeOperand}
import com.tjclp.xl.formula.parser.ParseError

import com.tjclp.xl.cells.{CellError, CellValue}

/**
 * GH-669: Excel's reference operators — union `,` (inside parentheses: `SUM((A1,A2))`) and
 * intersection ` ` (a single space: `A1:C3 B2:D4`).
 *
 * Both are unregistered `TExpr.Call` specs, like `@x` (SINGLE) and `x#` (ANCHORARRAY): every walker
 * already reaches a call's operands through its ArgSpec, so dependency edges, shifting, sheet
 * renames and externality come for free, and their names are not identifiers, so `=UNION(…)` stays
 * an unknown function and `xl functions` never lists them. The object lives outside the
 * `FunctionSpecs*` traits for that reason (FunctionRegistryMacro collects every spec those
 * declare).
 *
 * No multi-area value ever enters the evaluator's value plane. A union's areas exist only inside
 * [[areas]], which the consumers that accept several areas call — the variadic aggregates, INDEX's
 * area_num, AREAS and the intersection itself. Everywhere else a union evaluates through its own
 * `eval`, Excel's `#VALUE!` for a multi-area reference in a value position. An intersection that
 * resolves to one area is an ordinary reference, so every reference position takes it unchanged.
 *
 * Initialization: this object reads `Evaluator` only inside defs, never in a val initializer, so it
 * cannot join an object-initialization cycle with `FunctionSpecs` (whose AREAS uses it).
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
private[formula] object ReferenceOperators extends FunctionSpecsBase:

  /** The operators' spec names — never identifiers, so no formula can call them by name. */
  final val UnionName = "(union)"
  final val IntersectionName = "(intersection)"

  /** One operand: a location the parser recognized, or a reference-valued expression. */
  type Operand = Either[TExpr.RangeLocation, TExpr[Any]]

  /** A call to one of the two operators, looking through the parser's coercion wrapper. */
  object OperatorCall:
    def unapply(e: TExpr[?]): Option[TExpr.Call[?]] = e match
      case TExpr.Coerced(inner, _) => unapply(inner)
      case call: TExpr.Call[?]
          if call.spec.name == UnionName || call.spec.name == IntersectionName =>
        Some(call)
      case _ => None

  def isOperatorCall(e: TExpr[?]): Boolean = OperatorCall.unapply(e).isDefined

  /**
   * Can `e` be an operand of `,` or ` `? Every reference the parser can see: a cell, range, name,
   * LET name or error literal (`A1:B2 #REF!` is what a deletion leaves), a union or intersection,
   * and a call to a function that returns a reference (IF, IFS, CHOOSE, SWITCH, OFFSET, INDIRECT,
   * INDEX). Literals, arithmetic, comparisons, `@x`, `x#`, array constants and every other call are
   * values — Excel refuses them too.
   */
  def isReferenceShape(e: TExpr[?]): Boolean = e match
    case TExpr.Coerced(inner, _) => isReferenceShape(inner)
    case _: TExpr.RangeRef | _: TExpr.SheetRange | _: TExpr.ExternalRange | _: TExpr.PolyRef |
        _: TExpr.Ref[?] | _: TExpr.SheetPolyRef | _: TExpr.SheetRef[?] | _: TExpr.ExternalRef |
        _: TExpr.NameRef | _: TExpr.SheetNameRef | _: TExpr.ErrorLit | _: TExpr.BindingRef |
        _: TExpr.CoercedBindingRef[?] =>
      true
    case call: TExpr.Call[?] =>
      isOperatorCall(call) || Evaluator.referenceFunctions.contains(call.spec.name)
    case _ => false

  /**
   * An operand slot: a location as `Left` (every shape `ArgSpec.rangeLocation` accepts — a single
   * cell is its 1×1 range), otherwise the expression as `Right` when `acceptRight` admits it, else
   * the parse error a range slot gives. An operand is never an array-lifted scalar slot.
   */
  def operandSpec(acceptRight: TExpr[?] => Boolean): ArgSpec[Operand] = new ArgSpec[Operand]:
    def describeParts: List[String] = List("reference")

    def parse(
      args: List[TExpr[?]],
      pos: Int,
      fnName: String
    ): Either[ParseError, (Operand, List[TExpr[?]])] =
      args match
        case head :: tail =>
          ArgSpec.rangeLocation.parse(List(head), pos, fnName) match
            case Right((location, _)) => Right((Left(location), tail))
            case Left(_) if acceptRight(head) => Right((Right(head.asInstanceOf[TExpr[Any]]), tail))
            case Left(err) => Left(err)
        case Nil => Left(ParseError.InvalidArguments(fnName, pos, describe, "0 arguments"))

    def toValues(args: Operand): List[ArgValue] = args match
      case Left(location) => List(ArgValue.Range(location))
      case Right(expr) => List(ArgValue.Expr(expr))

    def map(args: Operand)(
      mapExpr: TExpr[?] => TExpr[?],
      mapRange: TExpr.RangeLocation => TExpr.RangeLocation,
      mapCells: com.tjclp.xl.CellRange => com.tjclp.xl.CellRange
    ): Operand = args match
      case Left(location) => Left(mapRange(location))
      case Right(expr) => Right(mapExpr(expr).asInstanceOf[TExpr[Any]])

  /** An operand of `,` or ` ` (and AREAS' argument): any reference shape. */
  val referenceOperand: ArgSpec[Operand] = operandSpec(isReferenceShape)

  /**
   * INDEX's first slot: a location exactly as before, an operator call (area_num picks among its
   * areas) or an array constant (INDEX indexes the constant's values). Any other expression keeps
   * the range slot's parse error.
   */
  val indexOperand: ArgSpec[Operand] = operandSpec(e =>
    isOperatorCall(e) || (e match
      case TExpr.Lit(_: ArrayResult) => true
      case _ => false)
  )

  private def operandText(operand: Operand, printer: ArgPrinter): String = operand match
    case Left(location) => printer.location(location)
    case Right(expr) => printer.expr(expr)

  /**
   * `(a,b,…)`. Printed with a bare `,` in every form: a space is itself an operator here, so the
   * human form and the file form are the same text.
   */
  val union: FunctionSpec[CellValue] { type Args = List[Operand] } =
    FunctionSpec.simple[CellValue, List[Operand]](
      UnionName,
      Arity.AtLeast(2),
      renderFn =
        Some((operands, printer) => operands.map(operandText(_, printer)).mkString("(", ",", ")"))
    )((_, _) => Left(multiAreaValue))(using ArgSpec.list(using referenceOperand))

  /** Excel's answer for a multi-area reference in a value position: `=(A1,A2)`, `ABS((A1,A2))`. */
  private def multiAreaValue: EvalError =
    EvalError.ErrorValue(
      CellError.Value,
      Some("a multi-area reference (a union) is not a value here")
    )

  /**
   * `a b` — the cells two references share. One area is an ordinary reference (a plain cell reads
   * it through the implicit intersection, array mode whole); two or more (a union intersecting
   * several areas) are `#VALUE!` outside the consumers that take several; none is `#NULL!`.
   */
  val intersection: FunctionSpec[ArrayResult] { type Args = (Operand, Operand) } =
    FunctionSpec.referencing[ArrayResult, (Operand, Operand)](
      IntersectionName,
      Arity.two,
      renderFn = Some { case ((left, right), printer) =>
        val rightText = operandText(right, printer)
        // left-associative: a right operand that is itself an intersection keeps its parens
        val grouped = right match
          case Right(OperatorCall(call)) if call.spec.name == IntersectionName => s"($rightText)"
          case _ => rightText
        s"${operandText(left, printer)} $grouped"
      }
    )((args, ctx) => singleArea(args, ctx)) { (args, ctx) =>
      singleArea(args, ctx).flatMap { case RangeOperand(sheet, range) =>
        referenceResult(range, sheet, ctx)(
          extractRangeAsMatrixEval(range, sheet, ctx).map(ArrayResult(_))
        )
      }
    }(using ArgSpec.tuple(using referenceOperand, ArgSpec.tuple(using referenceOperand)))

  private def singleArea(
    args: (Operand, Operand),
    ctx: EvalContext
  ): Either[EvalError, RangeOperand] =
    intersect(args._1, args._2, ctx).flatMap {
      case Vector(one) => Right(one)
      case _ =>
        Left(
          EvalError.ErrorValue(CellError.Value, Some("a multi-area reference is not accepted here"))
        )
    }

  /** Build `(a,b,…)` from parsed operands, each of which must be a reference. */
  def mkUnion(areas: List[TExpr[?]], pos: Int): Either[ParseError, TExpr[?]] =
    if areas.size < 2 then Left(ParseError.InvalidOperator(",", pos, "a union needs two areas"))
    else if !areas.forall(isReferenceShape) then
      Left(ParseError.InvalidOperator(",", pos, "every area of a union must be a reference"))
    else
      ArgSpec.list(using referenceOperand).parse(areas, pos, UnionName).map { case (operands, _) =>
        TExpr.Call(union, operands)
      }

  /** Build `l r`; both operands must be references. */
  def mkIntersection(l: TExpr[?], r: TExpr[?], pos: Int): Either[ParseError, TExpr[?]] =
    if !(isReferenceShape(l) && isReferenceShape(r)) then
      Left(ParseError.InvalidOperator(" ", pos, "intersection operands must be references"))
    else
      for
        (left, _) <- referenceOperand.parse(List(l), pos, IntersectionName)
        (right, _) <- referenceOperand.parse(List(r), pos, IntersectionName)
      yield TExpr.Call(intersection, (left, right))

  /**
   * The areas an operand denotes, left to right — the ONLY multi-area resolution. A location is one
   * area (its errors propagate: `#REF!`, a missing sheet, a name bound to a non-reference). A union
   * concatenates its operands' areas, duplicates and overlaps kept (`SUM((A1,A1))` is 2×A1; nested
   * unions flatten), all on one sheet. An intersection intersects every pair of its operands'
   * areas, left-major, keeping the non-empty ones; none is `#NULL!`. Any other expression is the
   * reference it evaluates to, or `#VALUE!` when it evaluates to a value. The first error wins.
   */
  private[formula] def areas(
    op: Operand,
    ctx: EvalContext
  ): Either[EvalError, Vector[RangeOperand]] =
    op match
      case Left(location) =>
        Evaluator
          .resolveRangeLocation(location, ctx.sheet, ctx.workbook)
          .map((sheet, range) => Vector(RangeOperand(sheet, range)))
      case Right(OperatorCall(call)) if call.spec.name == UnionName =>
        call.args
          .asInstanceOf[List[Operand]]
          .foldLeft[Either[EvalError, Vector[RangeOperand]]](Right(Vector.empty)) {
            (acc, operand) =>
              acc.flatMap(found => areas(operand, ctx).map(found ++ _))
          }
          .flatMap { found =>
            if found.map(_.sheet.name).distinct.size <= 1 then Right(found)
            else
              Left(
                EvalError.ErrorValue(CellError.Value, Some("a union's areas must be on one sheet"))
              )
          }
      case Right(OperatorCall(call)) =>
        val (left, right) = call.args.asInstanceOf[(Operand, Operand)]
        intersect(left, right, ctx)
      case Right(other) =>
        ctx.evalReference(other).flatMap {
          case operand: RangeOperand => Right(Vector(operand))
          case _ => Left(EvalError.ErrorValue(CellError.Value, Some("not a reference")))
        }

  private def intersect(
    left: Operand,
    right: Operand,
    ctx: EvalContext
  ): Either[EvalError, Vector[RangeOperand]] =
    for
      leftAreas <- areas(left, ctx)
      rightAreas <- areas(right, ctx)
      pairs = for l <- leftAreas; r <- rightAreas yield (l, r)
      _ <-
        if pairs.forall((l, r) => l.sheet.name == r.sheet.name) then Right(())
        else
          Left(
            EvalError.ErrorValue(
              CellError.Value,
              Some("an intersection's references must be on one sheet")
            )
          )
      shared = pairs.flatMap((l, r) => l.range.intersect(r.range).map(RangeOperand(l.sheet, _)))
      result <-
        if shared.nonEmpty then Right(shared)
        else Left(EvalError.ErrorValue(CellError.Null, Some("the intersection is empty")))
    yield result

  /**
   * The static dependency edges of an operator call: an intersection of two static locations on one
   * sheet depends only on the cells they share (whole-operand edges would falsely make
   * `C3: =SUM(A1:C3 A1:A3)` circular), nothing when they share none. Anything else — names, dynamic
   * operands, nesting — keeps every operand's edges, a sound over-approximation.
   */
  def dependencyValues(specName: String, values: List[ArgValue]): List[ArgValue] =
    if specName != IntersectionName then values
    else
      values match
        case List(ArgValue.Range(l), ArgValue.Range(r)) =>
          (l, r) match
            case (TExpr.RangeLocation.Local(a, _), TExpr.RangeLocation.Local(b, _)) =>
              a.intersect(b).map(s => ArgValue.Range(TExpr.RangeLocation.Local(s))).toList
            case (
                  TExpr.RangeLocation.CrossSheet(sa, a, _),
                  TExpr.RangeLocation.CrossSheet(sb, b, _)
                ) if sa.value.equalsIgnoreCase(sb.value) =>
              a.intersect(b).map(s => ArgValue.Range(TExpr.RangeLocation.CrossSheet(sa, s))).toList
            case _ => values
        case _ => values
