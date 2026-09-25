package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue, FormulaKind}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{ArrayResult, EvalError, Evaluator}
import com.tjclp.xl.formula.functions.{ArgValue, ArrayLift, FunctionRegistry, FunctionSpec}
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAllNoShrink

/**
 * The laws of Excel array lifting, over the whole registry: the slot law (rebuilding a call from
 * its own scalar slots is the identity), the flag law (a lifted function has only scalar and range
 * slots and answers a scalar), the pinned lifted and excluded rosters, and the lifting law —
 * `f(range)[i] == f(cell_i)` for every lifted slot of every flagged function, over the contents
 * `genCell` generates (numbers, the texts "abc", "" and "a", blanks, booleans, #N/A, #DIV/0! and a
 * date). The law does not hold beyond that domain for two cell shapes whose single-reference
 * decoders predate lifting — numeric text and cached formula cells — and the carve-out test pins
 * both, so a fix to the decoders has to revisit it.
 */
class ArrayLiftingLawsSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(12)

  private val name = SheetName.unsafe("Sheet1")

  /**
   * B2:D4 holds the Weaver block with two zeros (D3, D4) and one blank (C4), so a range argument
   * tells a blank from a zero; the lifting law overwrites B2:B4 with generated contents.
   */
  private val base: Sheet = List[(ARef, CellValue)](
    ref"A1" -> CellValue.Text("Label"),
    ref"B2" -> CellValue.Number(5),
    ref"C2" -> CellValue.Number(-3),
    ref"D2" -> CellValue.Number(1),
    ref"B3" -> CellValue.Number(-2),
    ref"C3" -> CellValue.Number(4),
    ref"D3" -> CellValue.Number(0),
    ref"B4" -> CellValue.Number(7),
    ref"D4" -> CellValue.Number(0)
  ).foldLeft(Sheet(name)) { case (s, (at, v)) => s.put(at, v) }

  private val lifted: List[FunctionSpec[?]] =
    FunctionRegistry.all.filter(_.flags.lift != ArrayLift.Off)

  // ===== Rosters =====

  private val liftedAll = Set(
    // math
    "ABS",
    "SIGN",
    "SQRT",
    "INT",
    "TRUNC",
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "CEILING",
    "FLOOR",
    "MOD",
    "POWER",
    "EXP",
    "LN",
    "LOG",
    // text
    "LEN",
    "LEFT",
    "RIGHT",
    "MID",
    "UPPER",
    "LOWER",
    "TRIM",
    "SUBSTITUTE",
    "FIND",
    "SEARCH",
    "TEXT",
    "VALUE",
    "CONCATENATE",
    // information
    "ISNUMBER",
    "ISTEXT",
    "ISBLANK",
    "ISERROR",
    "ISERR",
    "ISNA",
    "N",
    "ERROR.TYPE",
    // dates
    "DATE",
    "YEAR",
    "MONTH",
    "DAY",
    "DATEDIF",
    // time value of money
    "PMT",
    "PV",
    "FV",
    "NPER",
    "RATE",
    "RRI",
    // lookup, criteria, statistics k-slots, NPV's rate
    "VLOOKUP",
    "HLOOKUP",
    "MATCH",
    "INDEX",
    "SUMIF",
    "COUNTIF",
    "AVERAGEIF",
    "SUMIFS",
    "COUNTIFS",
    "AVERAGEIFS",
    "MAXIFS",
    "MINIFS",
    "LARGE",
    "SMALL",
    "PERCENTILE",
    "QUARTILE",
    "RANK",
    "NPV"
  )
  private val liftedArraysOnly =
    Set("MROUND", "EDATE", "EOMONTH", "WORKDAY", "NETWORKDAYS", "YEARFRAC")
  private val liftedFirstSlot = Set("IFERROR", "IFNA", "XLOOKUP")

  /** Functions with their own array semantics, array/reference results, or nothing to lift. */
  private val excluded = Set(
    "IF",
    "IFS",
    "SWITCH",
    "CHOOSE",
    "NOT",
    "AND",
    "OR",
    "SUM",
    "AVERAGE",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "MAX",
    "MIN",
    "MEDIAN",
    "STDEV",
    "STDEVP",
    "VAR",
    "VARP",
    "SUMPRODUCT",
    "TRANSPOSE",
    "SORT",
    "UNIQUE",
    "FILTER",
    "SEQUENCE",
    "ANCHORARRAY",
    "SINGLE",
    "ROW",
    "COLUMN",
    "ROWS",
    "COLUMNS",
    "CELL",
    "OFFSET",
    "INDIRECT",
    "ADDRESS",
    "HYPERLINK",
    "IRR",
    "XIRR",
    "XNPV",
    "RAND",
    "RANDBETWEEN",
    "NA",
    "NOW",
    "TODAY",
    "PI"
  )

  test("the lifted roster is pinned: 74 functions with their slot policies") {
    val byPolicy = lifted.groupMap(_.flags.lift)(_.name.toUpperCase).view.mapValues(_.toSet).toMap
    assertEquals(byPolicy.getOrElse(ArrayLift.all, Set.empty), liftedAll)
    assertEquals(byPolicy.getOrElse(ArrayLift.arraysOnly, Set.empty), liftedArraysOnly)
    assertEquals(byPolicy.getOrElse(ArrayLift.slots(0), Set.empty), liftedFirstSlot)
    assertEquals(lifted.size, 74)
  }

  test("the excluded functions do not lift") {
    val registered = FunctionRegistry.all.map(spec => spec.name.toUpperCase -> spec).toMap
    excluded.foreach { fn =>
      val spec = registered.getOrElse(fn, fail(s"$fn is not registered"))
      assertEquals(spec.flags.lift, ArrayLift.Off, fn)
    }
  }

  // ===== Sample calls for every registered function =====

  private val scalarKinds = Set("number", "integer", "text", "boolean", "date", "cell", "value")

  /** The slot kinds of a call with `arity` arguments, from the parser's slot description. */
  private def slotKinds(parts: List[String], arity: Int): Option[List[String]] =
    val kinds = parts.foldLeft(Vector.empty[String]) { (acc, part) =>
      if part.endsWith("...") then
        val group = part.stripSuffix("...").split(", ").toVector
        Iterator
          .iterate(acc)(_ ++ group)
          .takeWhile(_.size <= arity)
          .toVector
          .lastOption
          .getOrElse(acc)
      else if part.startsWith("optional ") then
        if acc.size < arity then acc :+ part.stripPrefix("optional ") else acc
      else acc :+ part
    }
    Option.when(kinds.size == arity)(kinds.toList)

  private def arities(arity: Arity): List[Int] = arity match
    case Arity.Exact(n) => List(n)
    case Arity.Range(min, max) => (min to math.min(max, min + 2)).toList
    case Arity.AtLeast(min) => (min to min + 2).toList

  private def scalarArgument(kind: String): String = kind match
    case "text" => "\"a\""
    case "boolean" => "TRUE"
    case "date" => "DATE(2026,1,1)"
    case "range" | "array or range" | "number or range" => "C2:D4"
    case _ => "1"

  private def sampleCalls(spec: FunctionSpec[?]): List[(Int, List[String])] =
    for
      arity <- arities(spec.arity)
      kinds <- slotKinds(spec.argSpec.describeParts, arity).toList
    yield (arity, kinds)

  private def parseCall(formula: String): Option[TExpr.Call[?]] =
    FormulaParser.parse(formula).toOption.collect { case call: TExpr.Call[?] => call }

  test("slot law: rebuilding a call from its own scalar slots is the identity") {
    val calls =
      for
        spec <- FunctionRegistry.all
        (_, kinds) <- sampleCalls(spec)
        call <- parseCall(s"=${spec.name}(${kinds.map(scalarArgument).mkString(",")})").toList
      yield call
    assert(calls.sizeIs > 150, s"only ${calls.size} sample calls parse")
    calls.foreach { call =>
      val slots = call.spec.argSpec.scalarSlots(call.args).map(_._1)
      assertEquals(
        call.spec.argSpec.replaceScalarSlots(call.args, slots),
        (call.args, Nil),
        call.spec.name
      )
      // scalar slots are argument expressions, never a range slot or an aggregate's arm
      val expressions = call.spec.argSpec.toValues(call.args).collect { case ArgValue.Expr(e) => e }
      slots.foreach(slot => assert(expressions.exists(_ eq slot), s"${call.spec.name}: $slot"))
    }
  }

  test("flag law: a lifted function has only scalar and range slots and answers a scalar") {
    // INDEX alone answers a reference (its per-element result collapses to the top-left cell)
    val referenceResults = Set("INDEX")
    lifted.foreach { spec =>
      val parts = spec.argSpec.describeParts.flatMap(
        _.stripPrefix("optional ").stripSuffix("...").split(", ")
      )
      parts.foreach(kind =>
        assert(scalarKinds(kind) || kind == "range", s"${spec.name} has a '$kind' slot")
      )
      if !referenceResults(spec.name) then
        sampleCalls(spec).foreach { (_, kinds) =>
          val formula = s"=${spec.name}(${kinds.map(scalarArgument).mkString(",")})"
          parseCall(formula).foreach { call =>
            Evaluator.arrayInstance.eval(call, base, Clock.system, Some(Workbook(base)), None) match
              case Right(ar: ArrayResult) => fail(s"$formula answered an array: $ar")
              case _ => ()
          }
        }
    }
  }

  // ===== The lifting law =====

  private val genCell: Gen[CellValue] = Gen.frequency(
    6 -> Gen.choose(-4, 30).map(n => CellValue.Number(BigDecimal(n))),
    2 -> Gen.oneOf("0.5", "2.75", "-1.25").map(s => CellValue.Number(BigDecimal(s))),
    2 -> Gen.oneOf(CellValue.Text("abc"), CellValue.Text(""), CellValue.Text("a")),
    2 -> Gen.const(CellValue.Empty),
    1 -> Gen.oneOf(CellValue.Bool(true), CellValue.Bool(false)),
    1 -> Gen.oneOf(CellValue.Error(CellError.NA), CellValue.Error(CellError.Div0)),
    1 -> Gen.const(CellValue.DateTime(java.time.LocalDateTime.of(2024, 2, 29, 0, 0)))
  )

  /**
   * One probe per (function, arity, lifted scalar position): the formula with `#` where the lifted
   * argument goes. The Analysis ToolPak lineage takes `+#` — an array, since a range reference is
   * `#VALUE!` there by design.
   */
  private val probes: List[(String, String)] =
    for
      spec <- lifted
      (_, kinds) <- sampleCalls(spec)
      position <- kinds.indices.toList
      if scalarKinds(kinds(position))
      hole = if spec.flags.lift == ArrayLift.arraysOnly then "+#" else "#"
      template = s"=${spec.name}(${kinds.zipWithIndex
          .map((k, i) => if i == position then hole else scalarArgument(k))
          .mkString(",")})"
      if liftsAt(template)
    yield (spec.name, template)

  /** The `#` argument sits in a lifted slot (IFERROR's fallback and XLOOKUP's others do not). */
  private def liftsAt(template: String): Boolean =
    parseCall(template.replace("#", "Z9:Z10")).exists { call =>
      call.spec.flags.lift match
        case ArrayLift.On(slots, _) =>
          call.spec.argSpec.scalarSlots(call.args).map(_._1).zipWithIndex.exists { (slot, i) =>
            containsRange(slot) && slots.forall(_.contains(i))
          }
        case ArrayLift.Off => false
    }

  private def containsRange(expr: TExpr[?]): Boolean =
    TExpr.collectRanges(expr).exists((_, range) => range.toA1 == "Z9:Z10")

  /** A scalar result as the element the lifted call would hold, or None for a host failure. */
  private def asElement(result: Either[EvalError, Any]): Option[CellValue] = result match
    case Right(value) =>
      Some(value match
        case cv: CellValue => cv
        case ar: ArrayResult => if ar.isEmpty then CellValue.Empty else ar(0, 0)
        case bd: BigDecimal => CellValue.Number(bd)
        case s: String => CellValue.Text(s)
        case b: Boolean => CellValue.Bool(b)
        case i: Int => CellValue.Number(BigDecimal(i))
        case d: java.time.LocalDate => CellValue.DateTime(d.atStartOfDay())
        case dt: java.time.LocalDateTime => CellValue.DateTime(dt)
        case other => CellValue.Text(other.toString))
    case Left(error) =>
      EvalError
        .toErrorValue(error)
        .map(CellValue.Error(_))
        .orElse(error match
          case EvalError.TypeMismatch(_, _, _) | EvalError.CodecFailed(_, _) =>
            Some(CellValue.Error(CellError.Value))
          case _ => None)

  test("the probes cover every lifted function, and most of them lift to an array") {
    assertEquals(probes.map(_._1).toSet, lifted.map(_.name).toSet)
    val wb = Some(Workbook(base))
    val arrays = probes.count { (_, template) =>
      FormulaParser.parse(template.replace("#", "B2:B4")).toOption.exists { expr =>
        Evaluator.arrayInstance.eval(expr, base, Clock.system, wb, None) match
          case Right(ar: ArrayResult) => ar.rows == 3 && ar.cols == 1
          case _ => false
      }
    }
    assert(probes.sizeIs > 150, s"only ${probes.size} probes")
    assert(arrays * 10 >= probes.size * 8, s"only $arrays of ${probes.size} probes lift to 3x1")
  }

  property("lifting law: f(range)[i] == f(cell_i) for every lifted slot") {
    forAllNoShrink(Gen.listOfN(3, genCell)) { cells =>
      val sheet = List(ref"B2", ref"B3", ref"B4").zip(cells).foldLeft(base) { case (s, (at, v)) =>
        s.put(at, v)
      }
      val wb = Some(Workbook(sheet))
      def eval(formula: String): Either[EvalError, Any] =
        FormulaParser.parse(formula) match
          case Left(e) => fail(s"$formula does not parse: $e")
          case Right(expr) => Evaluator.arrayInstance.eval(expr, sheet, Clock.system, wb, None)
      Prop.all(probes.map { (_, template) =>
        val liftedResult = eval(template.replace("#", "B2:B4"))
        val elements = List("B2", "B3", "B4").map(c => asElement(eval(template.replace("#", c))))
        val expected: Option[ArrayResult] =
          Option.when(elements.forall(_.isDefined))(
            ArrayResult(elements.flatten.toVector.map(Vector(_)))
          )
        val holds = (expected, liftedResult) match
          case (Some(array), Right(result)) => result == array
          case (None, Left(_)) => true
          case _ => false
        Prop(holds) :| s"$template over $cells: lifted $liftedResult, element-wise $elements"
      }*)
    }
  }

  test("lifting-law carve-outs: numeric text and cached formula cells through one reference") {
    // Both divergences are the scalar reference decoders', and the lifted element is Excel's:
    // a numeric-text cell coerces to its number element-wise, while one reference to it in a
    // numeric slot is a type mismatch (as `=B2+0` is); one reference to a cached formula cell in a
    // text slot reads the formula's text (TExprDecoders.decodeAsString), the element its cached value
    val sheet = base
      .put(ref"B2", CellValue.Text("5"))
      .put(ref"B3", CellValue.Text("5"))
      .put(ref"E2", CellValue.Formula("1+1", Some(CellValue.Number(2)), FormulaKind.Normal()))
      .put(ref"E3", CellValue.Formula("10+1", Some(CellValue.Number(11)), FormulaKind.Normal()))
    def eval(formula: String): Either[EvalError, Any] =
      FormulaParser.parse(formula) match
        case Left(e) => fail(s"$formula does not parse: $e")
        case Right(expr) =>
          Evaluator.arrayInstance.eval(expr, sheet, Clock.system, Some(Workbook(sheet)), None)
    def column(values: Int*): ArrayResult =
      ArrayResult(values.toVector.map(n => Vector(CellValue.Number(BigDecimal(n)))))

    assertEquals(eval("=ABS(B2:B3)"), Right(column(5, 5)))
    assertEquals(eval("=SUMPRODUCT(ABS(B2:B3))"), Right(BigDecimal(10)))
    eval("=ABS(B2)") match
      case Left(EvalError.CodecFailed(_, _)) => ()
      case other => fail(s"=ABS(B2) over the text 5: expected a type mismatch, got $other")

    assertEquals(eval("=LEN(E2:E3)"), Right(column(1, 2)))
    assertEquals(eval("=LEN(E2)"), Right(BigDecimal(3)), "the formula text 1+1")
  }
