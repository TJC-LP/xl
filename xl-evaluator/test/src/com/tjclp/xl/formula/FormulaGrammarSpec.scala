package com.tjclp.xl.formula

import scala.annotation.tailrec
import scala.util.{Failure, Success, Try}

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.error.XLError
import com.tjclp.xl.formula.functions.{ArgPrinter, FunctionRegistry, FunctionSpec}
import com.tjclp.xl.formula.parser.ParseError
import com.tjclp.xl.formula.printer.{FormulaOps, FormulaPrinter}
import com.tjclp.xl.ooxml.FormulaStorage
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.forAllNoShrink

import FormulaTextGens.*

/**
 * GH-653 (item 4): the grammar-complete parse ∘ print generator.
 *
 * The round-trip generators before this spec built ASTs, so they could only ever produce what the
 * AST already modelled. Every generator here produces formula TEXT — the parser decides the tree
 * and the printer every paren — over the grammar Excel writes: sheet qualifiers on both sides of
 * the quoting boundary, external `[n]Name` qualifiers, refs and ranges in every spelling under
 * every qualifier, error literals, defined names, the spill reference `x#`, `@` on every primary,
 * string literals with doubled quotes, every registry function at every legal arity with any subset
 * of slots left empty, the `_xlfn.` storage spellings, LET, and a recursive mixer over all of it.
 *
 * Two laws. For text the parser ACCEPTS: `parse(print(parse(s))) == parse(s)` (the GH-455 pattern —
 * two parser-built trees compare, so `Ref` decoders and numeric re-spellings are handled), `print`
 * is a fixpoint, the file form re-parses, and — for the sub-grammars with one canonical spelling —
 * `print(parse(s)) == s` byte-for-byte. For the grammar the parser REJECTS (3-D spans, array
 * constants, structured references, union, intersection, LAMBDA — no AST node exists for any of
 * them) the law is totality: a `Left` carrying a position inside the text, never a throw; the
 * `FormulaStorage` boundary is the identity on the text, so a workbook carrying it rides through
 * read → write byte-for-byte; and `FormulaOps.shift` refuses with a `FormulaError` rather than
 * rewriting.
 *
 * Determinism: the suite pins the ScalaCheck seed like OoxmlGenerativeRoundTripSpec (GH-308).
 * `XL_ROUNDTRIP_MIN_SUCCESS=2000 XL_ROUNDTRIP_RANDOM_SEED=true` is the weekly law-fuzz setting;
 * `-Dxl.roundtrip.seed=<printed seed>` replays a failure.
 */
class FormulaGrammarSpec extends ScalaCheckSuite:

  // ==================== Seed plumbing (GH-308) ====================

  private val minSuccess: Int =
    sys.props
      .get("xl.roundtrip.minSuccess")
      .orElse(sys.env.get("XL_ROUNDTRIP_MIN_SUCCESS"))
      .flatMap(_.toIntOption)
      .getOrElse(100)

  private val explicitSeed: Option[org.scalacheck.rng.Seed] =
    sys.props.get("xl.roundtrip.seed").flatMap { raw =>
      raw.toLongOption
        .map(org.scalacheck.rng.Seed.apply)
        .orElse(org.scalacheck.rng.Seed.fromBase64(raw).toOption)
    }

  private val randomSeed: Boolean =
    explicitSeed.isEmpty &&
      sys.env.get("XL_ROUNDTRIP_RANDOM_SEED").exists(_.equalsIgnoreCase("true"))

  override def scalaCheckTestParameters: org.scalacheck.Test.Parameters =
    val base = super.scalaCheckTestParameters.withMinSuccessfulTests(minSuccess)
    if randomSeed then
      val seed = org.scalacheck.rng.Seed.random()
      println(
        s"[law-fuzz] random seed for this run: ${seed.toBase64} ($minSuccess cases, GH-653 grammar)"
      )
      base.withInitialSeed(seed)
    else base.withInitialSeed(explicitSeed.getOrElse(org.scalacheck.rng.Seed(20260913L)))

  // ==================== Laws ====================

  private def parsed(source: String): TExpr[?] =
    FormulaParser.parse(source) match
      case Right(expr) => expr
      case Left(err) => fail(s"'$source' should parse: ${ParseError.describe(err)}")

  /**
   * The strong law for accepted text: the printed form re-parses to the SAME tree (decoders
   * included — both trees are parser-built), keeps its typed shape, is a fixpoint of `print`, and
   * its file form re-parses to the same tree too.
   */
  private def assertRoundTrips(source: String): TExpr[?] =
    val e1 = parsed(source)
    val printed = FormulaPrinter.print(e1)
    val reparsed = FormulaParser.parse(printed)
    assertEquals(reparsed, Right(e1), s"'$source' printed as '$printed'")
    assertEquals(
      reparsed.map(FormulaPrinter.printWithTypes),
      Right(FormulaPrinter.printWithTypes(e1)),
      s"typed shape must survive the round-trip of '$source' via '$printed'"
    )
    assertEquals(
      reparsed.map(FormulaPrinter.print(_)),
      Right(printed),
      s"print is not a fixpoint for '$source'"
    )
    val fileForm = FormulaPrinter.printFileForm(e1)
    assertEquals(
      FormulaParser.parse(fileForm),
      Right(e1),
      s"file form '$fileForm' of '$source' must re-parse"
    )
    e1

  /** Text identity: the sub-grammar has one canonical spelling and the printer reproduces it. */
  private def assertPrintsAs(source: String, expected: String): Unit =
    assertEquals(FormulaPrinter.print(parsed(source)), expected, s"source '$source'")

  private def position(err: ParseError): Option[Int] = err match
    case ParseError.UnexpectedChar(_, pos, _) => Some(pos)
    case ParseError.UnexpectedEOF(pos, _) => Some(pos)
    case ParseError.InvalidCellRef(_, pos, _) => Some(pos)
    case ParseError.InvalidNumber(_, pos, _) => Some(pos)
    case ParseError.UnbalancedDelimiter(pos, _, _) => Some(pos)
    case ParseError.UnknownFunction(_, pos, _) => Some(pos)
    case ParseError.InvalidArguments(_, pos, _, _) => Some(pos)
    case ParseError.InvalidOperator(_, pos, _) => Some(pos)
    case ParseError.GenericError(_, pos) => pos
    case ParseError.EmptyFormula | ParseError.FormulaTooLong(_, _) |
        ParseError.NestingTooDeep(_, _) =>
      None

  // ==================== Generator scaffolding ====================

  /**
   * The LET bindings visible where a leaf is generated. A bound name is never a spill anchor (`x#`
   * needs a cell or a defined name) and never fills a strict range slot (a scalar binding there is
   * a parse error, as in Excel), so those generators draw unbound names.
   */
  private final case class Scope(bound: List[String]):
    def isBound(name: String): Boolean = bound.exists(_.equalsIgnoreCase(name))
    def extend(names: List[String]): Scope = Scope(names ++ bound)

  private val NoScope = Scope(Nil)

  private def sequenceGens(gens: List[Gen[String]]): Gen[List[String]] =
    gens.foldRight(Gen.const(List.empty[String])) { (g, acc) =>
      for
        x <- g
        xs <- acc
      yield x :: xs
    }

  private def genUnboundName(scope: Scope): Gen[String] =
    genNameText.map(n => LazyList.iterate(n)(_ + "_u").find(c => !scope.isBound(c)).getOrElse(n))

  private val genQualifiedRef: Gen[String] =
    for
      q <- genQualifier
      r <- genRefText
    yield q + r

  private val genQualifiedRange: Gen[String] =
    for
      q <- genQualifier
      r <- genRangeText
    yield q + r

  private def genQualifiedName(scope: Scope): Gen[String] =
    for
      q <- genLocalOrSheetQualifier
      n <- genUnboundName(scope)
    yield q + n

  /** `x#` on every anchor shape: a cell (qualified or not, any anchor) or a defined name. */
  private def genSpill(scope: Scope): Gen[String] =
    Gen.frequency(3 -> genQualifiedRef, 1 -> genQualifiedName(scope)).map(_ + "#")

  private val genBoolText: Gen[String] = Gen.oneOf("TRUE", "FALSE", "true", "False")
  private val genErrorWritten: Gen[String] = genErrorLitText.map(_._1)

  /** What a strict range slot accepts: a range, a cell, a defined name, or an error literal. */
  private def genRangeArg(scope: Scope): Gen[String] =
    Gen.frequency(
      5 -> genQualifiedRange,
      2 -> genQualifiedRef,
      2 -> genQualifiedName(scope),
      1 -> genErrorWritten
    )

  private def genLeaf(scope: Scope): Gen[String] =
    val bound = if scope.bound.isEmpty then Nil else List(3 -> Gen.oneOf(scope.bound))
    val leaves = List(
      4 -> genNumberText,
      2 -> genStringLitText,
      1 -> genBoolText,
      1 -> genErrorWritten,
      4 -> genQualifiedRef,
      2 -> genQualifiedRange,
      2 -> genQualifiedName(scope),
      1 -> genSpill(scope)
    )
    Gen.frequency((leaves ++ bound)*)

  /** The operand shapes the `@` arm accepts: one primary. */
  private def genPrimary(depth: Int, scope: Scope): Gen[String] =
    val leaves = List(
      3 -> genQualifiedRef,
      3 -> genQualifiedRange,
      2 -> genQualifiedName(scope),
      1 -> genStringLitText,
      1 -> genIntText,
      1 -> genBoolText,
      1 -> genErrorWritten,
      1 -> genSpill(scope)
    )
    val nested =
      if depth <= 0 then Nil
      else
        List(
          2 -> Gen.lzy(genCall(depth - 1, scope)),
          1 -> Gen.lzy(genExpr(depth - 1, scope).map(e => s"($e)")),
          1 -> Gen.lzy(genAt(depth - 1, scope))
        )
    Gen.frequency((leaves ++ nested)*)

  private def genAt(depth: Int, scope: Scope): Gen[String] =
    for
      space <- Gen.frequency(4 -> Gen.const(""), 1 -> Gen.const(" "))
      primary <- genPrimary(depth, scope)
    yield s"@$space$primary"

  private val genBinaryOp: Gen[String] =
    Gen.oneOf("+", "-", "*", "/", "^", "&", "=", "<>", "<", "<=", ">", ">=")

  private val genSpace: Gen[String] = Gen.frequency(3 -> Gen.const(""), 1 -> Gen.const(" "))

  private def genComparison(scope: Scope): Gen[String] =
    for
      l <- genLeaf(scope)
      op <- Gen.oneOf("=", "<>", "<", ">=")
      r <- genLeaf(scope)
    yield s"$l$op$r"

  /** The recursive mixer: every leaf, every operator, calls, `@`, LET, grouping parens. */
  private def genExpr(depth: Int, scope: Scope): Gen[String] =
    if depth <= 0 then genLeaf(scope)
    else
      Gen.frequency(
        3 -> genLeaf(scope),
        6 -> Gen.lzy(for
          l <- genExpr(depth - 1, scope)
          op <- genBinaryOp
          r <- genExpr(depth - 1, scope)
          s1 <- genSpace
          s2 <- genSpace
        yield s"$l$s1$op$s2$r"),
        1 -> Gen.lzy(genExpr(depth - 1, scope).map(e => s"-$e")),
        1 -> Gen.lzy(genExpr(depth - 1, scope).map(e => s"+$e")),
        1 -> Gen.lzy(genExpr(depth - 1, scope).map(e => s"NOT $e")),
        1 -> Gen.lzy(genPrimary(depth - 1, scope).map(p => s"$p%")),
        1 -> Gen.lzy(genExpr(depth - 1, scope).map(e => s"($e)")),
        4 -> Gen.lzy(genCall(depth - 1, scope)),
        1 -> Gen.lzy(genAt(depth - 1, scope)),
        1 -> Gen.lzy(genLet(depth - 1, scope))
      )

  // ==================== Registry-driven calls ====================

  private val registry: List[FunctionSpec[?]] = FunctionRegistry.all

  /** One declared argument slot of a FunctionSpec, read off `argSpec.describeParts`. */
  private enum Slot:
    case Fixed(kind: String)
    case Optional(kind: String)

    /** A repeating group (`number or range...`, `range, value...`), always the last slot. */
    case Variadic(kinds: List[String])

  private def slotsOf(spec: FunctionSpec[?]): List[Slot] =
    spec.argSpec.describeParts.map { part =>
      if part.endsWith("...") then Slot.Variadic(part.dropRight(3).split(", ").toList)
      else if part.startsWith("optional ") then Slot.Optional(part.stripPrefix("optional "))
      else Slot.Fixed(part)
    }

  /** The slot kinds an n-argument call fills, in order, or None when n is not a shape it parses. */
  private def kindsFor(slots: List[Slot], n: Int): Option[List[String]] =
    @tailrec
    def loop(rest: List[Slot], remaining: Int, acc: List[String]): Option[List[String]] =
      rest match
        case Nil => if remaining == 0 then Some(acc.reverse) else None
        case Slot.Fixed(kind) :: tail =>
          if remaining == 0 then None else loop(tail, remaining - 1, kind :: acc)
        case Slot.Optional(kind) :: tail =>
          if remaining == 0 then Some(acc.reverse) else loop(tail, remaining - 1, kind :: acc)
        case Slot.Variadic(kinds) :: Nil =>
          if kinds.nonEmpty && remaining % kinds.size == 0 then
            Some(acc.reverse ++ List.fill(remaining / kinds.size)(kinds).flatten)
          else None
        case Slot.Variadic(_) :: _ => None
    loop(slots, n, Nil)

  /** Every argument count the spec both declares (Arity) and structurally parses (ArgSpec). */
  private def legalArities(spec: FunctionSpec[?]): List[Int] =
    val slots = slotsOf(spec)
    val fixed = slots.count { case Slot.Fixed(_) => true; case _ => false }
    val optional = slots.count { case Slot.Optional(_) => true; case _ => false }
    val group = slots.collectFirst { case Slot.Variadic(kinds) => kinds.size }.getOrElse(0)
    (0 to fixed + optional + group * 3).toList.filter { n =>
      kindsFor(slots, n).isDefined && spec.arity.validate(n, spec.name, 0).isRight
    }

  private val callShapes: List[(FunctionSpec[?], List[Int])] =
    registry.map(spec => (spec, legalArities(spec)))

  private def declaredMin(arity: Arity): Int = arity match
    case Arity.Exact(n) => n
    case Arity.Range(min, _) => min
    case Arity.AtLeast(n) => n

  /** A strict range slot refuses an empty argument (`SUMIF(,">0")` is InvalidArguments). */
  private def isStrictRange(kind: String): Boolean = kind == "range"

  private def genArg(kind: String, depth: Int, scope: Scope): Gen[String] = kind match
    case "range" => genRangeArg(scope)
    case "number or range" | "array or range" =>
      Gen.frequency(3 -> genRangeArg(scope), 3 -> Gen.lzy(genExpr(depth, scope)))
    case "boolean" =>
      Gen.frequency(
        2 -> genBoolText,
        1 -> genComparison(scope),
        2 -> Gen.lzy(genExpr(depth, scope))
      )
    case "text" => Gen.frequency(3 -> genStringLitText, 2 -> Gen.lzy(genExpr(depth, scope)))
    case "integer" =>
      Gen.frequency(3 -> Gen.choose(0, 12).map(_.toString), 2 -> Gen.lzy(genExpr(depth, scope)))
    case "number" => Gen.frequency(3 -> genNumberText, 2 -> Gen.lzy(genExpr(depth, scope)))
    case "date" =>
      Gen.frequency(
        2 -> genQualifiedRef,
        1 -> Gen.const("DATE(2024, 1, 15)"),
        1 -> genIntText,
        2 -> Gen.lzy(genExpr(depth, scope))
      )
    case _ => Gen.lzy(genExpr(depth, scope)) // value, cell

  private def storedPrefix(name: String): String =
    if FormulaStorage.WorksheetScoped.contains(name) then
      FormulaStorage.XlfnPrefix + FormulaStorage.XlwsPrefix
    else FormulaStorage.XlfnPrefix

  /**
   * The spellings of a function name the parser accepts: as declared, lower-case, stored, and with
   * whitespace before the paren (`SUM (A1)` is the call — the arm that caught `NOT (A1)^2` taking
   * the keyword path).
   */
  private val genNameSpelling: Gen[String => String] =
    Gen.frequency[String => String](
      6 -> Gen.const(identity),
      2 -> Gen.const(_.toLowerCase),
      1 -> Gen.oneOf(" ", "  ").map(space => (name: String) => name + space),
      1 -> Gen.const(name =>
        if FormulaStorage.FutureFunctions.contains(name) then storedPrefix(name) + name else name
      )
    )

  /**
   * A generated call: `kinds` are the slot kinds filled, `emptied` the positions written empty —
   * never a strict range slot, and never when n == 1 (`F()` is the zero-argument call).
   */
  private final case class CallShape(
    spec: FunctionSpec[?],
    kinds: List[String],
    emptied: Set[Int],
    text: String
  )

  private def genCallShape(
    spec: FunctionSpec[?],
    arities: List[Int],
    depth: Int,
    scope: Scope,
    spelling: Gen[String => String]
  ): Gen[CallShape] =
    for
      n <- Gen.oneOf(arities)
      kinds = kindsFor(slotsOf(spec), n).getOrElse(Nil)
      mask <- Gen.listOfN(n, Gen.frequency(4 -> Gen.const(false), 1 -> Gen.const(true)))
      emptied =
        if n < 2 then Set.empty[Int]
        else
          kinds
            .zip(mask)
            .zipWithIndex
            .collect {
              case ((kind, true), i) if !isStrictRange(kind) =>
                i
            }
            .toSet
      args <- sequenceGens(kinds.zipWithIndex.map { (kind, i) =>
        if emptied(i) then Gen.const("") else genArg(kind, depth, scope)
      })
      sep <- Gen.oneOf(",", ", ", " , ")
      spell <- spelling
    yield CallShape(spec, kinds, emptied, s"${spell(spec.name)}(${args.mkString(sep)})")

  private def genCall(depth: Int, scope: Scope): Gen[String] =
    Gen
      .oneOf(callShapes.filter(_._2.nonEmpty))
      .flatMap((spec, arities) => genCallShape(spec, arities, depth, scope, genNameSpelling))
      .map(_.text)

  // ==================== LET ====================

  /** A LET binding name: a letter or `_` first (the declaration position), then name characters. */
  private val genLetName: Gen[String] =
    Gen.frequency(
      4 -> genNameText.map(n => nameFrom(if n.startsWith("\\") then "L" + n.drop(1) else n)),
      1 -> Gen.oneOf("sum", "max", "if", "let", "rate", "npv", "x", "_v", "a.b")
    )

  private def distinctNames(raw: List[String]): List[String] =
    raw.foldLeft(List.empty[String]) { (acc, n) =>
      val unique =
        LazyList.iterate(n)(_ + "_").find(c => !acc.exists(_.equalsIgnoreCase(c))).getOrElse(n)
      acc :+ unique
    }

  private def genLet(depth: Int, scope: Scope): Gen[String] =
    for
      k <- Gen.choose(1, 3)
      raw <- Gen.listOfN(k, genLetName)
      names = distinctNames(raw)
      values <- sequenceGens(names.indices.toList.map { i =>
        genExpr(depth, scope.extend(names.take(i)))
      })
      body <- genExpr(depth, scope.extend(names))
      sep <- Gen.oneOf(",", ", ")
    yield
      val pairs = names.zip(values).flatMap((n, v) => List(n, v))
      s"LET(${(pairs :+ body).mkString(sep)})"

  // ==================== Positions a leaf can occupy ====================

  private type Shape = String => String

  /** Positions every reference-like leaf occupies, each with one canonical spelling. */
  private val commonShapes: List[Shape] = List(
    x => x,
    x => s"SUM($x)",
    x => s"COUNTIF($x, 1)",
    x => s"IF(A1, $x, 1)",
    x => s"$x+1",
    x => s"1/$x",
    x => s"-$x",
    x => s"$x%",
    x => s"@$x",
    x => s"IFERROR($x, 0)",
    x => s"$x=1",
    x => s"VLOOKUP(1, $x, 2)",
    x => s"INDEX($x, 1)",
    x => s"SUMPRODUCT($x, B1:B3)"
  )

  /**
   * The spill positions — only a cell or a name may carry `#`. A strict range slot is not among
   * them: `INDEX(A1#, 1)` / `COUNTIF(A1#, x)` are InvalidArguments today (ArgSpec.rangeLocation has
   * no case for the ANCHORARRAY call, while Excel accepts a spill reference wherever a range goes);
   * that needs a RangeLocation case with resolution semantics, tracked as a follow-up of #653.
   */
  private val spillShapes: List[Shape] = List(
    x => s"$x#",
    x => s"@$x#",
    x => s"$x#%",
    x => s"SUM($x#)",
    x => s"$x#*2",
    x => s"IF(A1, $x#, 1)"
  )

  // ==================== Properties: the grammar the parser accepts ====================

  property(
    "GH-653: recursive mixer — parse ∘ print ∘ parse = parse, print is a fixpoint, the file form re-parses"
  ) {
    forAllNoShrink(genExpr(3, NoScope)) { body =>
      assertRoundTrips(s"=$body")
      true
    }
  }

  property("GH-653: refs and ranges print back byte-for-byte under every sheet-qualifier shape") {
    val genCase =
      for
        qualifier <- genQualifier
        isRef <- Gen.oneOf(true, false)
        target <- if isRef then genRefText else genRangeText
        shape <- Gen.oneOf(if isRef then commonShapes ++ spillShapes else commonShapes)
      yield shape(qualifier + target)
    forAllNoShrink(genCase) { body =>
      val source = s"=$body"
      assertPrintsAs(source, source)
      assertRoundTrips(source)
      true
    }
  }

  property("GH-653: error literals round-trip in every position and print upper-cased") {
    val genCase =
      for
        (written, canonical) <- genErrorLitText
        shape <- Gen.oneOf(commonShapes :+ ((x: String) => s""""a"&$x"""))
      yield (shape(written), shape(canonical))
    forAllNoShrink(genCase) { (body, canonicalBody) =>
      assertPrintsAs(s"=$body", s"=$canonicalBody")
      assertRoundTrips(s"=$body")
      true
    }
  }

  property("GH-653: defined names print verbatim, bare or sheet-qualified, in every position") {
    val genCase =
      for
        name <- genQualifiedName(NoScope)
        shape <- Gen.oneOf(
          commonShapes ++ spillShapes ++ List[Shape](x => s"$x&\"a\"", x => s"LEN($x)")
        )
      yield shape(name)
    forAllNoShrink(genCase) { body =>
      val source = s"=$body"
      assertPrintsAs(source, source)
      assertRoundTrips(source)
      true
    }
  }

  property(
    "GH-653: spill references round-trip on every anchor shape; range#, literal#, call# are rejected"
  ) {
    val genAccepted =
      for
        anchor <- Gen.frequency(3 -> genQualifiedRef, 1 -> genQualifiedName(NoScope))
        shape <- Gen.oneOf(spillShapes)
      yield shape(anchor)
    val genRejected =
      Gen
        .frequency(
          3 -> genQualifiedRange,
          1 -> genNumberText,
          1 -> genStringLitText,
          1 -> genBoolText,
          1 -> genErrorWritten,
          1 -> Gen.lzy(genCall(0, NoScope))
        )
        .map(_ + "#")
    forAllNoShrink(genAccepted, genRejected) { (accepted, rejected) =>
      assertPrintsAs(s"=$accepted", s"=$accepted")
      assertRoundTrips(s"=$accepted")
      Try(FormulaParser.parse(s"=$rejected")) match
        case Success(Left(err)) => assert(position(err).isDefined, s"'$rejected': $err")
        case Success(Right(expr)) =>
          fail(s"'$rejected' should be rejected, parsed as ${FormulaPrinter.printWithTypes(expr)}")
        case Failure(t) => fail(s"'$rejected' threw ${t.getClass.getSimpleName}: ${t.getMessage}")
      true
    }
  }

  property("GH-653: implicit intersection @ over every primary shape re-parses to the same tree") {
    forAllNoShrink(genAt(2, NoScope)) { body =>
      val e1 = assertRoundTrips(s"=$body")
      assert(FormulaPrinter.print(e1).startsWith("=@"), s"'$body' should print as an @ operand")
      true
    }
  }

  private def renderedSlots(expr: TExpr[?]): List[Option[String]] = expr match
    case call: TExpr.Call[?] =>
      val printer = ArgPrinter(
        expr = e => FormulaPrinter.print(e, includeEquals = false),
        location = _ => "range",
        cellRange = _ => "cells"
      )
      call.spec.argSpec.renderSlots(call.args, printer)
    case other => fail(s"expected a call, got ${FormulaPrinter.printWithTypes(other)}")

  /**
   * Missing sits at every emptied position; a trailing empty slot keeps its comma in both forms.
   */
  private def assertEmptiedSlots(shape: CallShape, expr: TExpr[?]): Unit =
    val slots = renderedSlots(expr)
    val n = shape.kinds.size
    shape.emptied.foreach { i =>
      assertEquals(slots.lift(i), Some(Some("")), s"slot $i of '${shape.text}' should be empty")
    }
    (0 until n).filterNot(shape.emptied).foreach { i =>
      assert(slots.lift(i).exists(_.exists(_.nonEmpty)), s"slot $i of '${shape.text}' is present")
    }
    if n > 0 && shape.emptied.contains(n - 1) then
      assert(FormulaPrinter.print(expr).endsWith(",)"), s"'${shape.text}' keeps its trailing comma")
      assert(
        FormulaPrinter.printFileForm(expr).endsWith(",)"),
        s"'${shape.text}' keeps its trailing comma in the file form"
      )

  property(
    "GH-653: every registry function round-trips at a legal arity with any subset of slots left empty"
  ) {
    val genShape =
      Gen
        .oneOf(callShapes.filter(_._2.nonEmpty))
        .flatMap((spec, arities) =>
          genCallShape(spec, arities, 1, NoScope, Gen.const(identity[String]))
        )
    forAllNoShrink(genShape) { shape =>
      val e1 = assertRoundTrips(s"=${shape.text}")
      assertEmptiedSlots(shape, e1)
      true
    }
  }

  property("GH-653: _xlfn. and _xlfn._xlws. spellings parse to the bare tree and print bare") {
    val future = callShapes.filter((spec, arities) =>
      arities.nonEmpty && FormulaStorage.FutureFunctions.contains(spec.name)
    )
    assert(future.nonEmpty, "the registry carries future functions")
    val genCase =
      for
        (spec, arities) <- Gen.oneOf(future)
        shape <- genCallShape(spec, arities, 1, NoScope, Gen.const(identity[String]))
        prefix <-
          if FormulaStorage.WorksheetScoped.contains(spec.name) then
            Gen.oneOf("_xlfn._xlws.", "_XLFN._XLWS.", "_xlws.")
          else Gen.oneOf("_xlfn.", "_XLFN.", "_Xlfn.")
      yield (shape.text, prefix + shape.text)
    forAllNoShrink(genCase) { (bare, stored) =>
      val e1 = assertRoundTrips(s"=$bare")
      val e2 = assertRoundTrips(s"=$stored")
      assertEquals(e2, e1, s"'$stored' should parse to the tree of '$bare'")
      val printed = FormulaPrinter.print(e2)
      assertEquals(printed, FormulaPrinter.print(e1))
      // the call is the whole formula, so a surviving prefix would sit right after the '='
      assert(
        !printed.drop(1).toLowerCase.startsWith("_xl"),
        s"'$stored' should print bare, got '$printed'"
      )
      true
    }
  }

  property("GH-653: LET bindings and bodies round-trip, function-name shadows included") {
    forAllNoShrink(genLet(2, NoScope)) { body =>
      val e1 = assertRoundTrips(s"=$body")
      assert(FormulaPrinter.print(e1).startsWith("=LET("), s"'$body' should print as a LET")
      true
    }
  }

  property(
    "GH-653: string literals with doubled quotes, separators, parens and Unicode print back byte-for-byte"
  ) {
    val shapes: List[Shape] = List(
      x => x,
      x => s"LEN($x)",
      x => s"$x&\"a\"",
      x => s"IF(A1=$x, 1, 0)",
      x => s"COUNTIF(A1:A3, $x)",
      x => s"SUBSTITUTE($x, $x, \"\")",
      x => s"@$x",
      x => s"$x=$x"
    )
    val genCase =
      for
        literal <- genStringLitText
        shape <- Gen.oneOf(shapes)
      yield shape(literal)
    forAllNoShrink(genCase) { body =>
      val source = s"=$body"
      assertPrintsAs(source, source)
      assertRoundTrips(source)
      true
    }
  }

  // ==================== Properties: the grammar the parser rejects ====================

  /** The three spellings of a 3-D span, with the two sheet names it spans. */
  private val genThreeD: Gen[(String, String, String)] =
    for
      s1 <- genSheetNameText
      s2raw <- genSheetNameText
      s2 = if s2raw.equalsIgnoreCase(s1) then s2raw.take(30) + "x" else s2raw
      tail <- Gen.oneOf(genRefText, genRangeText)
      spelling <- Gen.choose(0, 2)
      wrap <- Gen.oneOf[Shape](x => x, x => s"SUM($x)", x => s"$x+1")
    yield
      val q = SheetName.quoteForFormula
      def quoted(name: String) = s"'${name.replace("'", "''")}'"
      val span = spelling match
        case 0 => s"${q(s1)}:${q(s2)}!$tail"
        case 1 => s"${quoted(s1)}:${quoted(s2)}!$tail"
        case _ => s"${quoted(s"$s1:$s2")}!$tail"
      (wrap(span), s1, s2)

  private val genArrayConstant: Gen[String] =
    val element = Gen.frequency(
      3 -> genIntText,
      2 -> genStringLitText,
      1 -> genBoolText,
      1 -> genErrorWritten,
      1 -> Gen.oneOf("-1", "1.5", "2E3")
    )
    for
      rows <- Gen.choose(1, 3)
      cols <- Gen.choose(1, 3)
      cells <- Gen.listOfN(rows * cols, element)
      wrap <- Gen.oneOf[Shape](
        x => x,
        x => s"SUM($x)",
        x => s"INDEX($x, 1, 2)",
        x => s"$x+1",
        x => s"MATCH(1, $x, 0)"
      )
    yield wrap(cells.grouped(cols).map(_.mkString(",")).mkString("{", ";", "}"))

  private val genStructuredRef: Gen[String] =
    for
      table <- Gen.oneOf("Table1", "Sales", "tblData", "DeptSales")
      col <- Gen.oneOf("Col", "Amount", "Col Name", "Q1 Sales", "Total")
      spelling <- Gen.choose(0, 6)
      wrap <- Gen.oneOf[Shape](x => x, x => s"SUM($x)", x => s"$x*2", x => s"INDEX($x, 1)")
    yield
      val ref = spelling match
        case 0 => s"$table[$col]"
        case 1 => s"$table[[#Headers],[$col]]"
        case 2 => s"[@$col]"
        case 3 => s"$table[@[$col]]"
        case 4 => s"$table[#All]"
        case 5 => s"$table[[#Totals],[$col]]"
        case _ => s"$table[@$col]"
      wrap(ref)

  private val genRangeOrRef: Gen[String] = Gen.oneOf(genQualifiedRange, genQualifiedRef)

  private val genUnion: Gen[String] =
    for
      a <- genRangeOrRef
      b <- genRangeOrRef
      wrap <- Gen.oneOf[Shape](x => s"($x)", x => s"SUM(($x))", x => s"COUNT(($x))")
    yield wrap(s"$a,$b")

  private val genIntersection: Gen[String] =
    for
      a <- genRangeOrRef
      b <- genRangeOrRef
      wrap <- Gen.oneOf[Shape](x => x, x => s"SUM($x)")
    yield wrap(s"$a $b")

  private val genLambda: Gen[String] =
    Gen.oneOf(
      "LAMBDA(x,x+1)(2)",
      "LAMBDA(x, y, x+y)",
      "MAP(A1:A3, LAMBDA(v, v*2))",
      "LET(f, LAMBDA(x, x*2), f(3))",
      "BYROW(A1:C3, LAMBDA(r, SUM(r)))",
      "REDUCE(0, A1:A3, LAMBDA(a, b, a+b))",
      "LAMBDA(x, [y], x)(1)"
    )

  private val genRejected: Gen[String] =
    Gen.frequency(
      3 -> genThreeD.map(_._1),
      2 -> genArrayConstant,
      2 -> genStructuredRef,
      1 -> genUnion,
      1 -> genIntersection,
      1 -> genLambda
    )

  /** Total rejection: a Left with a position inside the text, never a throw. */
  private def assertRejected(body: String): ParseError =
    Try(FormulaParser.parse(s"=$body")) match
      case Success(Left(err)) =>
        position(err) match
          case Some(pos) =>
            assert(
              pos >= 0 && pos <= body.length,
              s"'$body': position $pos of $err is off the text"
            )
          case None => fail(s"'$body': $err carries no position")
        err
      case Success(Right(expr)) =>
        fail(s"'$body' should be rejected, parsed as ${FormulaPrinter.printWithTypes(expr)}")
      case Failure(t) => fail(s"'$body' threw ${t.getClass.getSimpleName}: ${t.getMessage}")

  /**
   * The FormulaStorage boundary is the identity on text the parser rejects: the stored form maps
   * back to the model text, is itself a fixpoint, and equals the text exactly when the scanner
   * finds no bare future call in it (the documented `bareFutureCalls` ⇔ `toStored` equivalence).
   */
  private def assertStorageIdentity(body: String): Unit =
    val stored = FormulaStorage.toStored(body)
    assertEquals(FormulaStorage.fromStored(stored), body, s"stored as '$stored'")
    assertEquals(FormulaStorage.toStored(stored), stored, s"toStored is not idempotent on '$body'")
    assertEquals(
      stored == body,
      FormulaStorage.bareFutureCalls(body).isEmpty,
      s"'$body' stored as '$stored' with bare future calls ${FormulaStorage.bareFutureCalls(body)}"
    )

  private def assertShiftRefuses(body: String): Unit =
    FormulaOps.shift(s"=$body", 1, 1) match
      case Left(XLError.FormulaError(_, reason)) =>
        assert(reason.startsWith("Cannot shift references"), s"'$body': $reason")
      case other => fail(s"shift of '$body' should refuse with a FormulaError, got $other")

  property(
    "GH-653: rejected grammar is total — a Left with a position, storage identity, shift refuses"
  ) {
    forAllNoShrink(genRejected) { body =>
      assertRejected(body)
      assertStorageIdentity(body)
      assertShiftRefuses(body)
      true
    }
  }

  property(
    "GH-653: a rename touching either end of a 3-D span refuses instead of leaving an end stale"
  ) {
    forAllNoShrink(genThreeD) { (body, start, end) =>
      assertRejected(body)
      List(start, end).flatMap(SheetName(_).toOption).foreach { from =>
        val to = SheetName.unsafe(if from.value == "Renamed" then "Renamed2" else "Renamed")
        FormulaOps.renameSheet(body, from, to) match
          case Left(XLError.FormulaError(_, reason)) =>
            assert(reason.contains("Cannot rewrite"), s"'$body' renaming '${from.value}': $reason")
          case other =>
            fail(s"renaming '${from.value}' in '$body' should refuse, got $other")
      }
      true
    }
  }

  // ==================== Deterministic sweeps and pins ====================

  test("GH-653: the slot model read off describeParts agrees with every spec's declared arity") {
    callShapes.foreach { (spec, arities) =>
      assert(arities.nonEmpty, s"${spec.name}: no argument count satisfies ${spec.arity}")
      assertEquals(
        arities.foldLeft(Int.MaxValue)(math.min),
        declaredMin(spec.arity),
        s"${spec.name}: the smallest parseable call differs from the declared minimum"
      )
    }
  }

  private def sweepFiller(kind: String): String = kind match
    case "range" => "A1:B2"
    case "number or range" | "array or range" => "A1:A3"
    case "boolean" => "TRUE"
    case "text" => "\"x\""
    case "integer" => "0"
    case "number" => "1"
    case _ => "A1" // date, cell, value

  private def emptySubsets(positions: List[Int]): List[Set[Int]] =
    if positions.size <= 4 then positions.toSet.subsets().toList
    else Set.empty[Int] :: positions.toSet :: positions.map(Set(_))

  test(
    "GH-653: every registry function round-trips at every legal arity with every subset of value slots empty"
  ) {
    callShapes.foreach { (spec, arities) =>
      arities.foreach { n =>
        val kinds = kindsFor(slotsOf(spec), n).getOrElse(Nil)
        val emptiable =
          if n < 2 then Nil
          else kinds.zipWithIndex.collect { case (kind, i) if !isStrictRange(kind) => i }
        emptySubsets(emptiable).foreach { emptied =>
          val args =
            kinds.zipWithIndex.map((kind, i) => if emptied(i) then "" else sweepFiller(kind))
          val shape = CallShape(spec, kinds, emptied, s"${spec.name}(${args.mkString(",")})")
          val e1 = assertRoundTrips(s"=${shape.text}")
          assertEmptiedSlots(shape, e1)
        }
      }
    }
  }

  test(
    "GH-653: every numeric literal spelling the grammar accepts parses and reprints as a fixpoint"
  ) {
    List(
      ".5",
      ".25",
      "5.",
      "1.5E10",
      "2.3E-7",
      "1e3",
      "1E+2",
      "4E0",
      "0.10",
      "1.50",
      "0",
      "1048576"
    )
      .foreach(n => assertRoundTrips(s"=$n"))
    // Excel's leading-dot fraction re-spells through BigDecimal; the value is what round-trips
    assertPrintsAs("=.5", "=0.5")
    assertPrintsAs("=A1+.5", "=A1+0.5")
    assertPrintsAs("=-.5", "=-0.5")
    assertPrintsAs("=SUM(.25, 1)", "=SUM(0.25, 1)")
  }

  test("GH-653: NOT( is the function call — a postfix after its closing paren binds to the call") {
    // before: `NOT(` took the lenient keyword path, so the parens were the operand's grouping
    // parens and the postfix bound inside them — NOT(A1)% read as NOT(A1%), a different value
    assertPrintsAs("=NOT(A1)^2", "=NOT(A1)^2")
    assertPrintsAs("=NOT(A1)%", "=NOT(A1)%")
    assertPrintsAs("=not(TRUE)*2", "=NOT(TRUE)*2")
    assertPrintsAs("=-NOT(TRUE)", "=-NOT(TRUE)")
    assertRejected("NOT(A1)#")
    // whitespace between the name and the paren is still the call, as for every other function
    // (`SUM (A1)`): Excel strips the space, so `NOT (A1)^2` is `(NOT(A1))^2` there too
    assertPrintsAs("=NOT (A1)^2", "=NOT(A1)^2")
    assertPrintsAs("=NOT  (A1)%", "=NOT(A1)%")
    assertPrintsAs("=-not (A1)", "=-NOT(A1)")
    assertEquals(parsed("=NOT (A1)^2"), parsed("=NOT(A1)^2"))
    assertRejected("NOT (A1)#")
    // the paren-less keyword form is xl's lenient prefix operator: its operand is one power term
    assertPrintsAs("=NOT A1^2", "=NOT(A1^2)")
    assertPrintsAs("=NOT A1", "=NOT(A1)")
    assertPrintsAs("=-NOT TRUE", "=-NOT(TRUE)")
    List("=NOT(A1)^2", "=NOT(A1)%", "=NOT (A1)^2", "=NOT A1^2", "=not(TRUE)*2")
      .foreach(assertRoundTrips)
  }

  test("GH-653: the rejected grammar's canonical examples name their failure") {
    def variant(body: String): String = assertRejected(body).productPrefix
    assertEquals(variant("Sheet1:Sheet3!A1"), "InvalidCellRef")
    assertEquals(variant("'Q1':'Q3'!A1"), "UnexpectedChar")
    assertEquals(variant("'Sheet1:Sheet 3'!A1"), "InvalidCellRef")
    assertEquals(variant("{1,2;3,4}"), "UnexpectedChar")
    assertEquals(variant("SUM(Table1[Col])"), "UnexpectedChar")
    assertEquals(variant("Table1[Col]"), "UnexpectedChar")
    assertEquals(variant("[@Col]"), "InvalidCellRef")
    assertEquals(variant("(A1,B1)"), "UnbalancedDelimiter")
    assertEquals(variant("A1:B2 B1:C3"), "UnexpectedChar")
    assertEquals(variant("LAMBDA(x,x+1)(2)"), "UnknownFunction")
    List(
      "Sheet1:Sheet3!A1",
      "{1,2;3,4}",
      "SUM(Table1[Col])",
      "[@Col]",
      "(A1,B1)",
      "A1:B2 B1:C3",
      "LAMBDA(x,x+1)(2)",
      "LET(f, LAMBDA(x, x*2), f(3))"
    ).foreach { body =>
      assertStorageIdentity(body)
      assertShiftRefuses(body)
    }
  }
