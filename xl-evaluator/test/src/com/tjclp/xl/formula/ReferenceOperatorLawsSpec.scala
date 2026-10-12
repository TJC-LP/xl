package com.tjclp.xl.formula

import com.tjclp.xl.{*, given}
import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*
import org.scalacheck.rng.Seed

/**
 * GH-669: the laws the union and intersection operators obey over static areas.
 *
 *   - Union desugaring: an aggregate over `(a,b,…)` equals the aggregate over `a,b,…`.
 *   - Intersection substitution: `a b` with a non-empty intersection `R` behaves as `R` in every
 *     context — a plain cell, array mode, SUM, arithmetic, INDEX, ROWS, `@`, AREAS; an empty one is
 *     `#NULL!`.
 *   - Dependencies: a union depends on its areas, an intersection on the cells it shares.
 *   - Shifting a union shifts each of its leaves.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
class ReferenceOperatorLawsSpec extends ScalaCheckSuite:

  override def scalaCheckTestParameters =
    super.scalaCheckTestParameters.withMinSuccessfulTests(60)

  /** A1:F10 of numbers, text and blanks (a deterministic mix), so every aggregate has cases. */
  private val sheet: Sheet =
    (for col <- 0 until 6; row <- 0 until 10 yield (col, row)).foldLeft(Sheet("S")) {
      case (acc, (col, row)) =>
        val at = ARef.from0(col, row)
        (col + row) % 7 match
          case 0 => acc
          case 3 => acc.put(at, CellValue.Text(s"t$col$row"))
          case _ => acc.put(at, CellValue.Number(BigDecimal(col * 10 + row + 1)))
    }
  private val book = Workbook(Vector(sheet))

  private val genArea: Gen[CellRange] =
    for
      c1 <- Gen.choose(0, 5)
      c2 <- Gen.choose(0, 5)
      r1 <- Gen.choose(0, 9)
      r2 <- Gen.choose(0, 9)
    yield CellRange(ARef.from0(c1 min c2, r1 min r2), ARef.from0(c1 max c2, r1 max r2))

  private val genAreas: Gen[List[CellRange]] = Gen.choose(2, 4).flatMap(Gen.listOfN(_, genArea))

  /** A plain cell's row: inside the grid's rows or past them (where intersections fail). */
  private val genHost: Gen[ARef] = Gen.choose(0, 14).map(row => ARef.from0(7, row))

  private def plain(at: ARef, formula: String): Either[String, CellValue] =
    val placed = sheet.put(at, CellValue.Formula(formula, None))
    placed.evaluateCell(at, Clock.system, Some(book.put(placed))).left.map(_.message)

  private def array(formula: String): Either[String, Any] =
    FormulaParser
      .parse(s"=$formula")
      .left
      .map(_.toString)
      .flatMap(expr =>
        Evaluator.arrayInstance
          .eval(expr.asInstanceOf[TExpr[Any]], sheet, workbook = Some(book))
          .left
          .map(e => EvalError.toErrorValue(e).fold(e.toString)(_.toExcel))
      )

  private def parsed(formula: String): TExpr[?] =
    FormulaParser.parse(formula).fold(err => fail(s"$formula: $err"), identity)

  private val aggregates =
    List(
      "SUM",
      "COUNT",
      "COUNTA",
      "AVERAGE",
      "MIN",
      "MAX",
      "MEDIAN",
      "STDEV",
      "STDEVP",
      "VAR",
      "VARP"
    )

  property("union desugaring: f((a,b,…)) == f(a,b,…) for every variadic aggregate") {
    forAll(genAreas, genHost, Gen.oneOf(aggregates)) { (areas, host, fn) =>
      val list = areas.map(_.toA1).mkString(",")
      assertEquals(plain(host, s"$fn(($list))"), plain(host, s"$fn($list)"), s"$fn(($list))")
      assertEquals(array(s"$fn(($list))"), array(s"$fn($list)"), s"evala $fn(($list))")
    }
  }

  property("intersection substitution: C[a b] == C[R]; an empty intersection is #NULL!") {
    val contexts: List[String => String] = List(
      x => x,
      x => s"SUM($x)",
      x => s"($x)*2",
      x => s"INDEX($x,1,1)",
      x => s"ROWS($x)",
      x => s"COLUMNS($x)",
      x => s"@($x)",
      x => s"AREAS($x)",
      x => s"COUNTA($x)"
    )
    forAll(genArea, genArea, genHost) { (a, b, host) =>
      val op = s"${a.toA1} ${b.toA1}"
      a.intersect(b) match
        case Some(shared) =>
          contexts.foreach { context =>
            val written = context(op)
            val substituted = context(shared.toA1)
            assertEquals(plain(host, written), plain(host, substituted), s"plain $written")
            assertEquals(array(written), array(substituted), s"evala $written")
          }
        case None =>
          assertEquals(plain(host, op), Right(CellValue.Error(CellError.Null)), op)
          assertEquals(plain(host, s"ISERROR($op)"), Right(CellValue.Bool(true)), op)
      true
    }
  }

  property("dependencies: a union's are its areas', an intersection's the shared cells'") {
    forAll(genAreas, genArea, genArea) { (areas, a, b) =>
      val list = areas.map(_.toA1).mkString(",")
      assertEquals(
        DependencyGraph.extractDependencies(parsed(s"=SUM(($list))")),
        DependencyGraph.extractDependencies(parsed(s"=SUM($list)"))
      )
      val shared = a.intersect(b).fold(Set.empty[ARef])(_.cells.toSet)
      assertEquals(
        DependencyGraph.extractDependencies(parsed(s"=SUM(${a.toA1} ${b.toA1})")),
        shared
      )
    }
  }

  // ===== GH-713: the range operator `:` =====

  private def hull(a: CellRange, b: CellRange): CellRange = a.expand(b.start).expand(b.end)

  /**
   * The ways an operand of `:` spells an area: as written, or as the reference IF, CHOOSE, OFFSET
   * or INDEX returns; a single cell also as one INDEX picks out of the whole grid.
   */
  private def spellings(area: CellRange): List[String] =
    val a = area.toA1
    val cell =
      if area.width == 1 && area.height == 1 then
        List(s"INDEX($$A$$1:$$F$$10,${area.start.row.index1},${area.start.col.index0 + 1})")
      else Nil
    List(a, s"IF(TRUE,$a,0)", s"CHOOSE(1,$a)", s"OFFSET($a,0,0)", s"INDEX($a,0,0)") ++ cell

  private val genSingleOrArea: Gen[CellRange] =
    Gen.frequency(
      3 -> genArea,
      1 -> genArea.map(area => CellRange(area.start, area.start))
    )

  private def genSpelled(area: CellRange): Gen[String] = Gen.oneOf(spellings(area))

  property("bounding substitution: C[a:b] == C[bbox(a,b)] whatever spells a and b") {
    val contexts: List[String => String] = List(
      x => x,
      x => s"SUM($x)",
      x => s"($x)*2",
      x => s"INDEX($x,1,1)",
      x => s"ROWS($x)",
      x => s"COLUMNS($x)",
      x => s"@($x)",
      x => s"AREAS($x)",
      x => s"COUNTA($x)"
    )
    val genCase =
      for
        a <- genSingleOrArea
        b <- genSingleOrArea
        sa <- genSpelled(a)
        sb <- genSpelled(b)
        host <- genHost
      yield (a, b, s"$sa:$sb", host)
    forAll(genCase) { (a, b, op, host) =>
      val box = hull(a, b).toA1
      contexts.foreach { context =>
        val written = context(op)
        val substituted = context(box)
        assertEquals(plain(host, written), plain(host, substituted), s"plain $written")
        assertEquals(array(written), array(substituted), s"evala $written")
      }
      true
    }
  }

  property("semilattice: a:b == b:a, (a:b):c == a:(b:c), a:a == a, a:(a b) == a") {
    val genCase =
      for
        a <- genArea
        b <- genArea
        c <- genArea
        sa <- genSpelled(a)
        sb <- genSpelled(b)
        sc <- genSpelled(c)
      yield (a, b, sa, sb, sc)
    forAll(genCase) { (a, b, sa, sb, sc) =>
      val host = ARef.from0(7, 0)
      def same(x: String, y: String): Unit =
        List((t: String) => s"SUM($t)", (t: String) => s"ROWS($t)*100+COLUMNS($t)").foreach { f =>
          assertEquals(plain(host, f(x)), plain(host, f(y)), s"${f(x)} vs ${f(y)}")
        }
      same(s"$sa:$sb", s"$sb:$sa")
      same(s"($sa:$sb):$sc", s"$sa:($sb:$sc)")
      same(s"$sa:$sa", a.toA1)
      if a.intersect(b).isDefined then same(s"$sa:(${a.toA1} ${b.toA1})", a.toA1)
      true
    }
  }

  property("edges: a computed range's edges hold every cell it reads, or it is dynamic") {
    val genCase =
      for
        a <- genSingleOrArea
        b <- genSingleOrArea
        sa <- genSpelled(a)
        sb <- genSpelled(b)
      yield (a, b, s"SUM($sa:$sb)")
    forAll(genCase) { (a, b, formula) =>
      val expr = parsed(s"=$formula")
      val read = hull(a, b).cells.toSet
      assert(
        DependencyGraph.containsDynamicReference(expr) ||
          read.subsetOf(DependencyGraph.extractDependencies(expr)),
        s"$formula reads $read, edges ${DependencyGraph.extractDependencies(expr)}"
      )
      // the pre-filter admits every formula the classifier flags
      if DependencyGraph.containsDynamicReference(expr) then
        assert(functions.ReferenceOperators.mayContainComputedRange(formula), formula)
      true
    }
  }

  property("static operands: a range's edges are exactly its bounding range") {
    forAll(genArea, genArea) { (a, b) =>
      assertEquals(
        DependencyGraph.extractDependencies(parsed(s"=SUM(${a.toA1}:${b.toA1})")),
        hull(a, b).cells.toSet
      )
    }
  }

  property("targeted recalculation after any one edit equals full recalculation") {
    val genCase =
      for
        formulas <- Gen.listOfN(
          4,
          for
            a <- genSingleOrArea
            b <- genSingleOrArea
            sa <- genSpelled(a)
            sb <- genSpelled(b)
          yield s"SUM($sa:$sb)"
        )
        col <- Gen.choose(0, 5)
        row <- Gen.choose(0, 9)
        value <- Gen.choose(-50, 50)
      yield (formulas, ARef.from0(col, row), value)
    forAll(genCase) { (formulas, edited, value) =>
      val withFormulas = formulas.zipWithIndex.foldLeft(sheet) { case (acc, (f, i)) =>
        acc.put(ARef.from0(7, i), CellValue.Formula(f, None))
      }
      val total = withFormulas.put(ARef.from0(8, 0), CellValue.Formula("SUM(H1:H4)", None))
      val settled = Workbook(Vector(total)).recalculate().workbook
      val name = SheetName.unsafe("S")
      val edit = settled(name)
        .fold(e => fail(e.message), identity)
        .put(edited, CellValue.Number(BigDecimal(value)))
      val book = settled.put(edit)
      val targeted = book.recalculateAfterEdit(name, Set(edited), RecalcOptions()).workbook
      val full = book.recalculate().workbook
      def caches(wb: Workbook): Map[ARef, Option[CellValue]] =
        wb(name).fold(e => fail(e.message), identity).cells.collect {
          case (at, cell) if cell.value.isInstanceOf[CellValue.Formula] =>
            at -> (cell.value match
              case CellValue.Formula(_, cached, _) => cached
              case _ => None)
        }
      assertEquals(
        caches(targeted),
        caches(full),
        s"${formulas.mkString("; ")} after ${edited.toA1}"
      )
      true
    }
  }

  // ===== GH-713: printing `:` — the structural parenthesization rule, against the parser =====

  /**
   * The operands a `:` chain is built from, as written: every shape the printer tells apart — bare
   * cells and names (one spelled like a column, `Jan`), qualified cells and names, lexical and
   * qualified ranges (whole columns and rows too), external references, error literals, calls, a
   * union and an intersection, and names using every token character (`.`, `_`, `\`). Inside the
   * LET context `x` and `n.b_1` are LET names bound to cells and `r` one bound to a range.
   */
  private val chainLeaves: List[String] = List(
    "A1",
    "$B$2",
    "C3",
    "Start",
    "Finish",
    "Jan",
    "Sheet1!A1",
    "'My Sheet'!B2",
    "Sheet1!Start",
    "Other!Jan",
    // every token character the readers share (ReferenceChars.isToken): `.`, `_`, `\`, `$`
    "Sales.Total",
    "_rate",
    "\\path",
    "A$1",
    "$A1",
    "Sheet1!Sales.Total",
    "'My Sheet'!_rate",
    "Other!\\path",
    "Sheet1!$A$1:B$2",
    "B2:C3",
    "A:A",
    "1:3",
    "$A$1:$B$2",
    "Sheet1!B2:C3",
    "'My Sheet'!A:A",
    "Other!1:2",
    "[1]Book!A1",
    "[1]Book!A1:B2",
    "#REF!",
    "Sheet1!#REF!",
    "INDEX(A:A,3)",
    "OFFSET(A1,1,1)",
    "XLOOKUP(2,A1:A3,B1:B3)",
    "_xlfn.XLOOKUP(2,A1:A3,B1:B3)",
    "INDEX(Sheet1!A:A,3)",
    "IF(TRUE,A1,B2)",
    "INDIRECT(\"A1\")",
    "CHOOSE(1,A1,B2)",
    "(A1,B2)",
    "(A1:C3 B2:D4)"
  )

  private val letPrefix = "LET(x,A1,r,B1:B3,n.b_1,C2,"

  /** A `:` tree over the leaves, each node and leaf grouped at random. */
  private def genChain(depth: Int, leaves: List[String]): Gen[String] =
    val leaf = Gen.oneOf(leaves)
    def grouped(text: String): Gen[String] = Gen.frequency(3 -> text, 1 -> s"($text)")
    if depth <= 0 then leaf.flatMap(grouped)
    else
      Gen.frequency(
        1 -> leaf.flatMap(grouped),
        3 -> (for
          l <- genChain(depth - 1, leaves)
          r <- genChain(depth - 1, leaves)
          node <- grouped(s"$l:$r")
        yield node)
      )

  /** A whole formula: a `:` tree in a reference position, in or out of a LET. */
  private val genChainFormula: Gen[(String, Boolean)] =
    for
      inLet <- Gen.oneOf(false, true)
      depth <- Gen.choose(1, 5)
      tree <- genChain(depth, if inLet then "x" :: "r" :: "n.b_1" :: chainLeaves else chainLeaves)
      consumer <- Gen.oneOf("SUM", "ROWS", "COUNTA")
    yield if inLet then (s"=$letPrefix$consumer($tree))", true) else (s"=$consumer($tree)", false)

  /** Every range-operator call in `expr` (call arguments, LET bindings and bodies, coercions). */
  private def rangeCalls(expr: TExpr[?]): List[TExpr.Call[?]] =
    def operandCalls(value: functions.ArgValue): List[TExpr.Call[?]] = value match
      case functions.ArgValue.Expr(e) => rangeCalls(e)
      case _ => Nil
    expr match
      case TExpr.Coerced(inner, _) => rangeCalls(inner)
      case TExpr.Let(bindings, body) =>
        bindings.flatMap((_, value) => rangeCalls(value)) ++ rangeCalls(body)
      case call: TExpr.Call[?] =>
        val own =
          if call.spec.name == functions.ReferenceOperators.RangeName then List(call) else Nil
        own ++ call.spec.argSpec.toValues(call.args).flatMap(operandCalls)
      case _ => Nil

  /**
   * The parse-probing printer the structural rule replaced, kept as the oracle: for `l:r`, the
   * first of `l:r`, `l:(r)`, `(l):r` that the parser reads back as this very call, else `(l):(r)`.
   * Unlike the old printer it reads each candidate in the formula's LET scope, so a LET name reads
   * as one.
   */
  private def oracleRange(call: TExpr.Call[?], inLet: Boolean): String =
    val (l, r) = functions.ReferenceOperators.RangeOperatorCall.operands(call)
    def text(op: functions.ReferenceOperators.Operand): String =
      op.fold(FormulaPrinter.formatLocation, e => FormulaPrinter.printFileForm(e))
    val (left, right) = (text(l), text(r))
    def reads(candidate: String): Option[TExpr.Call[?]] =
      val formula = if inLet then s"=$letPrefix$candidate)" else s"=$candidate"
      FormulaParser.parse(formula).toOption.flatMap {
        case TExpr.Let(_, body) if inLet =>
          functions.ReferenceOperators.RangeOperatorCall.unapply(body)
        case other if !inLet => functions.ReferenceOperators.RangeOperatorCall.unapply(other)
        case _ => None
      }
    List(s"$left:$right", s"$left:($right)", s"($left):$right")
      .find(candidate => reads(candidate).contains(call))
      .getOrElse(s"($left):($right)")

  property("GH-713: parse(print(t)) == t for `:` chains, with the parser's own choice of parens") {
    // ten formulas a case: the suite's 60 cases are 600 trees, about half of which parse
    forAllNoShrink(Gen.listOfN(10, genChainFormula)) { cases =>
      cases.foreach { (formula, inLet) =>
        FormulaParser.parse(formula) match
          case Left(_) => () // a tree the parser refuses (two bare cells fold, 3-D ends, …)
          case Right(expr) =>
            assertEquals(FormulaParser.parse(FormulaPrinter.print(expr)), Right(expr), formula)
            assertEquals(
              FormulaParser.parse("=" + FormulaPrinter.printFileForm(expr)),
              Right(expr),
              formula
            )
            // minimal, in the old printer's order: each `:` prints as the first candidate that
            // the parser reads back as itself
            rangeCalls(expr).foreach { call =>
              assertEquals(FormulaPrinter.printFileForm(call), oracleRange(call, inLet), formula)
            }
            // the dynamic pre-filter admits every formula whose computed range the classifier
            // flags, written or printed
            val dynamicRange = rangeCalls(expr).exists(call =>
              functions.ReferenceOperators
                .isComputedRangeDynamic(call, functions.ReferenceOperators.staticShape)
            )
            if dynamicRange then
              assert(functions.ReferenceOperators.mayContainComputedRange(formula), formula)
              assert(
                functions.ReferenceOperators.mayContainComputedRange(
                  FormulaPrinter.printFileForm(expr)
                ),
                formula
              )
      }
      true
    }
  }

  test("GH-713: the chain generator reaches the parser (the property above is not vacuous)") {
    val params = Gen.Parameters.default
    val samples = (0L until 400L).flatMap(i => genChainFormula(params, Seed(i)))
    def calls(formula: String): Int =
      FormulaParser.parse(formula).fold(_ => 0, expr => rangeCalls(expr).size)
    val parsedCount = samples.count((formula, _) => FormulaParser.parse(formula).isRight)
    val withRange = samples.count((formula, _) => calls(formula) > 0)
    assert(parsedCount >= samples.size / 4, s"$parsedCount of ${samples.size} parse")
    assert(withRange >= samples.size / 5, s"$withRange of ${samples.size} hold a `:` call")
  }

  property("GH-713: a long `:` chain parses or hits a nesting guard — never overflows") {
    // short operands any neighbour accepts, so a chain past both guards (128 levels, 1024
    // operators) fits the 8192-character limit (`x` only where it is a LET name: elsewhere it is
    // a name spelled like column X, which a token's `:` refuses)
    val operands = List("A1", "INDEX(B1,1)", "Start", "B2:C3", "(A1,B2)", "#REF!", "IF(1,A1,B1)")
    val genLong =
      for
        // under the nesting budget (each `:` takes a level) and far past it
        n <- Gen.frequency(1 -> Gen.choose(2, 60), 1 -> Gen.choose(61, 1200))
        inLet <- Gen.oneOf(false, true)
        first <- Gen.oneOf(operands)
        rest <- Gen.listOfN(n - 1, Gen.oneOf(if inLet then "x" :: operands else operands))
        nested <- Gen.oneOf(false, true)
      yield
        val all = first :: rest
        val chain =
          if nested then
            all.dropRight(1).foldRight(all.lastOption.getOrElse(first))((p, acc) => s"$p:($acc)")
          else all.mkString(":")
        if inLet then s"=LET(x,A1,SUM($chain))" else s"=SUM($chain)"
    forAllNoShrink(genLong) { formula =>
      FormulaParser.parse(formula) match
        case Right(expr) =>
          assertEquals(FormulaParser.parse(FormulaPrinter.print(expr)), Right(expr), formula)
        case Left(
              _: ParseError.NestingTooDeep | _: ParseError.TooManyOperators |
              _: ParseError.FormulaTooLong
            ) =>
          ()
        case Left(other) => fail(s"${formula.take(80)}…: $other")
      true
    }
  }

  test("drift: referenceFunctions is every Referencing spec, the selectors and the operators") {
    val referencing = functions.FunctionRegistry.all.collect {
      case spec: functions.FunctionSpec.Referencing[?, ?] => spec.name
    }.toSet
    // the static names the parser consults are exactly the registry's reference-computing specs:
    // a new FunctionSpec.Referencing fails here until ReferenceFunctionNames.computing lists it
    assertEquals(functions.ReferenceFunctionNames.computing, referencing)
    assertEquals(
      Evaluator.referenceFunctions,
      referencing ++ Set("IF", "IFS", "CHOOSE", "SWITCH") ++
        Set(functions.ReferenceOperators.IntersectionName, functions.ReferenceOperators.RangeName)
    )
    assertEquals(
      functions.ReferenceOperators.selectingFunctions,
      Set("IF", "IFS", "CHOOSE", "SWITCH")
    )
    // a range slot names the calls that compute a reference: every Referencing spec and `:` —
    // not the selectors (IF(…) there keeps the plain argument-count diagnostic) nor the
    // intersection (a variadic slot reads it; a range slot reports it like any other call)
    assertEquals(
      functions.ArgSpec.computedReferenceCalls,
      Evaluator.referenceFunctions -- functions.ReferenceOperators.selectingFunctions -
        functions.ReferenceOperators.IntersectionName
    )
    assertEquals(
      functions.ArgSpec.computedReferenceCalls,
      referencing + functions.ReferenceOperators.RangeName
    )
    // neither operator is registered, nor reachable by name
    List(
      functions.ReferenceOperators.UnionName,
      functions.ReferenceOperators.IntersectionName,
      functions.ReferenceOperators.RangeName
    ).foreach(op => assertEquals(functions.FunctionRegistry.lookup(op), None, op))
  }

  property("shifting a union shifts each leaf; an intersection shifts both operands") {
    forAll(genAreas, Gen.choose(0, 4), Gen.choose(0, 4)) { (areas, dc, dr) =>
      def shifted(formula: String): String =
        FormulaPrinter.print(FormulaShifter.shift(parsed(formula), dc, dr))
      val leaves = areas.map(area => shifted(s"=${area.toA1}").stripPrefix("="))
      assertEquals(
        shifted(s"=SUM((${areas.map(_.toA1).mkString(",")}))"),
        s"=SUM((${leaves.mkString(",")}))"
      )
      val (first, second) = (areas(0), areas(1))
      assertEquals(
        shifted(s"=${first.toA1} ${second.toA1}"),
        s"=${leaves(0)} ${leaves(1)}"
      )
    }
  }

  property("GH-713: shifting a range shifts both operands, a computed one inside its call") {
    forAll(genArea, genArea, Gen.choose(0, 4), Gen.choose(0, 4)) { (a, b, dc, dr) =>
      def shifted(formula: String): String =
        FormulaPrinter.printFileForm(FormulaShifter.shift(parsed(formula), dc, dr))
      val (la, lb) = (shifted(s"=${a.toA1}"), shifted(s"=${b.toA1}"))
      assertEquals(shifted(s"=SUM(${a.toA1}:INDEX(${b.toA1},1,1))"), s"SUM($la:INDEX($lb,1,1))")
      assertEquals(shifted(s"=SUM(INDEX(${a.toA1},1,1):${b.toA1})"), s"SUM(INDEX($la,1,1):$lb)")
    }
  }
