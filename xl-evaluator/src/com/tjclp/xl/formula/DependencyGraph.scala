package com.tjclp.xl.formula

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.CellRange
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.sheets.Sheet
import scala.annotation.{nowarn, tailrec}

/**
 * Dependency graph for formula cells.
 *
 * Represents the dependency relationships between cells containing formulas. Each node is a cell
 * reference (ARef), and each directed edge A → B means cell A depends on cell B (A uses B's value).
 *
 * Supports:
 *   - Cycle detection using Tarjan's strongly connected components algorithm
 *   - Topological sorting using Kahn's algorithm for correct evaluation order
 *   - Precedent/dependent queries for impact analysis
 *
 * Design principles:
 *   - Pure functional (no mutations, all operations return new data)
 *   - Total error handling (cycles reported via Either)
 *   - O(V + E) complexity for graph algorithms
 *   - O(1) lookups for precedent/dependent queries (Map-based adjacency lists)
 *
 * Example:
 * {{{
 * // Sheet with formulas: A1="=B1+C1", B1="=10", C1="=20"
 * val graph = DependencyGraph.fromSheet(sheet)
 * graph.precedents(ref"A1") // Set(B1, C1)
 * graph.dependents(ref"B1") // Set(A1)
 *
 * // Detect cycles
 * DependencyGraph.detectCycles(graph) // Right(()) - no cycles
 *
 * // Get evaluation order
 * DependencyGraph.topologicalSort(graph) // Right(List(B1, C1, A1))
 * }}}
 */
final case class DependencyGraph(
  // Forward edges: ref → cells this ref depends on
  dependencies: Map[ARef, Set[ARef]],
  // Reverse edges: ref → cells that depend on this ref
  dependents: Map[ARef, Set[ARef]]
)

object DependencyGraph:
  /**
   * Build dependency graph from Sheet.
   *
   * Iterates through all cells, extracts references from Formula cells, and constructs the
   * dependency graph. Non-formula cells (constants, text, etc.) are ignored.
   *
   * @param sheet
   *   The sheet to analyze
   * @return
   *   Dependency graph with nodes for all formula cells and edges for all references
   *
   * Example:
   * {{{
   * val sheet = Sheet.empty
   *   .put(ref"A1", CellValue.Formula("=B1+C1"))
   *   .put(ref"B1", CellValue.Number(10))
   *   .put(ref"C1", CellValue.Formula("=D1*2"))
   *
   * val graph = DependencyGraph.fromSheet(sheet)
   * // graph.dependencies = Map(A1 -> Set(B1, C1), C1 -> Set(D1))
   * // graph.dependents = Map(B1 -> Set(A1), C1 -> Set(A1), D1 -> Set(C1))
   * }}}
   */
  def fromSheet(sheet: Sheet): DependencyGraph =
    // Get bounds once for all extractions - constrains full column/row ranges
    val bounds = sheet.usedRange

    val formulaCells = sheet.cells.flatMap { case (ref, cell) =>
      cell.value match
        case CellValue.Formula(expression, _) => Some(ref -> expression)
        case _ => None
    }

    // Build forward edges (dependencies) - use bounded extraction to avoid 1M+ cells
    val dependencies = formulaCells.map { case (ref, formulaStr) =>
      val deps = FormulaParser.parse(formulaStr) match
        case scala.util.Right(expr) => extractDependenciesBounded(expr, bounds)
        case scala.util.Left(_) => Set.empty[ARef] // Parse error: no dependencies
      ref -> deps
    }.toMap

    // Build reverse edges (dependents)
    val dependents = dependencies.foldLeft(Map.empty[ARef, Set[ARef]]) { case (acc, (ref, deps)) =>
      deps.foldLeft(acc) { (acc2, dep) =>
        acc2.updated(dep, acc2.getOrElse(dep, Set.empty) + ref)
      }
    }

    DependencyGraph(dependencies, dependents)

  /**
   * Extract all cell references from TExpr.
   *
   * Recursively traverses the expression AST and collects all cell references, including:
   *   - Single cell references (Ref)
   *   - Range references (FoldRange) expanded to all cells in range
   *
   * @param expr
   *   The expression to analyze
   * @return
   *   Set of all cell references used in the expression
   *
   * Example:
   * {{{
   * val expr = TExpr.Add(TExpr.Ref(ref"A1", ...), TExpr.Lit(10))
   * extractDependencies(expr) // Set(A1)
   *
   * val sumExpr = TExpr.sum(CellRange.parse("B1:B10").toOption.get)
   * extractDependencies(sumExpr) // Set(B1, B2, ..., B10)
   * }}}
   */
  // nowarn: Compiler incorrectly reports PolyRef as unreachable, but tests confirm it IS reached at runtime
  @nowarn("msg=Unreachable case")
  def extractDependencies[A](expr: TExpr[A]): Set[ARef] =
    expr match
      // Single cell reference
      case TExpr.Ref(at, _, _) => Set(at)

      // Polymorphic reference (type resolved at evaluation time)
      case TExpr.PolyRef(at, _) => Set(at)

      // Cross-sheet references return Set.empty in same-sheet dependency extraction.
      // This is intentional: extractDependencies builds intra-sheet graphs only.
      // For workbook-level dependency tracking, use extractQualifiedDependencies + fromWorkbook.
      case TExpr.SheetRef(_, _, _, _) => Set.empty
      case TExpr.SheetPolyRef(_, _, _) => Set.empty
      case TExpr.SheetRange(_, _) => Set.empty
      case TExpr.SheetFoldRange(_, _, _, _, _) => Set.empty
      case TExpr.SheetSum(_, _) => Set.empty
      case TExpr.SheetMin(_, _) => Set.empty
      case TExpr.SheetMax(_, _) => Set.empty
      case TExpr.SheetAverage(_, _) => Set.empty
      case TExpr.SheetCount(_, _) => Set.empty

      // Range reference (expand to all cells)
      case TExpr.FoldRange(range, _, _, _) =>
        range.cells.toSet

      // Recursive cases (binary operators)
      case TExpr.Add(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Sub(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Mul(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Div(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Eq(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Neq(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Lt(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Lte(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Gt(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Gte(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.And(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.Or(l, r) => extractDependencies(l) ++ extractDependencies(r)
      case TExpr.ToInt(expr) =>
        extractDependencies(expr) // Type conversion - extract from wrapped expr
      case TExpr.Concatenate(xs) => xs.flatMap(extractDependencies).toSet
      case TExpr.Left(text, n) => extractDependencies(text) ++ extractDependencies(n)
      case TExpr.Right(text, n) => extractDependencies(text) ++ extractDependencies(n)
      case TExpr.Date(y, m, d) =>
        extractDependencies(y) ++ extractDependencies(m) ++ extractDependencies(d)
      case TExpr.Year(date) => extractDependencies(date)
      case TExpr.Month(date) => extractDependencies(date)
      case TExpr.Day(date) => extractDependencies(date)

      // Date calculation functions
      case TExpr.Eomonth(startDate, months) =>
        extractDependencies(startDate) ++ extractDependencies(months)
      case TExpr.Edate(startDate, months) =>
        extractDependencies(startDate) ++ extractDependencies(months)
      case TExpr.Datedif(startDate, endDate, unit) =>
        extractDependencies(startDate) ++ extractDependencies(endDate) ++ extractDependencies(unit)
      case TExpr.Networkdays(startDate, endDate, holidaysOpt) =>
        extractDependencies(startDate) ++
          extractDependencies(endDate) ++
          holidaysOpt.map(_.cells.toSet).getOrElse(Set.empty)
      case TExpr.Workday(startDate, days, holidaysOpt) =>
        extractDependencies(startDate) ++
          extractDependencies(days) ++
          holidaysOpt.map(_.cells.toSet).getOrElse(Set.empty)
      case TExpr.Yearfrac(startDate, endDate, basis) =>
        extractDependencies(startDate) ++ extractDependencies(endDate) ++ extractDependencies(basis)

      // Financial functions
      case TExpr.Npv(rate, values) =>
        extractDependencies(rate) ++ values.localCells
      case TExpr.Irr(values, guessOpt) =>
        values.localCells ++ guessOpt.map(extractDependencies).getOrElse(Set.empty)
      case TExpr.Xnpv(rate, values, dates) =>
        extractDependencies(rate) ++ values.localCells ++ dates.localCells
      case TExpr.Xirr(values, dates, guessOpt) =>
        values.localCells ++ dates.localCells ++ guessOpt
          .map(extractDependencies)
          .getOrElse(Set.empty)
      case TExpr.VLookup(lookup, table, colIndex, rangeLookup) =>
        extractDependencies(lookup) ++
          table.localCells ++
          extractDependencies(colIndex) ++
          extractDependencies(rangeLookup)

      // Conditional aggregation functions
      case TExpr.SumIf(range, criteria, sumRangeOpt) =>
        range.localCells ++
          extractDependencies(criteria) ++
          sumRangeOpt.map(_.localCells).getOrElse(Set.empty)
      case TExpr.CountIf(range, criteria) =>
        range.localCells ++ extractDependencies(criteria)
      case TExpr.SumIfs(sumRange, conditions) =>
        sumRange.localCells ++
          conditions.flatMap { case (range, criteria) =>
            range.localCells ++ extractDependencies(criteria)
          }.toSet
      case TExpr.CountIfs(conditions) =>
        conditions.flatMap { case (range, criteria) =>
          range.localCells ++ extractDependencies(criteria)
        }.toSet

      // Array and advanced lookup functions
      case TExpr.SumProduct(arrays) =>
        arrays.flatMap(_.localCells).toSet

      case TExpr.XLookup(
            lookupValue,
            lookupArray,
            returnArray,
            ifNotFound,
            matchMode,
            searchMode
          ) =>
        extractDependencies(lookupValue) ++
          lookupArray.localCells ++
          returnArray.localCells ++
          ifNotFound.map(extractDependencies).getOrElse(Set.empty) ++
          extractDependencies(matchMode) ++
          extractDependencies(searchMode)

      // Ternary operator
      case TExpr.If(cond, thenBranch, elseBranch) =>
        extractDependencies(cond) ++ extractDependencies(thenBranch) ++ extractDependencies(
          elseBranch
        )

      // Unary operators
      case TExpr.Not(x) => extractDependencies(x)
      case TExpr.Len(x) => extractDependencies(x)
      case TExpr.Upper(x) => extractDependencies(x)
      case TExpr.Lower(x) => extractDependencies(x)

      // Error handling functions
      case TExpr.Iferror(value, valueIfError) =>
        extractDependencies(value) ++ extractDependencies(valueIfError)
      case TExpr.Iserror(value) => extractDependencies(value)

      // Rounding and math functions
      case TExpr.Round(value, numDigits) =>
        extractDependencies(value) ++ extractDependencies(numDigits)
      case TExpr.RoundUp(value, numDigits) =>
        extractDependencies(value) ++ extractDependencies(numDigits)
      case TExpr.RoundDown(value, numDigits) =>
        extractDependencies(value) ++ extractDependencies(numDigits)
      case TExpr.Abs(value) => extractDependencies(value)
      case TExpr.Sqrt(value) => extractDependencies(value)
      case TExpr.Mod(number, divisor) =>
        extractDependencies(number) ++ extractDependencies(divisor)
      case TExpr.Power(number, power) =>
        extractDependencies(number) ++ extractDependencies(power)
      case TExpr.Log(number, base) =>
        extractDependencies(number) ++ extractDependencies(base)
      case TExpr.Ln(value) => extractDependencies(value)
      case TExpr.Exp(value) => extractDependencies(value)
      case TExpr.Floor(number, significance) =>
        extractDependencies(number) ++ extractDependencies(significance)
      case TExpr.Ceiling(number, significance) =>
        extractDependencies(number) ++ extractDependencies(significance)
      case TExpr.Trunc(number, numDigits) =>
        extractDependencies(number) ++ extractDependencies(numDigits)
      case TExpr.Sign(value) => extractDependencies(value)
      case TExpr.Int_(value) => extractDependencies(value)

      // Lookup functions
      case TExpr.Index(array, rowNum, colNum) =>
        array.localCells ++ extractDependencies(rowNum) ++ colNum
          .map(extractDependencies)
          .getOrElse(Set.empty)
      case TExpr.Match(lookupValue, lookupArray, matchType) =>
        extractDependencies(lookupValue) ++ lookupArray.localCells ++ extractDependencies(
          matchType
        )

      // Range aggregate functions (direct enum cases)
      case TExpr.Sum(range) => range.localCells
      case TExpr.Count(range) => range.localCells
      case TExpr.Min(range) => range.localCells
      case TExpr.Max(range) => range.localCells
      case TExpr.Average(range) => range.localCells

      // Unified aggregate function (typeclass-based)
      case TExpr.Aggregate(_, location) => location.localCells

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty
      case TExpr.Today() => Set.empty
      case TExpr.Now() => Set.empty
      case TExpr.Pi() => Set.empty

      // Date-to-serial converters - extract from inner expression
      case TExpr.DateToSerial(dateExpr) => extractDependencies(dateExpr)
      case TExpr.DateTimeToSerial(dtExpr) => extractDependencies(dtExpr)

  /**
   * Extract all cell references from TExpr, bounded by the sheet's used range.
   *
   * This optimized version constrains full column/row references (like A:A or 1:1) to the
   * intersection with bounds, avoiding iteration over 1M+ cells. Use this when building dependency
   * graphs from sheets.
   *
   * @param expr
   *   The expression to analyze
   * @param bounds
   *   Optional bounding range (typically sheet.usedRange) to constrain full ranges
   * @return
   *   Set of all cell references used in the expression, bounded by the given range
   */
  @nowarn("msg=Unreachable case")
  def extractDependenciesBounded[A](expr: TExpr[A], bounds: Option[CellRange]): Set[ARef] =
    // Helper to bound a CellRange
    def boundRange(range: CellRange): Set[ARef] =
      bounds match
        case Some(b) => range.intersect(b).map(_.cells.toSet).getOrElse(Set.empty)
        case None => range.cells.toSet

    expr match
      // Single cell reference
      case TExpr.Ref(at, _, _) => Set(at)

      // Polymorphic reference (type resolved at evaluation time)
      case TExpr.PolyRef(at, _) => Set(at)

      // Cross-sheet references return Set.empty in same-sheet dependency extraction.
      case TExpr.SheetRef(_, _, _, _) => Set.empty
      case TExpr.SheetPolyRef(_, _, _) => Set.empty
      case TExpr.SheetRange(_, _) => Set.empty
      case TExpr.SheetFoldRange(_, _, _, _, _) => Set.empty
      case TExpr.SheetSum(_, _) => Set.empty
      case TExpr.SheetMin(_, _) => Set.empty
      case TExpr.SheetMax(_, _) => Set.empty
      case TExpr.SheetAverage(_, _) => Set.empty
      case TExpr.SheetCount(_, _) => Set.empty

      // Range reference (expand to all cells, BOUNDED)
      case TExpr.FoldRange(range, _, _, _) => boundRange(range)

      // Recursive cases (binary operators)
      case TExpr.Add(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Sub(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Mul(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Div(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Eq(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Neq(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Lt(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Lte(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Gt(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Gte(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.And(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.Or(l, r) =>
        extractDependenciesBounded(l, bounds) ++ extractDependenciesBounded(r, bounds)
      case TExpr.ToInt(expr) => extractDependenciesBounded(expr, bounds)
      case TExpr.Concatenate(xs) => xs.flatMap(extractDependenciesBounded(_, bounds)).toSet
      case TExpr.Left(text, n) =>
        extractDependenciesBounded(text, bounds) ++ extractDependenciesBounded(n, bounds)
      case TExpr.Right(text, n) =>
        extractDependenciesBounded(text, bounds) ++ extractDependenciesBounded(n, bounds)
      case TExpr.Date(y, m, d) =>
        extractDependenciesBounded(y, bounds) ++ extractDependenciesBounded(m, bounds) ++
          extractDependenciesBounded(d, bounds)
      case TExpr.Year(date) => extractDependenciesBounded(date, bounds)
      case TExpr.Month(date) => extractDependenciesBounded(date, bounds)
      case TExpr.Day(date) => extractDependenciesBounded(date, bounds)

      // Date calculation functions
      case TExpr.Eomonth(startDate, months) =>
        extractDependenciesBounded(startDate, bounds) ++ extractDependenciesBounded(months, bounds)
      case TExpr.Edate(startDate, months) =>
        extractDependenciesBounded(startDate, bounds) ++ extractDependenciesBounded(months, bounds)
      case TExpr.Datedif(startDate, endDate, unit) =>
        extractDependenciesBounded(startDate, bounds) ++
          extractDependenciesBounded(endDate, bounds) ++
          extractDependenciesBounded(unit, bounds)
      case TExpr.Networkdays(startDate, endDate, holidaysOpt) =>
        extractDependenciesBounded(startDate, bounds) ++
          extractDependenciesBounded(endDate, bounds) ++
          holidaysOpt.map(boundRange).getOrElse(Set.empty)
      case TExpr.Workday(startDate, days, holidaysOpt) =>
        extractDependenciesBounded(startDate, bounds) ++
          extractDependenciesBounded(days, bounds) ++
          holidaysOpt.map(boundRange).getOrElse(Set.empty)
      case TExpr.Yearfrac(startDate, endDate, basis) =>
        extractDependenciesBounded(startDate, bounds) ++
          extractDependenciesBounded(endDate, bounds) ++
          extractDependenciesBounded(basis, bounds)

      // Financial functions
      case TExpr.Npv(rate, values) =>
        extractDependenciesBounded(rate, bounds) ++ values.localCellsBounded(bounds)
      case TExpr.Irr(values, guessOpt) =>
        values.localCellsBounded(bounds) ++
          guessOpt.map(extractDependenciesBounded(_, bounds)).getOrElse(Set.empty)
      case TExpr.Xnpv(rate, values, dates) =>
        extractDependenciesBounded(rate, bounds) ++
          values.localCellsBounded(bounds) ++
          dates.localCellsBounded(bounds)
      case TExpr.Xirr(values, dates, guessOpt) =>
        values.localCellsBounded(bounds) ++
          dates.localCellsBounded(bounds) ++
          guessOpt.map(extractDependenciesBounded(_, bounds)).getOrElse(Set.empty)
      case TExpr.VLookup(lookup, table, colIndex, rangeLookup) =>
        extractDependenciesBounded(lookup, bounds) ++
          table.localCellsBounded(bounds) ++
          extractDependenciesBounded(colIndex, bounds) ++
          extractDependenciesBounded(rangeLookup, bounds)

      // Conditional aggregation functions
      case TExpr.SumIf(range, criteria, sumRangeOpt) =>
        range.localCellsBounded(bounds) ++
          extractDependenciesBounded(criteria, bounds) ++
          sumRangeOpt.map(_.localCellsBounded(bounds)).getOrElse(Set.empty)
      case TExpr.CountIf(range, criteria) =>
        range.localCellsBounded(bounds) ++ extractDependenciesBounded(criteria, bounds)
      case TExpr.SumIfs(sumRange, conditions) =>
        sumRange.localCellsBounded(bounds) ++
          conditions.flatMap { case (range, criteria) =>
            range.localCellsBounded(bounds) ++ extractDependenciesBounded(criteria, bounds)
          }.toSet
      case TExpr.CountIfs(conditions) =>
        conditions.flatMap { case (range, criteria) =>
          range.localCellsBounded(bounds) ++ extractDependenciesBounded(criteria, bounds)
        }.toSet

      // Array and advanced lookup functions
      case TExpr.SumProduct(arrays) =>
        arrays.flatMap(_.localCellsBounded(bounds)).toSet

      case TExpr.XLookup(
            lookupValue,
            lookupArray,
            returnArray,
            ifNotFound,
            matchMode,
            searchMode
          ) =>
        extractDependenciesBounded(lookupValue, bounds) ++
          lookupArray.localCellsBounded(bounds) ++
          returnArray.localCellsBounded(bounds) ++
          ifNotFound.map(extractDependenciesBounded(_, bounds)).getOrElse(Set.empty) ++
          extractDependenciesBounded(matchMode, bounds) ++
          extractDependenciesBounded(searchMode, bounds)

      // Ternary operator
      case TExpr.If(cond, thenBranch, elseBranch) =>
        extractDependenciesBounded(cond, bounds) ++
          extractDependenciesBounded(thenBranch, bounds) ++
          extractDependenciesBounded(elseBranch, bounds)

      // Unary operators
      case TExpr.Not(x) => extractDependenciesBounded(x, bounds)
      case TExpr.Len(x) => extractDependenciesBounded(x, bounds)
      case TExpr.Upper(x) => extractDependenciesBounded(x, bounds)
      case TExpr.Lower(x) => extractDependenciesBounded(x, bounds)

      // Error handling functions
      case TExpr.Iferror(value, valueIfError) =>
        extractDependenciesBounded(value, bounds) ++
          extractDependenciesBounded(valueIfError, bounds)
      case TExpr.Iserror(value) => extractDependenciesBounded(value, bounds)

      // Rounding and math functions
      case TExpr.Round(value, numDigits) =>
        extractDependenciesBounded(value, bounds) ++ extractDependenciesBounded(numDigits, bounds)
      case TExpr.RoundUp(value, numDigits) =>
        extractDependenciesBounded(value, bounds) ++ extractDependenciesBounded(numDigits, bounds)
      case TExpr.RoundDown(value, numDigits) =>
        extractDependenciesBounded(value, bounds) ++ extractDependenciesBounded(numDigits, bounds)
      case TExpr.Abs(value) => extractDependenciesBounded(value, bounds)
      case TExpr.Sqrt(value) => extractDependenciesBounded(value, bounds)
      case TExpr.Mod(number, divisor) =>
        extractDependenciesBounded(number, bounds) ++ extractDependenciesBounded(divisor, bounds)
      case TExpr.Power(number, power) =>
        extractDependenciesBounded(number, bounds) ++ extractDependenciesBounded(power, bounds)
      case TExpr.Log(number, base) =>
        extractDependenciesBounded(number, bounds) ++ extractDependenciesBounded(base, bounds)
      case TExpr.Ln(value) => extractDependenciesBounded(value, bounds)
      case TExpr.Exp(value) => extractDependenciesBounded(value, bounds)
      case TExpr.Floor(number, significance) =>
        extractDependenciesBounded(number, bounds) ++ extractDependenciesBounded(
          significance,
          bounds
        )
      case TExpr.Ceiling(number, significance) =>
        extractDependenciesBounded(number, bounds) ++ extractDependenciesBounded(
          significance,
          bounds
        )
      case TExpr.Trunc(number, numDigits) =>
        extractDependenciesBounded(number, bounds) ++ extractDependenciesBounded(numDigits, bounds)
      case TExpr.Sign(value) => extractDependenciesBounded(value, bounds)
      case TExpr.Int_(value) => extractDependenciesBounded(value, bounds)

      // Lookup functions
      case TExpr.Index(array, rowNum, colNum) =>
        array.localCellsBounded(bounds) ++
          extractDependenciesBounded(rowNum, bounds) ++
          colNum.map(extractDependenciesBounded(_, bounds)).getOrElse(Set.empty)
      case TExpr.Match(lookupValue, lookupArray, matchType) =>
        extractDependenciesBounded(lookupValue, bounds) ++
          lookupArray.localCellsBounded(bounds) ++
          extractDependenciesBounded(matchType, bounds)

      // Range aggregate functions (direct enum cases)
      case TExpr.Sum(range) => range.localCellsBounded(bounds)
      case TExpr.Count(range) => range.localCellsBounded(bounds)
      case TExpr.Min(range) => range.localCellsBounded(bounds)
      case TExpr.Max(range) => range.localCellsBounded(bounds)
      case TExpr.Average(range) => range.localCellsBounded(bounds)

      // Unified aggregate function (typeclass-based)
      case TExpr.Aggregate(_, location) => location.localCellsBounded(bounds)

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty
      case TExpr.Today() => Set.empty
      case TExpr.Now() => Set.empty
      case TExpr.Pi() => Set.empty

      // Date-to-serial converters - extract from inner expression
      case TExpr.DateToSerial(dateExpr) => extractDependenciesBounded(dateExpr, bounds)
      case TExpr.DateTimeToSerial(dtExpr) => extractDependenciesBounded(dtExpr, bounds)

  /**
   * Get cells this cell depends on (precedents).
   *
   * Returns the set of cells whose values are used in calculating this cell's value. If the cell
   * has no formula or is not in the graph, returns empty set.
   *
   * @param graph
   *   The dependency graph
   * @param ref
   *   The cell reference to query
   * @return
   *   Set of cells this cell depends on (may be empty)
   *
   * Example:
   * {{{
   * // A1 = "=B1+C1"
   * precedents(graph, ref"A1") // Set(B1, C1)
   * precedents(graph, ref"B1") // Set() - B1 is a constant
   * }}}
   */
  def precedents(graph: DependencyGraph, ref: ARef): Set[ARef] =
    graph.dependencies.getOrElse(ref, Set.empty)

  /**
   * Get cells that depend on this cell (dependents).
   *
   * Returns the set of cells that use this cell's value in their calculations. If no cells depend
   * on this cell, returns empty set.
   *
   * @param graph
   *   The dependency graph
   * @param ref
   *   The cell reference to query
   * @return
   *   Set of cells that depend on this cell (may be empty)
   *
   * Example:
   * {{{
   * // A1 = "=B1+C1", D1 = "=B1*2"
   * dependents(graph, ref"B1") // Set(A1, D1)
   * dependents(graph, ref"A1") // Set() - nothing depends on A1
   * }}}
   */
  def dependents(graph: DependencyGraph, ref: ARef): Set[ARef] =
    graph.dependents.getOrElse(ref, Set.empty)

  /**
   * Compute transitive dependencies for a set of cells.
   *
   * Given a set of starting cells, returns all cells that are directly or transitively depended
   * upon. This is useful for targeted evaluation - to evaluate only formulas in a range, we need to
   * also evaluate all cells they depend on (recursively).
   *
   * @param graph
   *   The dependency graph to traverse
   * @param refs
   *   The starting cell references
   * @return
   *   Set of all cells that the starting cells depend on (transitively)
   *
   * Example:
   * {{{
   * // A1="=B1+C1", B1="=D1", C1="=10", D1="=20"
   * // transitiveDependencies(graph, Set(A1))
   * // Returns: Set(B1, C1, D1) - all cells A1 depends on directly or indirectly
   * }}}
   */
  @scala.annotation.tailrec
  def transitiveDependencies(
    graph: DependencyGraph,
    refs: Set[ARef],
    visited: Set[ARef] = Set.empty
  ): Set[ARef] =
    val toVisit = refs -- visited
    if toVisit.isEmpty then visited
    else
      // Get direct dependencies of all cells in toVisit
      val directDeps = toVisit.flatMap(ref => graph.dependencies.getOrElse(ref, Set.empty))
      // Recurse with direct deps as new frontier
      transitiveDependencies(graph, directDeps, visited ++ toVisit)

  /**
   * Detect circular references using Tarjan's strongly connected components algorithm.
   *
   * A circular reference occurs when a cell's formula depends (directly or transitively) on its own
   * value. For example: A1="=B1", B1="=C1", C1="=A1" forms a cycle.
   *
   * This uses Tarjan's SCC algorithm which runs in O(V + E) time with a single DFS traversal. The
   * algorithm maintains a stack and low-link values to detect strongly connected components
   * (cycles).
   *
   * @param graph
   *   The dependency graph to analyze
   * @return
   *   Left(CircularRef) if a cycle is detected (includes cycle path), Right(()) if acyclic
   *
   * Example:
   * {{{
   * // No cycle: A1="=10", B1="=A1+5"
   * detectCycles(graph) // Right(())
   *
   * // Cycle: A1="=B1", B1="=A1"
   * detectCycles(graph) // Left(EvalError.CircularRef(List(A1, B1, A1)))
   * }}}
   */
  @SuppressWarnings(
    Array(
      "org.wartremover.warts.Var",
      "org.wartremover.warts.IterableOps",
      "org.wartremover.warts.Return",
      "org.wartremover.warts.IsInstanceOf",
      "org.wartremover.warts.AsInstanceOf"
    )
  )
  def detectCycles(graph: DependencyGraph): Either[EvalError.CircularRef, Unit] =
    // Tarjan's SCC algorithm: Intentional imperative implementation
    // Rationale: Classic algorithm uses mutable state for O(V+E) performance.
    // Functional version sacrifices clarity without benefit. Compile-time only.
    var index = 0
    var stack = List.empty[ARef]
    var indices = Map.empty[ARef, Int]
    var lowLinks = Map.empty[ARef, Int]
    var onStack = Set.empty[ARef]

    def strongConnect(v: ARef): Option[List[ARef]] =
      // Set the depth index for v
      indices = indices.updated(v, index)
      lowLinks = lowLinks.updated(v, index)
      index += 1
      stack = v :: stack
      onStack = onStack + v

      // Consider successors of v (cells that v depends on)
      val successors = graph.dependencies.getOrElse(v, Set.empty)
      val cycleFound = successors.foldLeft(Option.empty[List[ARef]]) { (acc, w) =>
        acc match
          case Some(cycle) => Some(cycle) // Already found cycle, propagate
          case None =>
            if !indices.contains(w) then
              // Successor w has not yet been visited; recurse on it
              strongConnect(w) match
                case Some(cycle) => Some(cycle)
                case None =>
                  lowLinks = lowLinks.updated(v, math.min(lowLinks(v), lowLinks(w)))
                  None
            else if onStack.contains(w) then
              // Successor w is on stack and hence in the current SCC
              lowLinks = lowLinks.updated(v, math.min(lowLinks(v), indices(w)))
              // Found a cycle! Reconstruct it from stack (w to top of stack forms the cycle)
              val cycleNodes = (stack.takeWhile(_ != w) :+ w).reverse
              Some(cycleNodes :+ cycleNodes.head) // Add first node again to show cycle closes
            else None // w is not on stack, already processed
      }

      cycleFound match
        case Some(cycle) => Some(cycle)
        case None =>
          // If v is a root node, pop the stack and check for SCC
          if lowLinks(v) == indices(v) then
            // Pop nodes from stack until v
            val (scc, remaining) = stack.span(_ != v)
            stack = remaining.tail // Remove v from stack
            onStack = onStack -- (scc :+ v)

            // Check if SCC has more than one node (cycle)
            if scc.nonEmpty then
              // Multiple nodes in SCC means cycle
              val cycleNodes = (scc :+ v).reverse
              Some(cycleNodes :+ cycleNodes.head) // Add first node again to show cycle closes
            else
              // Single node - only a cycle if it has self-loop
              if graph.dependencies.get(v).exists(_.contains(v)) then
                Some(List(v, v)) // Self-loop: v -> v
              else None
          else None

    // Run Tarjan's algorithm on all unvisited nodes
    val allNodes = graph.dependencies.keySet
    val cycleFound = allNodes.foldLeft(Option.empty[List[ARef]]) { (acc, node) =>
      acc match
        case Some(cycle) => Some(cycle) // Already found cycle
        case None =>
          if !indices.contains(node) then strongConnect(node)
          else None
    }

    cycleFound match
      case Some(cycle) => scala.util.Left(EvalError.CircularRef(cycle))
      case None => scala.util.Right(())

  /**
   * Topological sort using Kahn's algorithm.
   *
   * Returns a linear ordering of cells such that for every dependency A → B, cell B appears before
   * cell A in the ordering. This ensures formulas are evaluated in the correct order (dependencies
   * before dependents).
   *
   * Uses Kahn's algorithm which runs in O(V + E) time. The algorithm maintains a queue of nodes
   * with in-degree 0 and processes them in order, removing edges as it goes.
   *
   * @param graph
   *   The dependency graph to sort
   * @return
   *   Left(CircularRef) if a cycle is detected, Right(evaluation order) if acyclic
   *
   * Example:
   * {{{
   * // A1="=B1+C1", B1="=10", C1="=20"
   * topologicalSort(graph) // Right(List(B1, C1, A1))
   *
   * // A1="=B1", B1="=A1" (cycle)
   * topologicalSort(graph) // Left(EvalError.CircularRef(List(A1, B1, A1)))
   * }}}
   */
  def topologicalSort(graph: DependencyGraph): Either[EvalError.CircularRef, List[ARef]] =
    import scala.util.boundary, boundary.break

    boundary:
      // All formula cells (only process formulas, not constants)
      val allNodes = graph.dependencies.keySet

      // If no formula cells, early exit
      if allNodes.isEmpty then break(scala.util.Right(List.empty[ARef]))

      // Calculate in-degree for each node (number of formula cells it depends on)
      // Only count dependencies on other formula cells, not constants
      val inDegree = allNodes.map { node =>
        val deps = graph.dependencies.getOrElse(node, Set.empty)
        val formulaDeps = deps.filter(allNodes.contains)
        node -> formulaDeps.size
      }.toMap

      // Start with nodes that have in-degree 0 (no dependencies)
      val initialQueue = allNodes.filter(node => inDegree(node) == 0).toList

      @tailrec
      def process(
        queue: List[ARef],
        processedInDegree: Map[ARef, Int],
        acc: List[ARef]
      ): (List[ARef], Map[ARef, Int]) =
        queue match
          case Nil => (acc, processedInDegree)
          case node :: rest =>
            val deps = graph.dependents.getOrElse(node, Set.empty).filter(allNodes.contains)
            val (nextInDegree, newlyZero) = deps.foldLeft((processedInDegree, List.empty[ARef])) {
              case ((degreeAcc, zeros), dep) =>
                val newInDegree = degreeAcc(dep) - 1
                val updatedDegree = degreeAcc.updated(dep, newInDegree)
                val nextZeros = if newInDegree == 0 then zeros :+ dep else zeros
                (updatedDegree, nextZeros)
            }
            process(rest ++ newlyZero, nextInDegree, acc :+ node)

      val (result, _) = process(initialQueue, inDegree, List.empty)

      // If all nodes are processed, graph is acyclic
      if result.size == allNodes.size then scala.util.Right(result)
      else
        // Cycle detected: find one cycle for error reporting
        val remainingNodes = allNodes -- result.toSet
        val cycle = remainingNodes.headOption match
          case Some(start) =>
            // Follow dependencies to reconstruct cycle
            def findCycle(current: ARef, visited: Set[ARef]): List[ARef] =
              if visited.contains(current) then
                // Found cycle
                List(current)
              else
                graph.dependencies.getOrElse(current, Set.empty).headOption match
                  case Some(next) if remainingNodes.contains(next) =>
                    current :: findCycle(next, visited + current)
                  case _ => List(current)

            val cyclePath = findCycle(start, Set.empty)
            cyclePath.headOption.map(first => cyclePath :+ first).getOrElse(List.empty)
          case None => List.empty

        scala.util.Left(EvalError.CircularRef(cycle))

  // ===== Cross-Sheet Dependency Tracking =====

  /**
   * Cell reference qualified with sheet name for cross-sheet tracking.
   *
   * Used to track dependencies across sheets within a workbook. Each QualifiedRef uniquely
   * identifies a cell in the workbook by combining the sheet name and cell reference.
   *
   * Example:
   * {{{
   * val ref = QualifiedRef(SheetName.unsafe("Sales"), ref"A1")
   * // Represents Sales!A1
   * }}}
   */
  final case class QualifiedRef(sheet: SheetName, ref: ARef):
    override def toString: String = s"${sheet.value}!${ref.toA1}"

  /**
   * Build dependency graph from Workbook (cross-sheet aware).
   *
   * Iterates through all sheets and cells, extracting references from Formula cells and
   * constructing a workbook-level dependency graph. Cross-sheet references are properly tracked
   * using QualifiedRef.
   *
   * @param workbook
   *   The workbook to analyze
   * @return
   *   Dependency graph with QualifiedRef nodes covering all sheets
   *
   * Example:
   * {{{
   * // Sheet1!A1 = "=Sheet2!B1", Sheet2!B1 = 10
   * val graph = DependencyGraph.fromWorkbook(workbook)
   * // graph contains: QualifiedRef(Sheet1, A1) -> Set(QualifiedRef(Sheet2, B1))
   * }}}
   */
  def fromWorkbook(
    workbook: com.tjclp.xl.workbooks.Workbook
  ): Map[QualifiedRef, Set[QualifiedRef]] =
    workbook.sheets.flatMap { sheet =>
      sheet.cells.flatMap { case (cellRef, cell) =>
        cell.value match
          case CellValue.Formula(expression, _) =>
            val deps = FormulaParser.parse(expression) match
              case scala.util.Right(expr) => extractQualifiedDependencies(expr, sheet.name)
              case scala.util.Left(_) => Set.empty[QualifiedRef]
            Some(QualifiedRef(sheet.name, cellRef) -> deps)
          case _ => None
      }
    }.toMap

  /**
   * Convert a RangeLocation to qualified cell references.
   *
   * For Local ranges, uses the current sheet. For CrossSheet ranges, uses the specified sheet.
   */
  private def locationToQualifiedRefs(
    location: TExpr.RangeLocation,
    currentSheet: SheetName
  ): Set[QualifiedRef] =
    location match
      case TExpr.RangeLocation.Local(range) =>
        range.cells.map(ref => QualifiedRef(currentSheet, ref)).toSet
      case TExpr.RangeLocation.CrossSheet(sheet, range) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet

  /**
   * Extract all qualified cell references from TExpr.
   *
   * Similar to extractDependencies but returns QualifiedRef to track cross-sheet references.
   * Same-sheet references are qualified with the current sheet name.
   *
   * @param expr
   *   The expression to analyze
   * @param currentSheet
   *   The sheet containing the formula (used for same-sheet ref qualification)
   * @return
   *   Set of qualified cell references used in the expression
   */
  @nowarn("msg=Unreachable case")
  private def extractQualifiedDependencies[A](
    expr: TExpr[A],
    currentSheet: SheetName
  ): Set[QualifiedRef] =
    expr match
      // Same-sheet references - qualify with current sheet
      case TExpr.Ref(at, _, _) => Set(QualifiedRef(currentSheet, at))
      case TExpr.PolyRef(at, _) => Set(QualifiedRef(currentSheet, at))
      case TExpr.FoldRange(range, _, _, _) =>
        range.cells.map(ref => QualifiedRef(currentSheet, ref)).toSet

      // Cross-sheet references - use target sheet
      case TExpr.SheetRef(sheet, at, _, _) => Set(QualifiedRef(sheet, at))
      case TExpr.SheetPolyRef(sheet, at, _) => Set(QualifiedRef(sheet, at))
      case TExpr.SheetRange(sheet, range) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet
      case TExpr.SheetFoldRange(sheet, range, _, _, _) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet
      case TExpr.SheetSum(sheet, range) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet
      case TExpr.SheetMin(sheet, range) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet
      case TExpr.SheetMax(sheet, range) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet
      case TExpr.SheetAverage(sheet, range) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet
      case TExpr.SheetCount(sheet, range) =>
        range.cells.map(ref => QualifiedRef(sheet, ref)).toSet

      // Recursive cases (binary operators)
      case TExpr.Add(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Sub(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Mul(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Div(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Eq(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Neq(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Lt(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Lte(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Gt(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Gte(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.And(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )
      case TExpr.Or(l, r) =>
        extractQualifiedDependencies(l, currentSheet) ++ extractQualifiedDependencies(
          r,
          currentSheet
        )

      // Unary operators
      case TExpr.Not(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.Len(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.Upper(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.Lower(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.ToInt(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.Abs(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.Sqrt(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.Mod(n, d) =>
        extractQualifiedDependencies(n, currentSheet) ++ extractQualifiedDependencies(
          d,
          currentSheet
        )
      case TExpr.Power(n, p) =>
        extractQualifiedDependencies(n, currentSheet) ++ extractQualifiedDependencies(
          p,
          currentSheet
        )
      case TExpr.Log(n, b) =>
        extractQualifiedDependencies(n, currentSheet) ++ extractQualifiedDependencies(
          b,
          currentSheet
        )
      case TExpr.Ln(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.Exp(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.Floor(n, s) =>
        extractQualifiedDependencies(n, currentSheet) ++ extractQualifiedDependencies(
          s,
          currentSheet
        )
      case TExpr.Ceiling(n, s) =>
        extractQualifiedDependencies(n, currentSheet) ++ extractQualifiedDependencies(
          s,
          currentSheet
        )
      case TExpr.Trunc(n, d) =>
        extractQualifiedDependencies(n, currentSheet) ++ extractQualifiedDependencies(
          d,
          currentSheet
        )
      case TExpr.Sign(x) => extractQualifiedDependencies(x, currentSheet)
      case TExpr.Int_(x) => extractQualifiedDependencies(x, currentSheet)

      // Ternary
      case TExpr.If(cond, thenBranch, elseBranch) =>
        extractQualifiedDependencies(cond, currentSheet) ++
          extractQualifiedDependencies(thenBranch, currentSheet) ++
          extractQualifiedDependencies(elseBranch, currentSheet)

      // Text functions
      case TExpr.Concatenate(xs) =>
        xs.flatMap(extractQualifiedDependencies(_, currentSheet)).toSet
      case TExpr.Left(text, n) =>
        extractQualifiedDependencies(text, currentSheet) ++ extractQualifiedDependencies(
          n,
          currentSheet
        )
      case TExpr.Right(text, n) =>
        extractQualifiedDependencies(text, currentSheet) ++ extractQualifiedDependencies(
          n,
          currentSheet
        )

      // Date functions
      case TExpr.Date(y, m, d) =>
        extractQualifiedDependencies(y, currentSheet) ++
          extractQualifiedDependencies(m, currentSheet) ++
          extractQualifiedDependencies(d, currentSheet)
      case TExpr.Year(date) => extractQualifiedDependencies(date, currentSheet)
      case TExpr.Month(date) => extractQualifiedDependencies(date, currentSheet)
      case TExpr.Day(date) => extractQualifiedDependencies(date, currentSheet)

      // Range functions (direct, not FoldRange)
      case TExpr.Sum(range) => locationToQualifiedRefs(range, currentSheet)
      case TExpr.Count(range) => locationToQualifiedRefs(range, currentSheet)
      case TExpr.Min(range) => locationToQualifiedRefs(range, currentSheet)
      case TExpr.Max(range) => locationToQualifiedRefs(range, currentSheet)
      case TExpr.Average(range) => locationToQualifiedRefs(range, currentSheet)

      // Unified aggregate function (typeclass-based)
      case TExpr.Aggregate(_, location) => locationToQualifiedRefs(location, currentSheet)

      // Literals and nullary functions (no dependencies)
      case TExpr.Lit(_) => Set.empty
      case TExpr.Today() => Set.empty
      case TExpr.Now() => Set.empty
      case TExpr.Pi() => Set.empty

      // Date-to-serial converters - extract from inner expression
      case TExpr.DateToSerial(dateExpr) => extractQualifiedDependencies(dateExpr, currentSheet)
      case TExpr.DateTimeToSerial(dtExpr) => extractQualifiedDependencies(dtExpr, currentSheet)

      // Date calculation functions
      case TExpr.Eomonth(startDate, months) =>
        extractQualifiedDependencies(startDate, currentSheet) ++
          extractQualifiedDependencies(months, currentSheet)
      case TExpr.Edate(startDate, months) =>
        extractQualifiedDependencies(startDate, currentSheet) ++
          extractQualifiedDependencies(months, currentSheet)
      case TExpr.Datedif(startDate, endDate, unit) =>
        extractQualifiedDependencies(startDate, currentSheet) ++
          extractQualifiedDependencies(endDate, currentSheet) ++
          extractQualifiedDependencies(unit, currentSheet)
      case TExpr.Networkdays(startDate, endDate, holidaysOpt) =>
        extractQualifiedDependencies(startDate, currentSheet) ++
          extractQualifiedDependencies(endDate, currentSheet) ++
          holidaysOpt
            .map(_.cells.map(ref => QualifiedRef(currentSheet, ref)).toSet)
            .getOrElse(Set.empty)
      case TExpr.Workday(startDate, days, holidaysOpt) =>
        extractQualifiedDependencies(startDate, currentSheet) ++
          extractQualifiedDependencies(days, currentSheet) ++
          holidaysOpt
            .map(_.cells.map(ref => QualifiedRef(currentSheet, ref)).toSet)
            .getOrElse(Set.empty)
      case TExpr.Yearfrac(startDate, endDate, basis) =>
        extractQualifiedDependencies(startDate, currentSheet) ++
          extractQualifiedDependencies(endDate, currentSheet) ++
          extractQualifiedDependencies(basis, currentSheet)

      // Financial functions
      case TExpr.Npv(rate, values) =>
        extractQualifiedDependencies(rate, currentSheet) ++
          locationToQualifiedRefs(values, currentSheet)
      case TExpr.Irr(values, guessOpt) =>
        locationToQualifiedRefs(values, currentSheet) ++
          guessOpt.map(extractQualifiedDependencies(_, currentSheet)).getOrElse(Set.empty)
      case TExpr.Xnpv(rate, values, dates) =>
        extractQualifiedDependencies(rate, currentSheet) ++
          locationToQualifiedRefs(values, currentSheet) ++
          locationToQualifiedRefs(dates, currentSheet)
      case TExpr.Xirr(values, dates, guessOpt) =>
        locationToQualifiedRefs(values, currentSheet) ++
          locationToQualifiedRefs(dates, currentSheet) ++
          guessOpt.map(extractQualifiedDependencies(_, currentSheet)).getOrElse(Set.empty)
      case TExpr.VLookup(lookup, table, colIndex, rangeLookup) =>
        extractQualifiedDependencies(lookup, currentSheet) ++
          locationToQualifiedRefs(table, currentSheet) ++
          extractQualifiedDependencies(colIndex, currentSheet) ++
          extractQualifiedDependencies(rangeLookup, currentSheet)

      // Conditional aggregation functions
      case TExpr.SumIf(range, criteria, sumRangeOpt) =>
        locationToQualifiedRefs(range, currentSheet) ++
          extractQualifiedDependencies(criteria, currentSheet) ++
          sumRangeOpt
            .map(locationToQualifiedRefs(_, currentSheet))
            .getOrElse(Set.empty)
      case TExpr.CountIf(range, criteria) =>
        locationToQualifiedRefs(range, currentSheet) ++
          extractQualifiedDependencies(criteria, currentSheet)
      case TExpr.SumIfs(sumRange, conditions) =>
        locationToQualifiedRefs(sumRange, currentSheet) ++
          conditions.flatMap { case (range, criteria) =>
            locationToQualifiedRefs(range, currentSheet) ++
              extractQualifiedDependencies(criteria, currentSheet)
          }.toSet
      case TExpr.CountIfs(conditions) =>
        conditions.flatMap { case (range, criteria) =>
          locationToQualifiedRefs(range, currentSheet) ++
            extractQualifiedDependencies(criteria, currentSheet)
        }.toSet
      case TExpr.SumProduct(arrays) =>
        arrays.flatMap(locationToQualifiedRefs(_, currentSheet)).toSet

      // Advanced lookup functions
      case TExpr.XLookup(
            lookupValue,
            lookupArray,
            returnArray,
            ifNotFound,
            matchMode,
            searchMode
          ) =>
        extractQualifiedDependencies(lookupValue, currentSheet) ++
          locationToQualifiedRefs(lookupArray, currentSheet) ++
          locationToQualifiedRefs(returnArray, currentSheet) ++
          ifNotFound.map(extractQualifiedDependencies(_, currentSheet)).getOrElse(Set.empty) ++
          extractQualifiedDependencies(matchMode, currentSheet) ++
          extractQualifiedDependencies(searchMode, currentSheet)
      case TExpr.Index(array, rowNum, colNum) =>
        locationToQualifiedRefs(array, currentSheet) ++
          extractQualifiedDependencies(rowNum, currentSheet) ++
          colNum.map(extractQualifiedDependencies(_, currentSheet)).getOrElse(Set.empty)
      case TExpr.Match(lookupValue, lookupArray, matchType) =>
        extractQualifiedDependencies(lookupValue, currentSheet) ++
          locationToQualifiedRefs(lookupArray, currentSheet) ++
          extractQualifiedDependencies(matchType, currentSheet)

      // Error handling functions
      case TExpr.Iferror(value, valueIfError) =>
        extractQualifiedDependencies(value, currentSheet) ++
          extractQualifiedDependencies(valueIfError, currentSheet)
      case TExpr.Iserror(value) =>
        extractQualifiedDependencies(value, currentSheet)

      // Rounding and math functions
      case TExpr.Round(value, numDigits) =>
        extractQualifiedDependencies(value, currentSheet) ++
          extractQualifiedDependencies(numDigits, currentSheet)
      case TExpr.RoundUp(value, numDigits) =>
        extractQualifiedDependencies(value, currentSheet) ++
          extractQualifiedDependencies(numDigits, currentSheet)
      case TExpr.RoundDown(value, numDigits) =>
        extractQualifiedDependencies(value, currentSheet) ++
          extractQualifiedDependencies(numDigits, currentSheet)

  /**
   * Detect circular references across sheets using Tarjan's SCC algorithm.
   *
   * Similar to detectCycles but works with QualifiedRef to detect cycles that span multiple sheets.
   * A cross-sheet cycle occurs when cells across different sheets form a circular dependency (e.g.,
   * Sheet1!A1 = Sheet2!B1, Sheet2!B1 = Sheet1!A1).
   *
   * @param graph
   *   The cross-sheet dependency graph from fromWorkbook
   * @return
   *   Left(CircularRef) if cycle detected, Right(()) if acyclic
   *
   * Example:
   * {{{
   * val graph = DependencyGraph.fromWorkbook(workbook)
   * DependencyGraph.detectCrossSheetCycles(graph) match
   *   case Left(err) => println(s"Circular reference: $err")
   *   case Right(_) => println("No cycles")
   * }}}
   */
  @SuppressWarnings(
    Array(
      "org.wartremover.warts.Var",
      "org.wartremover.warts.IterableOps",
      "org.wartremover.warts.Return",
      "org.wartremover.warts.IsInstanceOf",
      "org.wartremover.warts.AsInstanceOf"
    )
  )
  def detectCrossSheetCycles(
    graph: Map[QualifiedRef, Set[QualifiedRef]]
  ): Either[EvalError.CircularRef, Unit] =
    // Tarjan's SCC algorithm adapted for QualifiedRef
    var index = 0
    var stack = List.empty[QualifiedRef]
    var indices = Map.empty[QualifiedRef, Int]
    var lowLinks = Map.empty[QualifiedRef, Int]
    var onStack = Set.empty[QualifiedRef]

    def strongConnect(v: QualifiedRef): Option[List[ARef]] =
      indices = indices.updated(v, index)
      lowLinks = lowLinks.updated(v, index)
      index += 1
      stack = v :: stack
      onStack = onStack + v

      val successors = graph.getOrElse(v, Set.empty)
      val cycleFound = successors.foldLeft(Option.empty[List[ARef]]) { (acc, w) =>
        acc match
          case Some(cycle) => Some(cycle)
          case None =>
            if !indices.contains(w) then
              strongConnect(w) match
                case Some(cycle) => Some(cycle)
                case None =>
                  lowLinks = lowLinks.updated(v, math.min(lowLinks(v), lowLinks(w)))
                  None
            else if onStack.contains(w) then
              lowLinks = lowLinks.updated(v, math.min(lowLinks(v), indices(w)))
              // Found cycle - reconstruct from stack
              val cycleNodes = (stack.takeWhile(_ != w) :+ w).reverse
              // Convert to List[ARef] with sheet prefix for error message
              Some(cycleNodes.map(_.ref) :+ cycleNodes.head.ref)
            else None
      }

      cycleFound match
        case Some(cycle) => Some(cycle)
        case None =>
          if lowLinks(v) == indices(v) then
            val (scc, remaining) = stack.span(_ != v)
            stack = remaining.tail
            onStack = onStack -- (scc :+ v)

            if scc.nonEmpty then
              // Multiple nodes in SCC = cycle
              val cycleNodes = (scc :+ v).reverse
              Some(cycleNodes.map(_.ref) :+ cycleNodes.head.ref)
            else if graph.get(v).exists(_.contains(v)) then Some(List(v.ref, v.ref)) // Self-loop
            else None
          else None

    val allNodes = graph.keySet
    val cycleFound = allNodes.foldLeft(Option.empty[List[ARef]]) { (acc, node) =>
      acc match
        case Some(cycle) => Some(cycle)
        case None =>
          if !indices.contains(node) then strongConnect(node)
          else None
    }

    cycleFound match
      case Some(cycle) => scala.util.Left(EvalError.CircularRef(cycle))
      case None => scala.util.Right(())
