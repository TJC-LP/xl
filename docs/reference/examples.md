
# Examples — End‑to‑End Snippets

## 1) Export a simple table
```scala
import com.tjclp.xl.api.Excel
import com.tjclp.xl.core.*, com.tjclp.xl.core.addr.*
import com.tjclp.xl.dsl.cell

final case class Person(name: String, age: Int, email: String) derives CanEqual
val people = Vector(Person("Ada", 34, "ada@ex.com"), Person("Linus", 55, "linus@ex.com"))

val s0 = Sheet(addr.SheetName("People"), Map.empty, Set.empty, Map.empty, Map.empty)
val s1 = s0.put(cell"A1" -> "Name", cell"B1" -> "Age", cell"C1" -> "Email")
val s2 = people.zipWithIndex.foldLeft(s1) { case (sh, (p, i)) =>
  val r = addr.Row.from0(i + 1)
  sh.updateCell(ARef(Column.from0(0), r))(_ => Cell(ARef(Column.from0(0), r), CellValue.Text(p.name)))
    .updateCell(ARef(Column.from0(1), r))(_ => Cell(ARef(Column.from0(1), r), CellValue.Number(BigDecimal(p.age))))
    .updateCell(ARef(Column.from0(2), r))(_ => Cell(ARef(Column.from0(2), r), CellValue.Text(p.email)))
}
val wb = Workbook(Vector(s2), styles = (), sharedStrings = (), metadata = ())
// Excel[IO].write(wb, Paths.get("people.xlsx"))
```

## 2) Formula Parsing & Typed AST

```scala
import com.tjclp.xl.*
import com.tjclp.xl.formula.{FormulaParser, FormulaPrinter, TExpr, ParseError}

// ==================== Parsing Formula Strings ====================

// Parse basic formulas
val sum = FormulaParser.parse("=SUM(A2:A10)")
// Right(TExpr.FoldRange(CellRange("A2:A10"), BigDecimal(0), sumFunc, decodeNumeric))

val conditional = FormulaParser.parse("=IF(A1>100, \"High\", \"Low\")")
// Right(TExpr.If(TExpr.Gt(...), TExpr.Lit("High"), TExpr.Lit("Low")))

val arithmetic = FormulaParser.parse("=(A1+B1)*C1/D1")
// Right(TExpr.Div(TExpr.Mul(TExpr.Add(...), ...), ...))

// Scientific notation supported
val scientific = FormulaParser.parse("=1.5E10")
// Right(TExpr.Lit(BigDecimal("1.5E10")))

// ==================== Programmatic Construction ====================

// Build formulas with type safety (GADT prevents type mixing)
val avgFormula: TExpr[BigDecimal] = TExpr.Div(
  TExpr.sum(CellRange.parse("A2:A10").toOption.get),
  TExpr.Lit(BigDecimal(9))
)

val nestedIf: TExpr[String] = TExpr.If(
  TExpr.Gt(
    TExpr.Ref(ref"A1", TExpr.decodeNumeric),
    TExpr.Lit(BigDecimal(100))
  ),
  TExpr.Lit("High"),
  TExpr.If(
    TExpr.Gt(
      TExpr.Ref(ref"A1", TExpr.decodeNumeric),
      TExpr.Lit(BigDecimal(50))
    ),
    TExpr.Lit("Medium"),
    TExpr.Lit("Low")
  )
)

// ==================== Round-Trip Verification ====================

// Print back to Excel syntax
val formulaString = FormulaPrinter.print(avgFormula, includeEquals = true)
// "=SUM(A2:A10)/9"

// Verify parse → print → parse = identity
FormulaParser.parse(formulaString) match
  case Right(reparsed) =>
    assert(FormulaPrinter.print(reparsed) == formulaString)
  case Left(error) =>
    println(s"Parse error: $error")

// ==================== Error Handling ====================

FormulaParser.parse("=SUMM(A1:A10)") match
  case Right(expr) => // Success
  case Left(ParseError.UnknownFunction(name, pos, suggestions)) =>
    println(s"Unknown function '$name' at position $pos")
    println(s"Did you mean: ${suggestions.mkString(", ")}")  // "SUM"
  case Left(other) =>
    println(s"Parse error: $other")

// ==================== Integration with CellValue ====================

// Store parsed formula in cell
val parsedFormula = FormulaParser.parse("=A1+B1")
parsedFormula.foreach { expr =>
  val formulaStr = FormulaPrinter.print(expr, includeEquals = false)
  val cell = Cell(ref"C1", CellValue.Formula(formulaStr))
  // Cell contains: CellValue.Formula("A1+B1")
}
```

**Note**: Formula evaluation is now **fully operational** (WI-07, WI-08, WI-09 complete). See runnable example scripts in `/examples/` directory.

## 3) Formula System Example Scripts

The `/examples/` directory contains ready-to-run scripts demonstrating the complete formula system (parsing, evaluation, dependency analysis, and cycle detection).

### Quick Start (5 minutes)

**File**: `examples/quick-start.sc`

Get up and running with formula evaluation in 5 minutes. Demonstrates:
- Parse Excel formulas to typed AST
- Evaluate formulas against sheets
- Handle errors gracefully (Either types, no exceptions)
- Detect circular references automatically
- Evaluate all formulas safely with dependency checking

```bash
scala-cli examples/quick-start.sc
```

**Key APIs shown**:
- `FormulaParser.parse("=SUM(A1:A10)")`
- `sheet.evaluateFormula(formula)`
- `sheet.evaluateCell(ref)`
- `sheet.evaluateWithDependencyCheck()`
- `DependencyGraph.detectCycles(graph)`

### Financial Modeling (~200 LOC)

**File**: `examples/financial-model.sc`

Complete 3-year income statement with:
- Revenue projections with growth rates
- COGS, operating expenses, EBITDA, net income
- Financial ratios (gross margin, operating margin, net margin)
- Dependency chain analysis (evaluation order)

**Perfect for**: FP&A teams, finance applications, reporting automation

**Demonstrates**:
- Complex formula chains (Revenue → COGS → Gross Profit → OpEx → Net Income)
- Division operations (margin percentages)
- Dependency analysis (impact of revenue changes)
- Production-ready evaluation with cycle detection

### Dependency Analysis (~100 LOC)

**File**: `examples/dependency-analysis.sc`

Advanced dependency graph features:
- Build dependency graphs from formula cells
- Detect circular references with precise cycle paths (Tarjan's SCC)
- Compute topological evaluation order (Kahn's algorithm)
- Query precedents and dependents (O(1) lookups)
- Perform impact analysis (transitive dependencies)

**Perfect for**: Meta-programming, formula auditing, testing frameworks

**Demonstrates**:
- `DependencyGraph.fromSheet(sheet)`
- `DependencyGraph.detectCycles(graph)` (with 4 cycle scenarios)
- `DependencyGraph.topologicalSort(graph)`
- `DependencyGraph.precedents/dependents(graph, ref)`
- Transitive dependency calculation
- What-if analysis (change propagation)

### Data Validation (~120 LOC)

**File**: `examples/data-validation.sc`

Quality control and error detection:
- Validate data ranges (MIN/MAX bounds checking)
- Detect missing data (COUNT vs expected)
- Normalize text (UPPER/LOWER for consistency)
- Validate text length (LEN for field constraints)
- Handle division by zero gracefully

**Perfect for**: ETL pipelines, data quality teams, validation frameworks

**Demonstrates**:
- Validation formulas: `IF(AND(value>=0, value<=100), "Valid", "INVALID")`
- Completeness checks: `COUNT(range) = expected`
- Text operations: `UPPER`, `LEFT`, `RIGHT`, `LEN`
- Error handling patterns

### Sales Pipeline Analytics (~150 LOC)

**File**: `examples/sales-pipeline.sc`

CRM and sales operations:
- Pipeline funnel with stage-by-stage conversion rates
- Tiered commission calculations (nested IF)
- Quota attainment tracking
- Average deal size metrics

**Perfect for**: Sales teams, RevOps, CRM applications

**Demonstrates**:
- Conversion funnels: `current_stage / previous_stage`
- Tiered logic: `IF(x<=t1, rate1, IF(x<=t2, rate2, rate3))`
- Safe division (automatic zero handling)
- Range aggregation across pipeline stages

### Comprehensive Evaluator Demo

**File**: `examples/evaluator-demo.sc`

Complete formula system tour:
- Low-level API (Evaluator.eval for programmatic use)
- High-level API (Sheet extensions)
- All 21 built-in functions
- Error handling patterns
- Dependency analysis
- Integration with fx macro

**Perfect for**: Learning all formula system capabilities systematically

### Running the Examples

All example scripts use `scala-cli` with explicit dependencies:

```bash
# Option 1: Run directly (downloads dependencies)
scala-cli examples/quick-start.sc

# Option 2: Publish locally first (faster, uses local build)
./mill xl-core.publishLocal && ./mill xl-evaluator.publishLocal
scala-cli examples/quick-start.sc

# Run specific examples
scala-cli examples/financial-model.sc
scala-cli examples/dependency-analysis.sc
scala-cli examples/data-validation.sc
scala-cli examples/sales-pipeline.sc
scala-cli examples/evaluator-demo.sc
```

## 4) Chart spec
```scala
import com.tjclp.xl.chart.*, com.tjclp.xl.dsl.range

val revenue = ChartSpec(
  title  = Some("Revenue by Quarter"),
  series = Vector(Series(MarkType.Column(true, false), Encoding(Field.Range(range"A2:A5"), Field.Range(range"B2:B5")), Some("2025"))),
  xAxis  = Axis(AxisType.Category, Scale.Linear, Some("Quarter")),
  yAxis  = Axis(AxisType.Value, Scale.Linear, Some("USD (mm)")),
  legend = Legend(true, "right"),
  plotAreaFill = None
)
```
