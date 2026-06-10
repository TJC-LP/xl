
# Examples — End‑to‑End Snippets

> **Start here**: [`examples/scripting_tour.sc`](../../examples/scripting_tour.sc) is the canonical, runnable tour of the scripting prelude — one import, compile-time refs, patch folds, recalculation, typed reads, and the `.unsafe` boundary in ~80 lines. The full catalog of runnable scripts lives in [`examples/README.md`](../../examples/README.md), and the prose companion is the [Scripting Guide](scripting.md).

Most snippets below use the **scripting prelude** (`import com.tjclp.xl.scripting.{*, given}`), the one-import path for scripts. Where an example deliberately demonstrates the **pure library import** (`com.tjclp.xl.{*, given}` — no `.unsafe`, no sync IO), it is labeled as such. Never combine the two imports in one file.

## 1) Export a simple table

```scala
import com.tjclp.xl.scripting.{*, given}

final case class Person(name: String, age: Int, email: String) derives CanEqual
val people = Vector(Person("Ada", 34, "ada@ex.com"), Person("Linus", 55, "linus@ex.com"))

// Headers + data rows: fold everything into one Patch, apply once
val header = (ref"A1" := "Name") ++ (ref"B1" := "Age") ++ (ref"C1" := "Email")
val rows = people.zipWithIndex.foldLeft(Patch.empty) { case (acc, (p, i)) =>
  val r = ref"A2".down(i) // total navigation — no Either in the loop
  acc ++ (r := p.name) ++ (r.right(1) := p.age) ++ (r.right(2) := p.email)
}

val sheet = Sheet("People").put(header ++ rows)
Excel.write(Workbook(sheet), "people.xlsx")
```

For production code with Cats Effect (no sync facade), use the pure import + `ExcelIO`:

```scala
// Pure library import (deliberately NOT the scripting prelude)
import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.ExcelIO
import cats.effect.IO
import java.nio.file.Path

def write(wb: Workbook, path: Path): IO[Unit] =
  ExcelIO.instance[IO].write(wb, path)
```

## 2) Formula Parsing & Typed AST

This example uses the **pure library import** — the formula AST works identically under the prelude.

```scala
import com.tjclp.xl.{*, given}
import com.tjclp.xl.formula.{FormulaParser, FormulaPrinter, TExpr, ParseError}

// ==================== Parsing Formula Strings ====================

// Parse basic formulas
val sum = FormulaParser.parse("=SUM(A2:A10)")
// Right(TExpr.Call(FunctionSpecs.sum, TExpr.RangeLocation.Local(CellRange("A2:A10"))))

val conditional = FormulaParser.parse("=IF(A1>100, \"High\", \"Low\")")
// Right(TExpr.Call(FunctionSpecs.ifFn, (TExpr.Gt(...), TExpr.Lit("High"), TExpr.Lit("Low"))))

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

val nestedIf: TExpr[String] = TExpr.cond(
  TExpr.Gt(
    TExpr.ref(ref"A1", TExpr.decodeNumeric),
    TExpr.Lit(BigDecimal(100))
  ),
  TExpr.Lit("High"),
  TExpr.cond(
    TExpr.Gt(
      TExpr.ref(ref"A1", TExpr.decodeNumeric),
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

**Note**: Evaluation is fully operational — 104 functions, whole-workbook `recalculate()`, dependency graphs with cycle detection. See the runnable scripts below.

## 3) Example Scripts Catalog

The `/examples/` directory contains ready-to-run scripts. Highlights (full catalog: [`examples/README.md`](../../examples/README.md)):

### Scripting Tour (start here)

**File**: `examples/scripting_tour.sc`

Canonical tour of the scripting prelude — ONE import gives a script everything:
- Compile-time literals (`ref"A1"`, `fx"..."`) and total navigation
- Bulk generation by folding data into a `Patch`
- Range fill (`ref"E2:E4" := 0`), smart detection (`"$1,234.56".toFormatted`)
- `recalculate()` with per-cell errors, typed reads, the `excel""` display interpolator
- Sync `Excel` IO and the one `.unsafe` boundary

### Quick Start (5 minutes)

**File**: `examples/quick_start.sc`

Formula system in 5 minutes (pure import):
- Parse Excel formulas to typed AST
- Evaluate formulas against sheets
- Handle errors gracefully (Either types, no exceptions)
- Detect circular references automatically
- Evaluate all formulas safely with dependency checking

**Key APIs shown**:
- `FormulaParser.parse("=SUM(A1:A10)")`
- `sheet.evaluateCell(ref)`
- `sheet.evaluateWithDependencyCheck()`
- `DependencyGraph.detectCycles(graph)`

### Financial Modeling

**File**: `examples/financial_model.sc`

Complete 3-year income statement with:
- Revenue projections with growth rates
- COGS, operating expenses, EBITDA, net income
- Financial ratios (gross margin, operating margin, net margin)
- Dependency chain analysis (evaluation order)

### Dependency Analysis

**File**: `examples/dependency_analysis.sc`

Advanced dependency graph features:
- Build dependency graphs from formula cells
- Detect circular references with precise cycle paths (Tarjan's SCC)
- Compute topological evaluation order (Kahn's algorithm)
- Query precedents and dependents
- Perform impact analysis (transitive dependencies)

### Data Validation

**File**: `examples/data_validation.sc`

Quality control and error detection:
- Validate data ranges (MIN/MAX bounds checking)
- Detect missing data (COUNT vs expected)
- Normalize text (UPPER/LOWER), validate lengths (LEN)
- Handle division by zero gracefully

### Sales Pipeline Analytics

**File**: `examples/sales_pipeline.sc`

CRM and sales operations:
- Pipeline funnel with stage-by-stage conversion rates
- Tiered commission calculations (nested IF)
- Quota attainment tracking

### Text Functions

**File**: `examples/text_functions_demo.sc` — TRIM, MID, FIND, SUBSTITUTE, VALUE, TEXT.

### Running the Examples

Examples resolve the library via the shared `examples/project.scala`, so publish locally first:

```bash
./mill __.publishLocal
scala-cli run examples/scripting_tour.sc
scala-cli run examples/quick_start.sc
scala-cli run examples/financial_model.sc
```

CI compiles every script via `scripts/test-examples.sh` (and runs a curated subset), so the catalog cannot silently rot.

## 4) Chart spec (Future — #222)
```scala
// Note: Chart support is a 0.12.0 candidate (GH #222, on the DrawingML layer #221).
// This is a design preview only.
import com.tjclp.xl.{*, given}
// import com.tjclp.xl.chart.*  // Future module

// val revenue = ChartSpec(
//   title  = Some("Revenue by Quarter"),
//   series = Vector(Series(MarkType.Column(true, false), Encoding(Field.Range(ref"A2:A5"), Field.Range(ref"B2:B5")), Some("2025"))),
//   xAxis  = Axis(AxisType.Category, Scale.Linear, Some("Quarter")),
//   yAxis  = Axis(AxisType.Value, Scale.Linear, Some("USD (mm)")),
//   legend = Legend(true, "right"),
//   plotAreaFill = None
// )
```
