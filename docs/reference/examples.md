
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

**Note**: Formula *evaluation* (WI-08) is not yet implemented. This example shows parsing and AST manipulation only.

## 3) Chart spec
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
