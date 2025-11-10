
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

## 2) Typed formula
```scala
import com.tjclp.xl.formula.TExpr.*
val avg: TExpr[BigDecimal] = Div(
  FoldRange(range"A2:A10", BigDecimal(0), (acc: BigDecimal, a: BigDecimal) => acc + a, decodeNumber),
  FoldRange(range"A2:A10", BigDecimal(0), (n: BigDecimal, _: BigDecimal) => n + 1, decodeNumber)
)
```

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
