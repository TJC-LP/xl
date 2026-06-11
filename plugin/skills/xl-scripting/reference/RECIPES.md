# XL Scripting Recipes

Seven complete, runnable scripts. Save as `recipe.sc`, adjust paths, run `scala-cli run recipe.sc`.

Convention: every fenced block starting with `//> using` is a complete script and is compile-verified in CI; the headers are byte-identical so a release bump is a mechanical substitution.

## 1. Bulk transform across many files

Apply a 5% price increase to column C of every workbook in a directory, in place (atomic).

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.12.0
import com.tjclp.xl.scripting.{*, given}
import java.nio.file.{Files, Paths}
import scala.jdk.CollectionConverters.*

val dir = Paths.get("invoices")
val files = Files.list(dir).iterator.asScala.filter(_.toString.endsWith(".xlsx")).toList

files.foreach { path =>
  Excel.modify(path.toString) { wb =>
    wb.sheets.foldLeft(wb) { (acc, sheet) =>
      val bumped = sheet.cells.keys
        .filter(r => r.col.index0 == 2 && r.row.index0 > 0) // column C, skip header
        .foldLeft(Patch.empty) { (p, r) =>
          sheet.readTypedOpt[BigDecimal](r) match
            case Some(price) => p ++ (r := price * BigDecimal("1.05"))
            case None => p
        }
      acc.put(sheet.put(bumped))
    }
  }
  println(s"✓ ${path.getFileName}")
}
println(s"updated ${files.size} files")
```

## 2. Typed extraction + validation report

Read rows into case classes; collect per-cell validation failures into a new Issues sheet.

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.12.0
import com.tjclp.xl.scripting.{*, given}

final case class Order(id: Int, customer: String, amount: BigDecimal)

val wb = Excel.read("orders.xlsx")
val sheet = wb("Orders").unsafe

val rows = sheet.usedRange.map(_.end.row.index0).getOrElse(0)
val (orders, issues) = (1 to rows).foldLeft((List.empty[Order], List.empty[String])) {
  case ((ok, bad), i) =>
    val r = ref"A1".down(i)
    (
      sheet.readTypedOpt[Int](r),
      sheet.readTypedOpt[String](r.right(1)),
      sheet.readTypedOpt[BigDecimal](r.right(2))
    ) match
      case (Some(id), Some(c), Some(a)) if a > 0 => (Order(id, c, a) :: ok, bad)
      case _ => (ok, s"row ${i + 1}: missing or invalid fields" :: bad)
}

val issueSheet = issues.reverse.zipWithIndex.foldLeft(Sheet("Issues").put(ref"A1", "Issue")) {
  case (s, (msg, i)) => s.put(ref"A2".down(i), msg)
}
Excel.write(wb.put(issueSheet), "orders-validated.xlsx")
println(s"${orders.size} valid orders, ${issues.size} issues")
```

## 3. Financial model: build, recalculate, verify

Three-year projection with growth formulas; fail the script if any formula errors.

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.12.0
import com.tjclp.xl.scripting.{*, given}

val header = CellStyle.default.bold.size(12.0).center
val currency = CellStyle.default.currency

val model = Sheet("Projection").put(
  (ref"A1" := "3-Year Revenue Projection") ++ ref"A1:D1".merge ++ ref"A1".styled(header) ++
    (ref"A3" := "") ++ (ref"B3" := "Y1") ++ (ref"C3" := "Y2") ++ (ref"D3" := "Y3") ++
    ref"A3:D3".styled(CellStyle.default.bold) ++
    (ref"A4" := "Revenue") ++ (ref"B4" := 1000000) ++
    (ref"C4" := fx"=B4*1.15") ++ (ref"D4" := fx"=C4*1.15") ++
    (ref"A5" := "Costs") ++ (ref"B5" := fx"=B4*0.6") ++
    (ref"C5" := fx"=C4*0.6") ++ (ref"D5" := fx"=D4*0.6") ++
    (ref"A6" := "Profit") ++ (ref"B6" := fx"=B4-B5") ++
    (ref"C6" := fx"=C4-C5") ++ (ref"D6" := fx"=D4-D5") ++
    ref"B4:D6".styled(currency)
)

Workbook(model).recalculate().toEither match
  case Right(wb) =>
    Excel.write(wb, "projection.xlsx")
    given Sheet = wb.sheets.headOption.getOrElse(sys.exit(1))
    println(excel"Y3 profit: ${ref"D6"}")
  case Left(errors) =>
    errors.foreach(e => println(s"✗ ${e.render}"))
    sys.exit(1)
```

## 4. Append many workbooks into one sheet

Stack the data rows of every input file under one header.

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.12.0
import com.tjclp.xl.scripting.{*, given}
import java.nio.file.{Files, Paths}
import scala.jdk.CollectionConverters.*

val files = Files.list(Paths.get("monthly")).iterator.asScala
  .filter(_.toString.endsWith(".xlsx")).toList.sortBy(_.toString)

val (combined, _) = files.foldLeft((Sheet("Combined").put(ref"A1", "Source"), 1)) {
  case ((acc, nextRow), path) =>
    val src = Excel.read(path.toString).sheets.headOption.getOrElse(sys.exit(1))
    val dataRefs = src.cells.keys.filter(_.row.index0 > 0).toList // skip per-file header
    val byRow = dataRefs.groupBy(_.row.index0).toList.sortBy(_._1)
    byRow.zipWithIndex.foldLeft((acc, nextRow)) { case ((s, row), ((_, refs), i)) =>
      val tagged = refs.foldLeft(
        s.put(ref"A1".down(row + i), path.getFileName.toString)
      ) { (s2, srcRef) =>
        srcRef.col.index0 match
          case c =>
            src.cells.get(srcRef).map(_.value) match
              case Some(v) => s2.put(ref"A1".down(row + i).right(c + 1), v)
              case None => s2
      }
      (tagged, row)
    } match
      case (s, row) => (s, row + byRow.size)
}

Excel.write(Workbook(combined), "combined.xlsx")
println(s"combined ${files.size} files, ${combined.cells.size} cells")
```

## 5. Streaming filter: 500k rows in, matching rows out (O(1) memory)

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.12.0
import com.tjclp.xl.scripting.{*, given}
import cats.effect.IO
import cats.effect.unsafe.implicits.global
import java.nio.file.Paths

val excel = ExcelIO.instance[IO]
val isLarge: RowData => Boolean = row =>
  row.cells.get(3).exists {
    case CellValue.Number(n) => n > 10000
    case _ => false
  }

excel
  .readStream(Paths.get("transactions.xlsx"))
  .filter(row => row.rowIndex == 1 || isLarge(row)) // keep header + matches
  .zipWithIndex
  .map { case (row, i) => row.copy(rowIndex = i.toInt + 1) } // reindex contiguously
  .through(excel.writeStream(Paths.get("large-transactions.xlsx"), "Filtered"))
  .compile
  .drain
  .unsafeRunSync()
println("✓ filtered (constant memory)")
```

## 6. Cell-level workbook diff

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.12.0
import com.tjclp.xl.scripting.{*, given}

val before = Excel.read("before.xlsx")
val after = Excel.read("after.xlsx")

val sheetNames = (before.sheets.map(_.name) ++ after.sheets.map(_.name)).distinct
val diffs = sheetNames.flatMap { name =>
  val b = before.sheets.find(_.name == name).map(_.cells).getOrElse(Map.empty)
  val a = after.sheets.find(_.name == name).map(_.cells).getOrElse(Map.empty)
  (b.keySet ++ a.keySet).toList.sortBy(r => (r.row.index0, r.col.index0)).flatMap { ref =>
    val left = b.get(ref).map(_.value)
    val right = a.get(ref).map(_.value)
    if left == right then None
    else Some(s"${name.value}!${ref.toA1}: ${left.getOrElse("∅")} → ${right.getOrElse("∅")}")
  }
}

if diffs.isEmpty then println("identical")
else
  diffs.foreach(println)
  println(s"${diffs.size} difference(s)")
```

## 7. CSV ingest → styled workbook (smart format detection)

```scala
//> using scala 3.8.3
//> using dep com.tjclp::xl:0.12.0
import com.tjclp.xl.scripting.{*, given}
import java.nio.file.{Files, Paths}
import scala.jdk.CollectionConverters.*

val lines = Files.readAllLines(Paths.get("data.csv")).asScala.toList
val headerStyle = CellStyle.default.bold.bgGray

val sheet = lines.zipWithIndex.foldLeft(Sheet("Imported")) { case (s, (line, rowIdx)) =>
  line.split(',').toList.zipWithIndex.foldLeft(s) { case (s2, (raw, colIdx)) =>
    val target = ref"A1".down(rowIdx).right(colIdx)
    if rowIdx == 0 then s2.put(target, raw.trim, headerStyle)
    else s2.put(target, raw.trim.toFormatted) // "$1,234.56" → Currency, "45.5%" → Percent, ...
  }
}

Excel.write(Workbook(sheet), "imported.xlsx")
println(s"imported ${lines.size} rows × ${lines.headOption.map(_.split(',').length).getOrElse(0)} cols")
```
