# Migration Guide: Apache POI → XL

## Overview

This guide helps Java developers migrate from Apache POI to XL. The two libraries have fundamentally different philosophies (imperative vs functional), but XL provides significant advantages in performance, type safety, and correctness.

## Why Migrate?

| Advantage | POI | XL | Improvement |
|-----------|-----|-----|-------------|
| **Performance** | 5s @ 800MB | 1.1s @ 10MB | **4.5x faster, 80x less memory** |
| **Type Safety** | Runtime (Object) | Compile-time | Catches errors at compile time |
| **Purity** | Mutable, imperative | Immutable, pure | Easier testing, no side effects |
| **Determinism** | Non-deterministic XML | Byte-identical | Stable diffs, reproducible builds |
| **Streaming** | SXSSF (complex) | fs2 (natural) | Cleaner API, better composability |

## API Comparison Table

| POI API | XL Equivalent | Notes |
|---------|---------------|-------|
| `new XSSFWorkbook()` | `Workbook.empty` | Returns `Workbook` (total) |
| `workbook.createSheet("Data")` | `Workbook(Sheet("Data"))` | Immutable; literal sheet names are compile-time validated |
| `sheet.createRow(0)` | N/A – XL is cell-oriented |
| `row.createCell(0)` | `sheet.put(ref"A1", value)` | Pure and immutable |
| `cell.setCellValue("Hello")` | `ref"A1" := "Hello"` | DSL syntax |
| `cell.setCellValue(42)` | `ref"A1" := 42` | Auto-converts |
| `cell.setCellValue(true)` | `ref"A1" := true` | Type-safe |
| `cell.getCellValue()` | `sheet(ref"A1")` | Returns `Cell` (empty cells yield `CellValue.Empty`) |
| `cell.getNumericCellValue()` | `sheet.readTyped[Int](ref"A1")` | Returns `Either[CodecError, Option[Int]]` |
| `cell.setCellStyle(style)` | `sheet.withCellStyle(ref"A1", style)` | Style auto-registered & deduplicated |
| `workbook.write(out)` | `Excel.write(wb, "out.xlsx")` / `ExcelIO.instance[IO].write(wb, path)` | Sync facade or effect-wrapped |

## Common Patterns

### Creating a Simple Spreadsheet

**POI (Java)**:
```java
XSSFWorkbook workbook = new XSSFWorkbook();
XSSFSheet sheet = workbook.createSheet("Sales");

XSSFRow headerRow = sheet.createRow(0);
headerRow.createCell(0).setCellValue("Product");
headerRow.createCell(1).setCellValue("Revenue");

XSSFRow dataRow = sheet.createRow(1);
dataRow.createCell(0).setCellValue("Widget");
dataRow.createCell(1).setCellValue(1000.0);

FileOutputStream out = new FileOutputStream("sales.xlsx");
workbook.write(out);
out.close();
```

**XL (Scala)**:
```scala
import com.tjclp.xl.scripting.{*, given}

val sheet = Sheet("Sales").put(
  ref"A1" -> "Product",
  ref"B1" -> "Revenue",
  ref"A2" -> "Widget",
  ref"B2" -> 1000
)

Excel.write(Workbook(sheet), "sales.xlsx")
```

- **Key Differences**:
  - XL is **cell-oriented** (no explicit rows)
  - XL uses **immutable updates** (no setters)
  - XL isolates **IO at the edge** (sync `Excel` facade for scripts, `ExcelIO.instance[IO]` for Cats Effect services)
  - XL uses **compile-time validated** references (`ref"A1"`)

---

### Applying Cell Styles

**POI (Java)**:
```java
CellStyle style = workbook.createCellStyle();
Font font = workbook.createFont();
font.setBold(true);
font.setColor(IndexedColors.RED.getIndex());
style.setFont(font);
style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

Cell cell = sheet.getRow(0).getCell(0);
cell.setCellStyle(style);
```

**XL (Scala)**:
```scala
import com.tjclp.xl.scripting.{*, given}

// Style DSL: compile-time validated hex colors, fluent builders
val style = CellStyle.default.bold.size(12.0).fontFamily("Arial")
  .hex("#FF0000")          // font color
  .bgRgb(200, 200, 200)    // fill

val updated = sheet.withCellStyle(ref"A1", style) // auto-registered & deduplicated

// Or use the Patch DSL:
val patch = ref"A1".styled(style)
sheet.put(patch)
```

**Key Differences**:
- XL styles are **immutable** (builder pattern)
- XL uses **RGB colors** (not indexed colors)
- XL **auto-registers and deduplicates** styles (no manual registry bookkeeping)
- XL supports **StylePatch Monoid** for composition

---

### Reading Cell Values

**POI (Java)**:
```java
Cell cell = sheet.getRow(0).getCell(0);

switch (cell.getCellType()) {
    case STRING:
        String s = cell.getStringCellValue();
        break;
    case NUMERIC:
        double d = cell.getNumericCellValue();
        break;
    case BOOLEAN:
        boolean b = cell.getBooleanCellValue();
        break;
}
```

**XL (Scala)**:
```scala
import com.tjclp.xl.scripting.{*, given}

// Type-safe reading with Either
sheet.readTyped[String](ref"A1") match
  case Right(Some(s)) => println(s"String: $s")
  case Right(None) => println("Cell empty")
  case Left(error) => println(s"Type error: $error")

// Or pattern match on CellValue
sheet.cells.get(ref"A1") match
  case Some(cell) =>
    cell.value match
      case CellValue.Text(s) => println(s"String: $s")
      case CellValue.Number(n) => println(s"Number: $n")
      case CellValue.Bool(b) => println(s"Boolean: $b")
      case _ => println("Other type")
  case None => println("Cell empty")
```

**Key Differences**:
- XL uses **Either** for error handling (no exceptions)
- XL has **Option[Cell]** for missing cells
- XL provides **type-safe reading** with `readTyped[A]`
- XL supports **pattern matching** on CellValue

---

### Streaming Large Files

**POI (Java)**:
```java
// SXSSFWorkbook for streaming write
SXSSFWorkbook workbook = new SXSSFWorkbook(100); // Keep 100 rows in memory

for (int i = 0; i < 1_000_000; i++) {
    Row row = workbook.getSheet("Data").createRow(i);
    row.createCell(0).setCellValue("Row " + i);
    row.createCell(1).setCellValue(i * 100.0);
}

FileOutputStream out = new FileOutputStream("large.xlsx");
workbook.write(out);
out.close();
workbook.dispose(); // Delete temp files
```

**XL (Scala)**:
```scala
import com.tjclp.xl.{*, given}
import com.tjclp.xl.io.{Excel, RowData}
import cats.effect.unsafe.implicits.global
import fs2.Stream

Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(
    0 -> CellValue.Text(s"Row $i"),
    1 -> CellValue.Number(BigDecimal(i * 100))
  )))
  .through(Excel.forIO.writeStream(path, "Data"))
  .compile.drain
  .unsafeRunSync()

// Memory: ~10MB constant (POI: ~800MB)
// No temp files!
```

**Key Differences**:
- XL uses **fs2 streams** (functional, composable)
- XL has **true O(1) memory** (POI keeps window in memory)
- XL **no temp files** (POI creates disk-backed SXSSFWorkbook)
- XL is **4.5x faster**

---

## Conceptual Differences

### Imperative vs Functional

**POI Philosophy**: Mutable objects, setters, side effects
```java
Cell cell = row.createCell(0);
cell.setCellValue("Hello");  // Mutates cell
cell.setCellValue("Goodbye"); // Overwrites
```

**XL Philosophy**: Immutable values, pure functions, explicit updates
```scala
val sheet1 = sheet.put(ref"A1", "Hello")     // New Sheet
val sheet2 = sheet1.put(ref"A1", "Goodbye")  // New Sheet (sheet1 unchanged)
```

**Migration Tip**: Think of XL operations as creating new versions, not modifying in place.

---

### Row-Based vs Cell-Based

**POI Philosophy**: Row-centric (create rows, then cells)
```java
Row row = sheet.createRow(5);
Cell cell = row.createCell(3);  // D6
```

**XL Philosophy**: Cell-centric (direct cell addressing)
```scala
sheet.put(ref"D6", "value")  // Direct, no row object needed
```

**Migration Tip**: XL doesn't have explicit Row objects. Use `ARef` (cell references) directly.

---

### Exceptions vs Either

**POI Philosophy**: Throws exceptions for errors
```java
try {
    Workbook wb = WorkbookFactory.create(file);
    // ...
} catch (InvalidFormatException | IOException e) {
    // Handle
}
```

**XL Philosophy**: Returns Either for errors (no exceptions in the pure core)
```scala
ExcelIO.instance[IO].readR(path).map {
  case Right(wb) => // Success
  case Left(error: XLError) => // Explicit error handling
}.unsafeRunSync()
```

**Migration Tip**: XL errors are values, not exceptions. Use pattern matching or flatMap. (`read` raises the error in `F` for monadic pipelines; `readR` surfaces it as `XLResult`. The sync `Excel.read` facade throws — only at the script IO edge.)

---

## Performance Migration Guide

### Small Files (<10k rows)

**POI**: Use `XSSFWorkbook` (in-memory)
**XL**: Use in-memory API (same approach)

**No change needed** - both libraries perform similarly for small files.

---

### Medium Files (10k-100k rows)

**POI**: Use `XSSFWorkbook` or `SXSSFWorkbook` with window
**XL**: Use in-memory API with batching (varargs `Sheet.put` / Patch folds)

**XL Advantage**: Simpler API, no temp files, faster

---

### Large Files (100k+ rows)

**POI**: Must use `SXSSFWorkbook` (complex setup, temp files)
**XL**: Use streaming write (simple, no temp files)

**XL Advantage**:
- 4.5x faster
- 80x less memory
- No temp file cleanup
- Cleaner functional API

---

## Common Gotchas

### Gotcha 1: IO[Unit] instead of void

**POI**:
```java
workbook.write(out);  // void
```

**XL**:
```scala
ExcelIO.instance[IO].write(wb, path)  // IO[Unit]
  .unsafeRunSync()  // Must call to execute!
```

**Why**: XL uses Cats Effect for IO management (referential transparency)

---

### Gotcha 2: Sheet Names Must Be Valid

**POI**:
```java
workbook.createSheet("My Sheet!");  // Works (POI allows most chars)
```

**XL**:
```scala
Sheet("Sheet[1]")  // Compile error! Brackets are invalid in sheet names

// Runtime strings validate via SheetName.apply:
SheetName(userInput) match
  case Right(name) => // Valid
  case Left(err) => // Invalid char
```

**Why**: XL validates sheet names against ECMA-376 spec (prevents Excel errors)

---

### Gotcha 3: No Auto-Create for Missing Cells

**POI**:
```java
Row row = sheet.getRow(0);
if (row == null) {
    row = sheet.createRow(0);  // Auto-create
}
```

**XL**:
```scala
sheet.cells.get(ref"A1") match
  case Some(cell) => // Exists
  case None => sheet.put(ref"A1", CellValue.Empty)  // Explicit creation
```

**Why**: XL is explicit (no magic auto-creation)

---

## Example Migration: Sales Report

### POI Version (Java)
```java
XSSFWorkbook workbook = new XSSFWorkbook();
XSSFSheet sheet = workbook.createSheet("Q1 Sales");

// Header
XSSFRow header = sheet.createRow(0);
header.createCell(0).setCellValue("Product");
header.createCell(1).setCellValue("Revenue");
header.createCell(2).setCellValue("Date");

// Header style
CellStyle headerStyle = workbook.createCellStyle();
Font font = workbook.createFont();
font.setBold(true);
headerStyle.setFont(font);
header.forEach(cell -> cell.setCellStyle(headerStyle));

// Data rows
for (int i = 0; i < 1000; i++) {
    XSSFRow row = sheet.createRow(i + 1);
    row.createCell(0).setCellValue("Product " + i);
    row.createCell(1).setCellValue(i * 100.0);
    row.createCell(2).setCellValue(LocalDate.now());
}

// Write
FileOutputStream out = new FileOutputStream("sales.xlsx");
workbook.write(out);
out.close();
```

### XL Version (Scala)
```scala
import com.tjclp.xl.scripting.{*, given}
import java.time.LocalDate

// Header style
val headerStyle = CellStyle.default.bold.size(12.0).fontFamily("Arial")

// Header + style
val header = Sheet("Q1 Sales")
  .put(
    ref"A1" -> "Product",
    ref"B1" -> "Revenue",
    ref"C1" -> "Date"
  )
  .withRangeStyle(ref"A1:C1", headerStyle)

// Data rows: fold into a Patch, apply once (codecs auto-infer the date format)
val dataRows = (1 to 1000).foldLeft(Patch.empty) { (p, i) =>
  val r = ref"A2".down(i - 1)
  p ++ (r := s"Product $i") ++ (r.right(1) := i * 100) ++
    (r.right(2) := LocalDate.now().atStartOfDay)
}
val withData = header.put(dataRows)

// Write
Excel.write(Workbook(withData), "sales.xlsx")
```

**Differences**:
- XL is more concise (fewer lines)
- XL batches operations (faster)
- XL is type-safe (compile-time checks)
- XL uses functional composition

---

## Performance Comparison

### Write Performance (100k rows)

**POI SXSSF**:
```
Time: ~5s
Memory: ~800MB
Code: 50 lines
Temp files: Yes
```

**XL Streaming**:
```
Time: ~1.1s (4.5x faster)
Memory: ~10MB (80x less)
Code: 10 lines
Temp files: No
```

**Winner**: XL by large margin

---

### Read Performance (100k rows)

**POI**:
```
Time: ~8s
Memory: ~1GB
```

**XL In-Memory**:
```
Time: ~1.8s (4.4x faster)
Memory: ~100MB (10x less)
```

**Winner**: XL

---

## When to Use Each Library

### Use POI When:
- You need 100% feature parity (charts, pivot tables, VBA macros)
- Your team only knows Java (learning curve issue)
- You have existing POI code (migration cost high)

### Use XL When:
- You value performance (4.5x faster, 80x less memory)
- You value correctness (pure functional, type-safe)
- You're writing new Scala code (no migration cost)
- You need deterministic output (git-friendly)
- You're processing large files (100k+ rows)

---

## Common Questions

### Q: Can I mix POI and XL?
**A**: Not easily. They use different in-memory representations. You'd need to write to file and read back.

### Q: Does XL support all POI features?
**A**: No. Still missing: charts, drawings, pivot tables, VBA. Formula **evaluation** is a point in XL's favor: 104 functions with whole-workbook recalculation (POI's evaluator exists but is partial and mutable). See [roadmap](../plan/roadmap.md) and [LIMITATIONS](../LIMITATIONS.md).

### Q: Is XL production-ready?
**A**: Yes for core features (read/write, styling, formulas, structural editing, streaming). No for visual features (charts, drawings).

### Q: How do I handle errors in XL?
**A**: XL uses Either[XLError, A]. Pattern match or use flatMap/map:
```scala
val excel = ExcelIO.instance[IO]
for
  workbook <- excel.read(path)              // errors raised in IO
  sheet <- IO.fromEither(workbook("Data"))  // XLResult[Sheet] → IO
yield sheet
```

### Q: Why is XL so much faster?
**A**:
1. Zero-overhead opaque types (no wrapper allocations)
2. fs2 streaming (true constant memory)
3. Optimized style deduplication (LinkedHashSet, memoized keys)
4. Single-pass algorithms (usedRange, SST building)

---

## Related Documentation

- [performance-guide.md](performance-guide.md) - When to use which I/O mode
- [examples.md](examples.md) - More XL code samples
- [QUICK-START.md](../QUICK-START.md) - Getting started
- [README.md](../../README.md) - Full documentation
