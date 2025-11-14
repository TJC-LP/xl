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
| `new XSSFWorkbook()` | `Workbook(SheetName.unsafe("Sheet1"))` | Returns Either |
| `workbook.createSheet("Data")` | `Workbook("Data").map(_._1)` | Immutable |
| `sheet.createRow(0)` | N/A (cell-oriented) | XL is cell-based, not row-based |
| `row.createCell(0)` | `sheet.put(cell"A1", value)` | Pure function |
| `cell.setCellValue("Hello")` | `cell"A1" := "Hello"` | DSL syntax |
| `cell.setCellValue(42)` | `cell"A1" := 42` | Auto-converts |
| `cell.setCellValue(true)` | `cell"A1" := true` | Type-safe |
| `cell.getCellValue()` | `sheet.get(cell"A1")` | Returns Option[Cell] |
| `cell.getNumericCellValue()` | `sheet.readTyped[Int](cell"A1")` | Type-safe, Either for errors |
| `cell.setCellStyle(style)` | `sheet.withCellStyle(cell"A1", styleId)` | Requires style registry |
| `workbook.write(out)` | `ExcelIO.instance.write[IO](wb, path)` | Effect-wrapped |

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
import com.tjclp.xl.api.*
import com.tjclp.xl.macros.*
import com.tjclp.xl.codec.syntax.*
import cats.effect.IO
import cats.effect.unsafe.implicits.global

val sheet = Sheet("Sales").get.put(
  cell"A1" -> "Product",
  cell"B1" -> "Revenue",
  cell"A2" -> "Widget",
  cell"B2" -> 1000
)

val workbook = Workbook(Vector(sheet))
ExcelIO.instance.write[IO](workbook, Path.of("sales.xlsx")).unsafeRunSync()
```

**Key Differences**:
- XL is **cell-oriented** (no explicit rows)
- XL uses **immutable updates** (no setters)
- XL requires **effect handling** (IO)
- XL uses **compile-time validated** cell references (`cell"A1"`)

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
import com.tjclp.xl.style.*

val style = CellStyle.default
  .withFont(Font("Arial", 12.0, bold = true, color = Color.fromHex("#FF0000")))
  .withFill(Fill.Solid(Color.fromRgb(200, 200, 200)))

val (registry, styleId) = sheet.styleRegistry.register(style)
val updated = sheet
  .copy(styleRegistry = registry)
  .withCellStyle(cell"A1", styleId)

// Or use Patch DSL:
import com.tjclp.xl.dsl.*
val patch = cell"A1".styled(style)
sheet.put(patch)
```

**Key Differences**:
- XL styles are **immutable** (builder pattern)
- XL uses **RGB colors** (not indexed colors)
- XL requires **style registration** (deduplication)
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
import com.tjclp.xl.codec.syntax.*

// Type-safe reading with Either
sheet.readTyped[String](cell"A1") match
  case Right(Some(s)) => println(s"String: $s")
  case Right(None) => println("Cell empty")
  case Left(error) => println(s"Type error: $error")

// Or pattern match on CellValue
sheet.get(cell"A1") match
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
import com.tjclp.xl.io.{Excel, RowData}
import fs2.Stream

Stream.range(1, 1_000_001)
  .map(i => RowData(i, Map(
    0 -> CellValue.Text(s"Row $i"),
    1 -> CellValue.Number(BigDecimal(i * 100))
  )))
  .through(Excel.forIO.writeStreamTrue(path, "Data"))
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
val sheet1 = sheet.put(cell"A1", "Hello")     // New Sheet
val sheet2 = sheet1.put(cell"A1", "Goodbye")  // New Sheet (sheet1 unchanged)
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
sheet.put(cell"D6", "value")  // Direct, no row object needed
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

**XL Philosophy**: Returns Either for errors (no exceptions)
```scala
ExcelIO.instance.read[IO](path).map {
  case Right(wb) => // Success
  case Left(error: XLError) => // Explicit error handling
}.unsafeRunSync()
```

**Migration Tip**: XL errors are values, not exceptions. Use pattern matching or flatMap.

---

## Performance Migration Guide

### Small Files (<10k rows)

**POI**: Use `XSSFWorkbook` (in-memory)
**XL**: Use in-memory API (same approach)

**No change needed** - both libraries perform similarly for small files.

---

### Medium Files (10k-100k rows)

**POI**: Use `XSSFWorkbook` or `SXSSFWorkbook` with window
**XL**: Use in-memory API with batching (putAll, putMixed)

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
ExcelIO.instance.write[IO](wb, path)  // IO[Unit]
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
Workbook("My Sheet!")  // Error! Exclamation mark invalid

// Use SheetName.apply for validation:
SheetName("My Sheet!") match
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
sheet.get(cell"A1") match
  case Some(cell) => // Exists
  case None => sheet.put(cell"A1", CellValue.Empty)  // Explicit creation
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
import com.tjclp.xl.api.*
import com.tjclp.xl.macros.*
import com.tjclp.xl.codec.syntax.*
import com.tjclp.xl.style.*
import java.time.LocalDate
import cats.effect.IO
import cats.effect.unsafe.implicits.global

// Header style
val headerStyle = CellStyle.default
  .withFont(Font("Arial", 12.0, bold = true))

// Create sheet
val sheet = Sheet("Q1 Sales").get
  .put(
    cell"A1" -> "Product",
    cell"B1" -> "Revenue",
    cell"C1" -> "Date"
  )
  .withRangeStyle(range"A1:C1", headerStyle)

// Data rows (batched for performance)
val dataRows = (1 to 1000).flatMap(i =>
  Seq(
    cell"A${i+1}" -> s"Product $i",
    cell"B${i+1}" -> i * 100,
    cell"C${i+1}" -> LocalDate.now()
  )
)
val withData = sheet.put(dataRows*)

// Write
val workbook = Workbook(Vector(withData))
ExcelIO.instance.write[IO](workbook, Path.of("sales.xlsx")).unsafeRunSync()
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

**Note**: XL streaming read has bug (P6.6), use in-memory for now

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
**A**: No. XL is ~55% feature complete. Missing: charts, drawings, pivot tables, formula evaluation. See roadmap.

### Q: Is XL production-ready?
**A**: Yes for core features (read/write, styling, streaming write). No for advanced features (charts, etc.).

### Q: How do I handle errors in XL?
**A**: XL uses Either[XLError, A]. Pattern match or use flatMap/map:
```scala
for
  workbook <- ExcelIO.instance.read[IO](path)
  sheet <- IO.fromEither(workbook.flatMap(_.sheet("Data")))
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
