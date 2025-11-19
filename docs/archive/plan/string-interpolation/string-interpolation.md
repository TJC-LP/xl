# String Interpolation for XL Macros

**Status**: Planned (Not Yet Implemented)
**Priority**: P7 (Post-1.0 Enhancement)
**Dependencies**: P1 (Addressing & Literals), P2 (Core + Patches), P3 (Styles)
**Strategy**: Hybrid Macro (Strategy A) - Compile-time optimization with runtime fallback

## Executive Summary

This document specifies string interpolation support for all four XL macro types: `ref`, `money`, `date`, and `fx`. The design uses a **hybrid compile-time/runtime validation pattern** that:

- **Preserves compile-time validation** for string literals
- **Enables runtime flexibility** for dynamic values
- **Maintains purity and totality** through explicit error handling
- **Provides zero-overhead** for compile-time paths
- **Offers unified API** (no separate `refDyn` macro needed)

### Design Principles

1. **Single API Surface**: `ref"$sheet!$cell"` works for both literals and dynamic values
2. **Purity Maintained**: Runtime path returns `Either[XLError, T]` (no exceptions)
3. **Type Stability**: Return union types for runtime interpolations
4. **Zero-Overhead Literals**: Compile-time detection emits constants directly
5. **Excel Compatibility**: `$$` escape for absolute references (`$A$1`)

### Key Trade-offs

- ✅ **Pro**: Ergonomic single API, maximum flexibility
- ✅ **Pro**: Compile-time optimization when possible
- ⚠️ **Con**: Type instability (returns `ARef` vs `Either[XLError, ARef]` depending on runtime detection)
- ⚠️ **Con**: More complex implementation than separate `refDyn` macro
- ⚠️ **Con**: `$$` escape may confuse Excel users initially (requires documentation)

## Motivation & Use Cases

### Current Limitations

XL macros currently only support compile-time string literals:

```scala
val ref = ref"A1"                    // ✅ Works
val range = ref"Sales!A1:B10"        // ✅ Works

// Dynamic values: Compile errors
val sheetName = "Sales"
val cell = ref"$sheetName!A1"        // ❌ Compile errors: "Use string interpolation (not yet supported)"
```

Users must fall back to runtime parsing:

```scala
val result = RefType.parse(s"$sheetName!A1")  // Verbose, loses macro ergonomics
```

### Proposed Capabilities

With string interpolation support:

```scala
// Dynamic sheets names
val sheet = "Q1 Sales"
val ref = ref"$sheet!A1"                       // ✅ Returns Either[XLError, RefType]

// Dynamic cell construction
val col = "B"
val row = 42
val cell = ref"$col$row"                       // ✅ Returns Either[XLError, ARef]

// Dynamic ranges
val start = cell"A1"
val end = cell"B10"
val range = ref"$start:$end"                   // ✅ Returns Either[XLError, CellRange]

// Dynamic formulas
val cellRef = "A1:A10"
val formula = fx"=SUM($cellRef)"               // ✅ Returns Either[XLError, CellValue.Formula]

// Dynamic formatted literals
val amount = 1234.56
val formatted = money"$$$amount"               // ✅ Returns Either[XLError, Formatted]

// Mixed literals and dynamic (compile-time optimization!)
val ref = ref"Sales!$cell"                     // Literal "Sales!" validated at compile time
```

### Real-World Scenarios

**1. User-Driven Sheet Navigation**
```scala
def getCellValue(sheetName: String, cellRef: String): Either[XLError, CellValue] =
  for
    ref <- ref"$sheetName!$cellRef"  // Runtime validation
    sheet <- workbook(sheetName)
    cell <- sheet(ref)
  yield cell.value
```

**2. Programmatic Formula Generation**
```scala
def createSumFormula(rangeStart: ARef, rangeEnd: ARef): Either[XLError, CellValue] =
  val range = s"${rangeStart.toA1}:${rangeEnd.toA1}"
  fx"=SUM($range)"  // Dynamic formula construction
```

**3. Financial Reporting with Dynamic Formats**
```scala
def formatRevenue(amount: BigDecimal, currency: String): Either[XLError, Formatted] =
  currency match
    case "USD" => money"$$$amount"               // $1,234.56
    case "EUR" => money"€$amount"                // €1,234.56 (future enhancement)
    case _     => Left(XLError.UnsupportedCurrency(currency))
```

**4. CSV Import with Header Mapping**
```scala
def importRow(row: Int, colMapping: Map[String, String]): Either[XLError, Sheet] =
  for
    nameCol <- colMapping.get("Name").toRight(XLError.MissingColumn("Name"))
    emailCol <- colMapping.get("Email").toRight(XLError.MissingColumn("Email"))
    nameRef <- ref"$nameCol$row"     // Dynamic cell reference
    emailRef <- ref"$emailCol$row"
    // ... use refs to read/write cells
  yield updatedSheet
```

## Technical Architecture

### Strategy A: Hybrid Compile-Time/Runtime Macro

The hybrid pattern extends XL's existing `.hex()` method pattern (in `style/dsl.scala:281-307`) to macro interpolation.

#### Core Detection Logic

```scala
def refImpl(sc: Expr[StringContext], args: Expr[Seq[Any]])(using Quotes): Expr[...] =
  import quotes.reflect.*

  val parts = sc.valueOrAbort.parts

  if parts.length == 1 then
    // No interpolation - use existing compile-time logic
    refImpl0(sc)  // Returns ARef | CellRange | RefType (transparent inline)
  else
    // Interpolation present - check if all args are compile-time constants
    args match
      case Varargs(argExprs) =>
        val literalValues = argExprs.map(_.value)  // Option[Any]

        if literalValues.forall(_.isDefined) then
          // All compile-time constants - validate at compile time
          compileTimePath(parts, literalValues.flatten)
        else
          // Runtime variables present - emit runtime validation code
          runtimePath(sc, args)
```

#### Compile-Time Path (All Literals)

When all interpolated values are compile-time constants:

```scala
def compileTimePath(parts: Seq[String], literals: Seq[Any])(using Quotes): Expr[...] =
  // Reconstruct full string at compile time
  val fullString = interleave(parts, literals.map(_.toString))

  // Parse using existing zero-allocation parser
  parseRef(fullString) match
    case Right(ref: ARef) =>
      // Emit constant directly (zero runtime overhead)
      '{ ARef.from0(${ Expr(ref.col.index0) }, ${ Expr(ref.row.index0) }) }

    case Right(range: CellRange) =>
      // Emit CellRange constant
      '{
        CellRange(
          ARef.from0(${ Expr(range.start.col.index0) }, ${ Expr(range.start.row.index0) }),
          ARef.from0(${ Expr(range.end.col.index0) }, ${ Expr(range.end.row.index0) })
        )
      }

    case Right(qualified: RefType.QualifiedCell) =>
      // Emit RefType constant
      '{ RefType.QualifiedCell(SheetName(${ Expr(qualified.sheet.value) }), ...) }

    case Left(err) =>
      // Compile errors with helpful message
      report.errorAndAbort(
        s"Invalid cell reference in interpolation: ${fullString}\n" +
        s"Error: $err\n" +
        s"Hint: Check that all parts form a valid Excel reference"
      )
```

**Benefits**:
- ✅ Zero runtime overhead
- ✅ Compile-time validation catches errors early
- ✅ Emits same code as non-interpolated literals
- ✅ Maintains totality (abort compilation on error)

#### Runtime Path (Dynamic Values)

When any interpolated value is a runtime variable:

```scala
def runtimePath(sc: Expr[StringContext], args: Expr[Seq[Any]])(using Quotes): Expr[Either[XLError, RefType]] =
  '{
    // Build string at runtime
    val parts = $sc.parts
    val argValues = $args.map(_.toString)
    val fullString = buildInterpolatedString(parts, argValues)

    // Parse with existing runtime parser (already pure and total)
    RefType.parse(fullString)
      .left.map(err => XLError.InvalidReference(fullString, err))
  }

// Helper for interleaving parts and args
private def buildInterpolatedString(parts: Seq[String], args: Seq[String]): String =
  val sb = new StringBuilder
  var i = 0
  while i < args.length do
    sb.append(parts(i))
    sb.append(args(i))
    i += 1
  sb.append(parts(i))
  sb.toString
```

**Benefits**:
- ✅ Maintains purity (returns `Either[XLError, T]`)
- ✅ Explicit error handling (no exceptions)
- ✅ Reuses existing runtime parsers (`RefType.parse`, `ARef.parse`)
- ✅ User gets clear error messages

**Trade-off**: Type instability (returns `Either` instead of `ARef`)

### Type Stability Challenge & Solution

The transparent inline feature allows different return types:

```scala
val a: ARef = ref"A1"                    // Compile-time: Returns ARef directly
val b: Either[XLError, ARef] = ref"$x1"  // Runtime: Returns Either
```

This creates **type instability** - same API, different return types based on runtime detection.

#### Solution 1: User Type Ascription (Recommended)

Users explicitly handle both cases:

```scala
// Compile-time path (detected automatically)
val a: ARef = ref"A1"  // No Either wrapper

// Runtime path (detected automatically)
val b: Either[XLError, ARef] = ref"$col$row"  // Must handle Either

// Or use for-comprehension
for
  ref <- ref"$sheet!$cell"  // Works with Either
  sheet <- workbook(ref.sheet)
  cell <- sheet(ref.cell)
yield cell.value
```

**Rationale**: Users opting into runtime interpolation accept the `Either` return type as the cost of flexibility.

#### Solution 2: Always Return Union Type

Alternative: Always return `ARef | Either[XLError, ARef]` for interpolated refs.

```scala
transparent inline def ref(inline args: Any*): ARef | Either[XLError, ARef] = ...
```

**Pros**: Type-safe without surprises
**Cons**: Forces pattern matching even for compile-time literals

**Verdict**: Solution 1 preferred (user type ascription) for ergonomics.

### Handling the `$` Conflict

Excel uses `$A$1` for absolute references (fixed column/row in formulas). Scala uses `$` for interpolation.

#### The Problem

```scala
// User intent: Excel absolute reference $A$1
ref"$A$1"  // Scala parser sees: StringContext("", "", "").ref(A, 1)
           // Expects variables named A and 1!

// User intent: Interpolate variable into absolute ref
val col = "A"
ref"$$col$1"  // Ambiguous: $$ means literal $, or escape for $col?
```

#### Solution: `$$` Escape Per Scala Standard

Follow Scala's string interpolation escaping rules:

- `$$` in interpolated string → Single literal `$` in output
- `$variable` → Interpolate variable

**Examples**:

```scala
// Absolute reference (no interpolation)
ref"$$A$$1"         // Parsed as: "$A$1" (Excel absolute reference)
                    // StringContext("", "A", "1").ref() with no args
                    // Result: ARef(0, 0) with absolute flags

// Interpolate into absolute reference
val col = "B"
ref"$$$col$$1"      // Parsed as: "$" + col + "$1"
                    // Runtime: "$B$1" (absolute column B, absolute row 1)

// Mixed absolute and relative
ref"$$A$row"        // Absolute column A, dynamic row
                    // Runtime: "$A42" (if row = 42)
```

#### Implementation: Absolute Reference Support

Extend `ARef` to track absolute flags (future enhancement):

```scala
// Current (P1-P6)
opaque type ARef = Long  // Packed: (row << 32) | col

// Future (with absolute ref support)
opaque type ARef = Long  // Packed: (flags << 60) | (row << 32) | col
// Flags: bit 60 = col absolute, bit 61 = row absolute

object ARef:
  def apply(col: Column, row: Row, colAbsolute: Boolean = false, rowAbsolute: Boolean = false): ARef = ...

  extension (ref: ARef)
    def isColAbsolute: Boolean = (ref >> 60) & 1 == 1
    def isRowAbsolute: Boolean = (ref >> 61) & 1 == 1

    def toA1: String =
      val colPart = if isColAbsolute then s"$$${col.toA1}" else col.toA1
      val rowPart = if isRowAbsolute then s"$$${row.index1}" else row.index1.toString
      s"$colPart$rowPart"
```

**Parser modifications**:

```scala
// In RefLiteral.scala parser
def parseARef(s: String): Either[String, ARef] =
  var i = 0
  val colAbsolute = s.charAt(0) == '$'
  if colAbsolute then i += 1

  val (col, newI) = parseColumn(s, i)  // Existing logic
  i = newI

  val rowAbsolute = i < s.length && s.charAt(i) == '$'
  if rowAbsolute then i += 1

  val (row, _) = parseRow(s, i)

  Right(ARef(col, row, colAbsolute, rowAbsolute))
```

#### Documentation & Migration

**Documentation** (in `docs/reference/examples.md`):

```scala
// Absolute References (Excel $A$1 syntax)
// Use $$ to escape the dollar sign in string interpolation

// Fixed column A, fixed row 1
val abs = ref"$$A$$1"  // Excel: $A$1

// Fixed column, dynamic row
val row = 42
val mixed = ref"$$A$row"  // Excel: $A42

// Fixed row, dynamic column
val col = "B"
val mixed2 = ref"$col$$1"  // Excel: B$1

// Both dynamic
val dynamic = ref"$col$row"  // Excel: B42 (relative)
```

**Migration**: No breaking changes (absolute refs are new feature).

## Implementation by Macro Type

### 1. `ref` Macro (Cell References)

**Supported patterns**:

```scala
// Simple cell references
ref"A1"              // ARef (compile-time)
ref"$col$row"        // Either[XLError, ARef] (runtime)

// Ranges
ref"A1:B10"          // CellRange (compile-time)
ref"$start:$end"     // Either[XLError, CellRange] (runtime)

// Sheet-qualified
ref"Sales!A1"        // RefType.QualifiedCell (compile-time)
ref"$sheet!$cell"    // Either[XLError, RefType] (runtime)

// Absolute references
ref"$$A$$1"          // ARef with absolute flags (compile-time)
ref"$$$col$$1"       // Either[XLError, ARef] with absolute col (runtime)
```

**Implementation location**: `xl-core/src/com/tjclp/xl/macros/RefLiteral.scala`

**Changes required**:

1. **Remove interpolation block**:
   ```scala
   // Current: Explicitly blocks interpolation
   inline def ref(inline args: Any*): ARef | CellRange | RefType =
     ${ errorNoInterpolation('sc, 'args, "ref") }

   // New: Implement hybrid logic
   transparent inline def ref(inline args: Any*): ARef | CellRange | RefType | Either[XLError, RefType] =
     ${ refImpl('sc, 'args) }
   ```

2. **Add hybrid detection** (see Technical Architecture above)

3. **Extend parser for absolute refs** (see `$` Conflict section)

4. **Add runtime parsing helper**:
   ```scala
   private def buildRuntimeRef(sc: Expr[StringContext], args: Expr[Seq[Any]])(using Quotes): Expr[Either[XLError, RefType]] =
     '{
       val fullString = buildInterpolatedString($sc.parts, $args.map(_.toString))
       RefType.parse(fullString).left.map(err => XLError.InvalidReference(fullString, err))
     }
   ```

**Testing requirements**:

```scala
// Compile-time tests
test("ref with all literal interpolation validates at compile time") {
  val sheet = "Sales"  // Literal constant
  val result = ref"${sheet}!A1"  // Should compile-time validate and emit constant
  assertEquals(result, RefType.QualifiedCell(SheetName("Sales"), cell"A1"))
}

// Runtime tests
test("ref with dynamic interpolation returns Either") {
  val sheetName = getUserInput()  // Runtime variable
  val result = ref"$sheetName!A1"

  result match
    case Right(RefType.QualifiedCell(name, ref)) =>
      // Success
    case Left(XLError.InvalidReference(input, err)) =>
      // Expected for invalid input
}

// Absolute reference tests
test("ref with absolute column and row") {
  val abs = ref"$$A$$1"
  assert(abs.isColAbsolute)
  assert(abs.isRowAbsolute)
  assertEquals(abs.toA1, "$A$1")
}

// Edge cases
test("ref with empty interpolation") {
  val empty = ""
  val result = ref"$empty!A1"  // Invalid: empty sheets name
  assert(result.isLeft)
}

test("ref with special characters in sheets name") {
  val sheet = "Q1 'Sales'"
  val result = ref"$sheet!A1"  // Should handle escaping
  assert(result.isRight)
}
```

### 2. `money` Macro (Currency Formatting)

**Supported patterns**:

```scala
// Compile-time
money"$$1,234.56"              // Formatted(Number(1234.56), NumFmt.Currency)

// Runtime interpolation
val amount = 1234.56
money"$$$amount"               // Either[XLError, Formatted]

// With explicit decimals
money"$$$amount" + ".00"       // Format with 2 decimals (future enhancement)
```

**Implementation location**: `xl-core/src/com/tjclp/xl/macros/FormattedLiterals.scala`

**Changes required**:

1. **Add interpolation support to `money` macro**:
   ```scala
   // Current: No interpolation
   inline def money(): Formatted = ${ moneyImpl('sc) }

   // New: Hybrid
   transparent inline def money(inline args: Any*): Formatted | Either[XLError, Formatted] =
     ${ moneyImplWithArgs('sc, 'args) }
   ```

2. **Parse and validate interpolated amounts**:
   ```scala
   def moneyImplWithArgs(sc: Expr[StringContext], args: Expr[Seq[Any]])(using Quotes): Expr[...] =
     val parts = sc.valueOrAbort.parts

     if parts.length == 1 then
       // No interpolation - existing logic
       moneyImpl(sc)
     else
       // Check for all-literal case
       args match
         case Varargs(argExprs) =>
           val literals = argExprs.flatMap(_.value)

           if literals.length == argExprs.length then
             // Compile-time: reconstruct and validate
             val fullString = interleave(parts, literals.map(_.toString))
             validateMoneyLiteral(fullString)  // Aborts on errors
           else
             // Runtime: parse and wrap in Either
             '{
               val fullString = buildInterpolatedString($sc.parts, $args.map(_.toString))
               parseMoneyRuntime(fullString)
             }

   // Runtime parser (pure)
   def parseMoneyRuntime(s: String): Either[XLError, Formatted] =
     // Remove currency symbols, commas
     val cleaned = s.replaceAll("[$$,]", "")
     BigDecimal(cleaned).toRight(XLError.InvalidNumber(s))
       .map(bd => Formatted(CellValue.Number(bd.toDouble), NumFmt.Currency))
   ```

3. **Handle `$$` escaping**:
   - `money"$$1,234"` → literal `$` followed by number → `$1,234`
   - `money"$$$amount"` → literal `$` + interpolated amount → `$1234.56`

**Testing requirements**:

```scala
test("money with compile-time interpolation") {
  val amount = 1234.56  // Literal constant
  val result = money"$$$amount"
  assertEquals(result.value, CellValue.Number(1234.56))
  assertEquals(result.format, NumFmt.Currency)
}

test("money with runtime interpolation") {
  val amount = getUserInput()  // Runtime variable
  val result = money"$$$amount"

  result match
    case Right(Formatted(CellValue.Number(value), NumFmt.Currency)) =>
      // Success
    case Left(XLError.InvalidNumber(input)) =>
      // Expected for invalid input
}

test("money with $$ escape") {
  val literal = money"$$$$100"  // Should parse as "$100"
  assertEquals(literal.value, CellValue.Number(100.0))
}
```

### 3. `date` Macro (Date Literals)

**Supported patterns**:

```scala
// Compile-time
date"2025-11-14"               // Formatted(ExcelDate(45976), NumFmt.Date)

// Runtime interpolation
val year = 2025
val month = 11
val day = 14
date"$year-$month-$day"        // Either[XLError, Formatted]

// Alternative format (future)
date"$month/$day/$year"        // US format
```

**Implementation location**: `xl-core/src/com/tjclp/xl/macros/FormattedLiterals.scala`

**Changes required**:

1. **Add interpolation support**:
   ```scala
   transparent inline def date(inline args: Any*): Formatted | Either[XLError, Formatted] =
     ${ dateImplWithArgs('sc, 'args) }
   ```

2. **Parse interpolated dates**:
   ```scala
   def dateImplWithArgs(sc: Expr[StringContext], args: Expr[Seq[Any]])(using Quotes): Expr[...] =
     // Similar to money macro
     if allLiterals then
       val fullString = reconstruct(parts, literals)
       validateDateLiteral(fullString)  // Uses LocalDate.parse
     else
       '{
         val fullString = buildInterpolatedString($sc.parts, $args.map(_.toString))
         parseDateRuntime(fullString)
       }

   // Runtime parser
   def parseDateRuntime(s: String): Either[XLError, Formatted] =
     Try(LocalDate.parse(s)).toEither
       .left.map(ex => XLError.InvalidDate(s, ex.getMessage))
       .map(ld => Formatted(CellValue.ExcelDate(ExcelDateConverter.toExcelSerial(ld)), NumFmt.Date))
   ```

**Testing requirements**:

```scala
test("date with compile-time interpolation") {
  val year = 2025
  val month = 11
  val day = 14
  val result = date"$year-$month-$day"
  // Should compile-time validate and emit constant
}

test("date with runtime interpolation") {
  val dateString = getUserInput()
  val result = date"$dateString"

  result match
    case Right(Formatted(CellValue.ExcelDate(_), NumFmt.Date)) => // Success
    case Left(XLError.InvalidDate(input, msg)) => // Expected for invalid
}

test("date with invalid format") {
  val invalid = "not-a-date"
  val result = date"$invalid"
  assert(result.isLeft)
}
```

### 4. `fx` Macro (Formulas)

**Supported patterns**:

```scala
// Compile-time
fx"=SUM(A1:A10)"               // CellValue.Formula("=SUM(A1:A10)")

// Runtime interpolation
val range = "A1:A10"
fx"=SUM($range)"               // Either[XLError, CellValue]

// Complex formulas
val ref1 = cell"A1"
val ref2 = cell"B1"
fx"=$ref1 + $ref2"             // Either[XLError, CellValue]

// Function name interpolation
val func = "SUM"
fx"=$func(A1:A10)"             // Either[XLError, CellValue]
```

**Implementation location**: `xl-core/src/com/tjclp/xl/macros/CellRangeLiterals.scala`

**Changes required**:

1. **Add interpolation support**:
   ```scala
   transparent inline def fx(inline args: Any*): CellValue.Formula | Either[XLError, CellValue] =
     ${ fxImplWithArgs('sc, 'args) }
   ```

2. **Validate interpolated formulas**:
   ```scala
   def fxImplWithArgs(sc: Expr[StringContext], args: Expr[Seq[Any]])(using Quotes): Expr[...] =
     if allLiterals then
       val fullString = reconstruct(parts, literals)
       validateFormulaLiteral(fullString)  // Check balanced parens, etc.
     else
       '{
         val fullString = buildInterpolatedString($sc.parts, $args.map(_.toString))
         parseFormulaRuntime(fullString)
       }

   // Runtime parser
   def parseFormulaRuntime(s: String): Either[XLError, CellValue] =
     if !s.startsWith("=") then
       Left(XLError.InvalidFormula(s, "Formula must start with '='"))
     else if !isBalanced(s) then
       Left(XLError.InvalidFormula(s, "Unbalanced parentheses"))
     else
       Right(CellValue.Formula(s))
   ```

**Note**: Formula validation is minimal (balanced parens only). Full formula parsing/evaluation is deferred to P7+ (Formula System).

**Testing requirements**:

```scala
test("fx with compile-time interpolation") {
  val range = "A1:A10"  // Literal constant
  val result = fx"=SUM($range)"
  assertEquals(result, CellValue.Formula("=SUM(A1:A10)"))
}

test("fx with runtime interpolation") {
  val cellRef = getUserInput()
  val result = fx"=$cellRef + 100"

  result match
    case Right(CellValue.Formula(formula)) => // Success
    case Left(XLError.InvalidFormula(input, msg)) => // Expected for invalid
}

test("fx with function name interpolation") {
  val func = "AVERAGE"
  val result = fx"=$func(B1:B10)"
  assertEquals(result, Right(CellValue.Formula("=AVERAGE(B1:B10)")))
}

test("fx with unbalanced parens") {
  val result = fx"=SUM(A1:A10"  // Missing closing paren
  assert(result.isLeft)
}
```

## Purity & Totality Preservation

### Compile-Time Path

**Totality**: Aborts compilation on invalid input (same as current macros)

```scala
val ref = ref"Invalid!@#$"  // Compile errors with helpful message
```

**Purity**: No runtime effects (emits constants directly)

### Runtime Path

**Totality**: Returns `Either[XLError, T]` (no exceptions thrown)

```scala
val result: Either[XLError, ARef] = ref"$userInput"  // Pure, total
```

**Purity**: No side effects, no null, no exceptions

### Error Handling Contract

All runtime interpolations return `Either[XLError, T]`:

```scala
sealed trait XLError:
  case InvalidReference(input: String, reason: String) extends XLError
  case InvalidNumber(input: String) extends XLError
  case InvalidDate(input: String, reason: String) extends XLError
  case InvalidFormula(input: String, reason: String) extends XLError
```

**Usage**:

```scala
// For-comprehension (idiomatic)
for
  ref <- ref"$sheet!$cell"
  sheet <- workbook(ref.sheet)
  cell <- sheet(ref.cell)
yield cell.value

// Pattern matching (explicit)
ref"$userInput" match
  case Right(ref) => processRef(ref)
  case Left(XLError.InvalidReference(input, reason)) =>
    log(s"Invalid reference: $input - $reason")
    fallbackBehavior()
```

### Law Compliance

**Identity Law** (for compile-time path):
```scala
property("ref interpolation identity") {
  forAll { (ref: ARef) =>
    // Compile-time literal should equal runtime-parsed
    val literal = ref"${ref.col.toA1}${ref.row.index1}"
    assertEquals(literal, ref)  // Same constant emitted
  }
}
```

**Round-Trip Law** (for runtime path):
```scala
property("ref interpolation round-trip") {
  forAll { (ref: ARef) =>
    val col = ref.col.toA1
    val row = ref.row.index1
    val result = ref"$col$row"  // Runtime path

    result match
      case Right(parsed) => assertEquals(parsed, ref)
      case Left(err) => fail(s"Should parse valid ref: $err")
  }
}
```

**Associativity** (for multi-part interpolation):
```scala
property("ref interpolation associativity") {
  forAll { (sheet: SheetName, col: Column, row: Row) =>
    // Different groupings should produce same result
    val a = ref"$sheet!${col.toA1}${row.index1}"
    val b = ref"${sheet.value}!${col.toA1}$row"
    assertEquals(a, b)
  }
}
```

## Implementation Phases

### Phase 1: Basic Runtime Support (P7.1)

**Goal**: Enable string interpolation with runtime parsing (no optimization)

**Deliverables**:
1. Remove interpolation blocks in all four macros
2. Implement runtime path for `ref`, `money`, `date`, `fx`
3. Return `Either[XLError, T]` for interpolated calls
4. Basic tests (runtime validation, error handling)

**Estimate**: 2-3 days

**Files modified**:
- `xl-core/src/com/tjclp/xl/macros/RefLiteral.scala` (~100 lines added)
- `xl-core/src/com/tjclp/xl/macros/FormattedLiterals.scala` (~150 lines added)
- `xl-core/src/com/tjclp/xl/macros/CellRangeLiterals.scala` (~50 lines added)
- `xl-core/test/src/com/tjclp/xl/InterpolationSpec.scala` (new file, ~300 lines)

**Success criteria**:
- ✅ All four macros accept interpolation without compile errors
- ✅ Runtime path returns `Either[XLError, T]`
- ✅ Invalid inputs return `Left(XLError.*)` (no exceptions)
- ✅ Basic tests passing (30+ tests)

### Phase 2: Compile-Time Optimization (P7.2)

**Goal**: Detect all-literal interpolations and emit constants (zero overhead)

**Deliverables**:
1. Implement literal detection in all macros (`Expr.value.isDefined`)
2. Reconstruct strings at compile time when all args are literals
3. Emit constants directly (same as non-interpolated literals)
4. Validate at compile time (abort on invalid literals)
5. Add property tests for identity and round-trip laws

**Estimate**: 3-4 days

**Files modified**:
- All macro files (add compile-time path logic)
- `xl-core/test/src/com/tjclp/xl/InterpolationSpec.scala` (add property tests)

**Success criteria**:
- ✅ Compile-time literals emit same code as non-interpolated versions
- ✅ Invalid compile-time literals abort compilation with helpful errors
- ✅ Property tests verify identity law (compile-time = non-interpolated)
- ✅ Zero overhead confirmed (inspect emitted bytecode)

### Phase 3: Absolute Reference Support (P7.3)

**Goal**: Support Excel absolute references (`$A$1`) with `$$` escaping

**Deliverables**:
1. Extend `ARef` to track absolute flags (4 high bits of Long)
2. Update parser to detect `$` prefix on column/row
3. Update printer (`toA1`) to emit `$` for absolute refs
4. Handle `$$` escaping in interpolated strings
5. Add tests for absolute refs (compile-time and runtime)

**Estimate**: 2-3 days

**Files modified**:
- `xl-core/src/com/tjclp/xl/addressing.scala` (extend ARef opaque type)
- `xl-core/src/com/tjclp/xl/macros/RefLiteral.scala` (parser updates)
- `xl-core/test/src/com/tjclp/xl/AddressingSpec.scala` (add tests)
- `xl-core/test/src/com/tjclp/xl/InterpolationSpec.scala` (add tests)

**Success criteria**:
- ✅ `ref"$$A$$1"` creates ARef with both absolute flags set
- ✅ `ref.toA1` prints `$A$1` for absolute refs
- ✅ Mixed absolute/relative refs work: `ref"$$A1"`, `ref"A$$1"`
- ✅ Interpolation with absolute refs: `ref"$$$col$$1"` → `$B$1`
- ✅ Round-trip law holds: `ref.toA1` → `ref"..."` → same `ref`

### Phase 4: Advanced Type Validation (P7.4) [Optional]

**Goal**: Type-specific interpolation positions (e.g., only `String` for sheet names)

**Deliverables**:
1. Define position-specific type constraints
2. Implement typeclass-based insertion (inspired by Contextual library)
3. Emit compile errors for type mismatches
4. Add tests for type safety

**Estimate**: 4-5 days (complex)

**Example**:

```scala
// Only String allowed for sheets name
val sheet: Int = 42
val ref = ref"$sheet!A1"  // Compile errors: "Expected String for sheets name, got Int"

// Only String or ARef allowed for cell reference
val cell: Double = 3.14
val ref = ref"Sales!$cell"  // Compile errors: "Expected String or ARef for cell, got Double"
```

**Note**: This phase is optional and may be deferred indefinitely. The type safety benefit is marginal compared to implementation complexity.

## Testing Requirements

### Unit Tests

**Test file**: `xl-core/test/src/com/tjclp/xl/InterpolationSpec.scala`

**Coverage**:

1. **Compile-time tests** (literal interpolation):
   ```scala
   test("ref with compile-time literal interpolation") {
     val sheet = "Sales"  // Literal constant
     val ref = ref"$sheet!A1"
     assertEquals(ref, RefType.QualifiedCell(SheetName("Sales"), cell"A1"))
   }
   ```

2. **Runtime tests** (dynamic interpolation):
   ```scala
   test("ref with runtime variable interpolation") {
     def getUserSheet(): String = "Sales"  // Simulates runtime value
     val ref = ref"${getUserSheet()}!A1"

     ref match
       case Right(RefType.QualifiedCell(sheet, cell)) =>
         assertEquals(sheet.value, "Sales")
       case Left(err) => fail(s"Should parse valid ref: $err")
   }
   ```

3. **Error handling**:
   ```scala
   test("ref with invalid runtime interpolation returns Left") {
     val invalid = "Not valid!@#$"
     val ref = ref"$invalid"

     assert(ref.isLeft)
     ref match
       case Left(XLError.InvalidReference(input, reason)) =>
         assertEquals(input, invalid)
       case other => fail(s"Expected InvalidReference errors, got: $other")
   }
   ```

4. **Edge cases**:
   ```scala
   test("ref with empty interpolation") {
     val empty = ""
     val ref = ref"$empty!A1"
     assert(ref.isLeft)  // Empty sheets name is invalid
   }

   test("ref with whitespace in sheets name") {
     val sheet = "Q1 Sales"
     val ref = ref"$sheet!A1"
     assert(ref.isRight)  // Whitespace is valid in sheets names
   }

   test("ref with quote escaping") {
     val sheet = "It's Q1"
     val ref = ref"$sheet!A1"  // Should handle internal quotes
     assert(ref.isRight)
   }

   test("money with $$ escape") {
     val literal = money"$$$$100"  // $$ → $, so "$$100"
     assertEquals(literal.value, CellValue.Number(100.0))
   }
   ```

5. **Multi-part interpolation**:
   ```scala
   test("ref with multiple interpolations") {
     val sheet = "Sales"
     val col = "B"
     val row = 42
     val ref = ref"$sheet!$col$row"

     ref match
       case Right(RefType.QualifiedCell(name, cell)) =>
         assertEquals(name.value, "Sales")
         assertEquals(cell.toA1, "B42")
       case Left(err) => fail(s"Should parse valid ref: $err")
   }
   ```

6. **Absolute references**:
   ```scala
   test("ref with absolute column and row") {
     val abs = ref"$$A$$1"
     assert(abs.isColAbsolute)
     assert(abs.isRowAbsolute)
     assertEquals(abs.toA1, "$A$1")
   }

   test("ref with mixed absolute and relative") {
     val mixed = ref"$$A1"  // Absolute column, relative row
     assert(mixed.isColAbsolute)
     assert(!mixed.isRowAbsolute)
     assertEquals(mixed.toA1, "$A1")
   }

   test("ref with interpolated absolute reference") {
     val col = "B"
     val ref = ref"$$$col$$1"  // $$ escape → $B$1

     ref match
       case Right(r: ARef) =>
         assert(r.isColAbsolute)
         assert(r.isRowAbsolute)
         assertEquals(r.toA1, "$B$1")
       case Left(err) => fail(s"Should parse: $err")
   }
   ```

### Property Tests

**Test file**: `xl-core/test/src/com/tjclp/xl/InterpolationSpec.scala`

**Coverage**:

1. **Identity law** (compile-time = non-interpolated):
   ```scala
   property("ref literal interpolation emits same constant") {
     forAll { (ref: ARef) =>
       val literal = ref"${ref.col.toA1}${ref.row.index1}"
       assertEquals(literal, ref)  // Should be identical
     }
   }
   ```

2. **Round-trip law** (parse . print = id):
   ```scala
   property("ref interpolation round-trip") {
     forAll { (ref: ARef) =>
       val col = ref.col.toA1
       val row = ref.row.index1
       val result = ref"$col$row"  // Runtime path

       result match
         case Right(parsed) => assertEquals(parsed, ref)
         case Left(err) => fail(s"Should parse valid ref: $err")
     }
   }

   property("money interpolation round-trip") {
     forAll { (amount: BigDecimal) =>
       val str = amount.toString
       val result = money"$$$str"

       result match
         case Right(Formatted(CellValue.Number(value), NumFmt.Currency)) =>
           assertEquals(value, amount.toDouble, 0.01)
         case Left(err) => fail(s"Should parse valid amount: $err")
     }
   }
   ```

3. **Associativity** (different groupings → same result):
   ```scala
   property("ref multi-part interpolation associativity") {
     forAll { (sheet: SheetName, col: Column, row: Row) =>
       val a = ref"$sheet!${col.toA1}${row.index1}"
       val b = ref"${sheet.value}!${col.toA1}$row"
       assertEquals(a, b)
     }
   }
   ```

4. **Error consistency** (invalid input → Left):
   ```scala
   property("ref with invalid input always returns Left") {
     forAll(Gen.asciiStr.filterNot(isValidRef)) { invalidStr =>
       val result = ref"$invalidStr"
       assert(result.isLeft)
     }
   }
   ```

### Integration Tests

Test interpolation in realistic workflows:

```scala
test("interpolation in for-comprehension") {
  val sheetName = "Sales"
  val cellRef = "A1"

  val result = for
    ref <- ref"$sheetName!$cellRef"
    sheet <- workbook(ref.sheet)
    cell <- sheet(ref.cell)
  yield cell.value

  assert(result.isRight)
}

test("interpolation with batch updates") {
  val amounts = List(1000.0, 2000.0, 3000.0)

  val updates = amounts.zipWithIndex.flatMap { case (amount, idx) =>
    val row = idx + 1
    money"$$$amount".toOption.map { formatted =>
      ref"A$row".toOption.map { ref =>
        ref -> formatted.value
      }
    }.flatten
  }.flatten

  // All amounts should parse successfully
  assertEquals(updates.length, amounts.length)
}
```

### Test Coverage Goals

- **Phase 1**: 30+ tests (basic runtime, error handling, edge cases)
- **Phase 2**: 20+ property tests (identity, round-trip, associativity)
- **Phase 3**: 15+ tests (absolute refs, escaping, mixed)
- **Phase 4**: 10+ tests (type validation, compile errors)

**Total**: 75+ tests covering all interpolation features

## Code Examples

### Example 1: Dynamic Sheet Navigation

```scala
import com.tjclp.xl.*

def getCellValue(
  workbook: Workbook,
  sheetName: String,
  cellRef: String
): Either[XLError, CellValue] =
  for
    // Parse user input with validation
    ref <- ref"$sheetName!$cellRef"  // Returns Either[XLError, RefType]

    // Extract sheets and cell
    (sheet, cell) <- ref match
      case RefType.QualifiedCell(name, cell) => Right((name, cell))
      case _ => Left(XLError.InvalidReference(cellRef, "Expected qualified cell"))

    // Lookup sheets and cell
    sheetObj <- workbook(sheet.value)
    cellObj <- sheetObj(cell)
  yield cellObj.value

// Usage
val result = getCellValue(wb, "Q1 Sales", "B42")
result match
  case Right(CellValue.Number(value)) => println(s"Value: $value")
  case Right(other) => println(s"Non-numeric value: $other")
  case Left(XLError.InvalidReference(input, reason)) =>
    println(s"Invalid reference '$input': $reason")
  case Left(other) => println(s"Error: $other")
```

### Example 2: Programmatic Formula Generation

```scala
def createSumColumn(
  sheet: Sheet,
  dataColumn: Column,
  startRow: Row,
  endRow: Row,
  resultColumn: Column
): Either[XLError, Sheet] =
  // Create formula for each row: =SUM(A1:A10)
  val formulas = (startRow.index1 to endRow.index1).map { rowNum =>
    val rangeStart = dataColumn.toA1 + startRow.index1.toString
    val rangeEnd = dataColumn.toA1 + endRow.index1.toString
    val resultRef = resultColumn.toA1 + rowNum.toString

    for
      formula <- fx"=SUM($rangeStart:$rangeEnd)"  // Dynamic formula
      resultCell <- ref"$resultRef"                // Dynamic cell reference
    yield (resultCell, formula)
  }

  // Collect all formulas (fail fast on first errors)
  formulas.toList.sequence.map { updates =>
    updates.foldLeft(sheet) { case (s, (ref, formula)) =>
      s.put(ref -> formula).getOrElse(s)  // Apply formula
    }
  }

// Usage
val result = createSumColumn(sheet, Column.from1(1), Row.from1(1), Row.from1(10), Column.from1(2))
```

### Example 3: CSV Import with Dynamic Headers

```scala
case class CSVRow(data: Map[String, String])

def importCSV(
  sheet: Sheet,
  rows: List[CSVRow],
  columnMapping: Map[String, Column]  // Header name → Excel column
): Either[XLError, Sheet] =
  rows.zipWithIndex.foldLeft(Right(sheet): Either[XLError, Sheet]) {
    case (Right(currentSheet), (row, idx)) =>
      val excelRow = Row.from1(idx + 2)  // Row 1 is headers

      // Create updates for all columns
      val updates = columnMapping.flatMap { case (header, col) =>
        row.data.get(header).flatMap { value =>
          val colStr = col.toA1
          val rowStr = excelRow.index1.toString

          // Dynamic cell reference
          ref"$colStr$rowStr".toOption.map { cellRef =>
            cellRef -> CellValue.Text(value)
          }
        }
      }

      // Apply all updates
      currentSheet.put(updates.toSeq: _*)

    case (Left(err), _) => Left(err)  // Fail fast
  }

// Usage
val mapping = Map(
  "Name" -> Column.from1(1),
  "Email" -> Column.from1(2),
  "Revenue" -> Column.from1(3)
)

val csvData = List(
  CSVRow(Map("Name" -> "Alice", "Email" -> "alice@example.com", "Revenue" -> "10000")),
  CSVRow(Map("Name" -> "Bob", "Email" -> "bob@example.com", "Revenue" -> "20000"))
)

val result = importCSV(sheet, csvData, mapping)
```

### Example 4: Financial Dashboard with Dynamic Formatting

```scala
def createRevenueReport(
  sheet: Sheet,
  quarters: List[(String, BigDecimal)]  // (Quarter name, Revenue amount)
): Either[XLError, Sheet] =
  quarters.zipWithIndex.foldLeft(Right(sheet): Either[XLError, Sheet]) {
    case (Right(currentSheet), ((quarter, revenue), idx)) =>
      val row = Row.from1(idx + 2)

      for
        // Label cell: A2, A3, A4...
        labelRef <- ref"A${row.index1}"

        // Amount cell: B2, B3, B4...
        amountRef <- ref"B${row.index1}"

        // Formatted amount with currency
        formatted <- money"$$$revenue"

        // Apply updates
        withLabel <- currentSheet.put(labelRef -> CellValue.Text(quarter))
        withAmount <- withLabel.put(amountRef -> formatted.value)
      yield withAmount

    case (Left(err), _) => Left(err)
  }

// Usage
val quarters = List(
  ("Q1 2025", BigDecimal("1000000.50")),
  ("Q2 2025", BigDecimal("1250000.75")),
  ("Q3 2025", BigDecimal("1500000.25")),
  ("Q4 2025", BigDecimal("1750000.00"))
)

val result = createRevenueReport(sheet, quarters)
```

### Example 5: Absolute Reference for Formulas

```scala
def createPercentageColumn(
  sheet: Sheet,
  dataColumn: Column,
  totalCell: ARef,  // Fixed cell with total (e.g., B10)
  startRow: Row,
  endRow: Row,
  resultColumn: Column
): Either[XLError, Sheet] =
  // Create formula: =A1/$B$10 (relative data cell, absolute total cell)
  val formulas = (startRow.index1 to endRow.index1).map { rowNum =>
    val dataRef = dataColumn.toA1 + rowNum.toString
    val totalRef = totalCell.toA1  // Already has $ if absolute
    val resultRef = resultColumn.toA1 + rowNum.toString

    for
      // Use absolute reference for total
      formula <- fx"=$dataRef/$$$totalRef"  // $$ escapes to $B$10
      resultCell <- ref"$resultRef"
    yield (resultCell, formula)
  }

  formulas.toList.sequence.map { updates =>
    updates.foldLeft(sheet) { case (s, (ref, formula)) =>
      s.put(ref -> formula).getOrElse(s)
    }
  }

// Usage
val totalCell = ref"$$B$$10"  // Absolute reference to total
val result = createPercentageColumn(
  sheet,
  Column.from1(1),  // Data in column A
  totalCell,        // Total in B10
  Row.from1(2),     // Start row
  Row.from1(9),     // End row
  Column.from1(3)   // Results in column C
)
```

## Migration Guide

### For Existing Code

**No breaking changes**: Existing macros continue to work identically.

```scala
// Before: Works
val ref = ref"A1"

// After: Still works identically
val ref = ref"A1"
```

### For New Code Using Interpolation

**Pattern 1: Handle Either Explicitly**

```scala
// New: Interpolation with errors handling
val sheetName = getUserInput()
val result = ref"$sheetName!A1"  // Returns Either[XLError, RefType]

result match
  case Right(ref) => processRef(ref)
  case Left(err) => handleError(err)
```

**Pattern 2: Use for-comprehension**

```scala
// New: Chain operations with Either
for
  ref <- ref"$sheetName!$cellRef"
  sheet <- workbook(ref.sheet)
  cell <- sheet(ref.cell)
yield cell.value
```

**Pattern 3: Convert to Option**

```scala
// New: Ignore errors with toOption
val maybeRef = ref"$userInput".toOption

maybeRef match
  case Some(ref) => processRef(ref)
  case None => handleMissing()
```

### Documentation Updates

**Update**: `docs/reference/examples.md`

Add section:

```markdown
## String Interpolation

XL macros support string interpolation for dynamic values:

### Basic Interpolation

\`\`\`scala
val sheet = "Sales"
val cell = "A1"
val ref = ref"$sheet!$cell"  // Returns Either[XLError, RefType]
\`\`\`

### Error Handling

Always handle the Either result:

\`\`\`scala
ref"$userInput" match
  case Right(ref) => // Use ref
  case Left(XLError.InvalidReference(input, reason)) =>
    // Handle error
\`\`\`

### Absolute References

Use `$$` to escape dollar signs for Excel absolute references:

\`\`\`scala
val abs = ref"$$A$$1"        // Excel: $A$1 (absolute column and row)
val mixed = ref"$$A1"         // Excel: $A1 (absolute column only)
val dynamic = ref"$$$col$$1"  // Excel: $B$1 (if col = "B")
\`\`\`

### Compile-Time Optimization

When all interpolated values are literals, validation happens at compile time:

\`\`\`scala
val sheet = "Sales"  // Literal constant
val ref = ref"$sheet!A1"  // Validated at compile time, emits constant
\`\`\`

This provides zero runtime overhead while maintaining type safety.
```

## Open Questions & Future Work

### Open Questions

1. **Type Ascription UX**: Is it acceptable for users to explicitly type `Either[XLError, T]` for runtime interpolations?
   - **Mitigation**: Provide clear documentation and examples
   - **Alternative**: Always return union type (e.g., `ARef | Either[XLError, ARef]`)

2. **Compile Error Messages**: How clear are the error messages when literal interpolation fails?
   - **Test**: Gather user feedback on error message quality
   - **Improve**: Add "Did you mean?" suggestions for common typos

3. **Performance Impact**: Does runtime parsing add measurable overhead?
   - **Benchmark**: Compare `ref"$x"` vs `ARef.parse(x)`
   - **Optimize**: If needed, cache parsed results

4. **Excel Compatibility**: Do users understand `$$` escaping for absolute refs?
   - **Document**: Provide clear examples and migration guide
   - **Test**: Gather user feedback on confusion level

### Future Enhancements

1. **Custom Interpolators** (P8+)
   - User-defined interpolators for domain-specific formats
   - Example: `ssn"$dyn"` for Social Security Number validation

2. **Position-Specific Type Validation** (P7.4)
   - Compile-time type checking for interpolation positions
   - Example: `ref"$sheet!$cell"` where `sheet: String`, `cell: String | ARef`

3. **Advanced Formula Parsing** (P7+)
   - Full Excel formula grammar validation
   - Type-safe formula AST construction
   - Integration with formula evaluator

4. **Localization Support** (P9+)
   - Locale-aware number/date parsing
   - Example: `date"$day/$month/$year"` vs `date"$month/$day/$year"`

5. **Multi-Currency Support** (P9+)
   - `money"€$amount"` for Euro formatting
   - `money"¥$amount"` for Yen formatting
   - Custom currency symbols

6. **Performance Optimizations**
   - Memoization for repeated runtime parsing
   - Fast-path for common patterns (e.g., `ref"A$row"`)
   - Specialized bytecode emission

### Related Features

- **P7: Formula System**: Full formula parsing/evaluation with string interpolation
- **P8: Named Tuples**: Derive interpolators from case classes
- **P9: Internationalization**: Locale-aware parsing for dates/numbers
- **P11: Security**: Validate interpolated formulas for injection attacks

## Summary

String interpolation support for XL macros provides a powerful, ergonomic API for working with dynamic values while maintaining XL's core principles of purity, totality, and type safety.

**Key Design Decisions**:

1. **Hybrid Strategy A**: Single API with compile-time optimization and runtime fallback
2. **All Macros**: Support `ref`, `money`, `date`, and `fx` interpolation
3. **`$$` Escaping**: Standard Scala approach for Excel absolute references
4. **Explicit Errors**: Runtime path returns `Either[XLError, T]` (no exceptions)
5. **Phased Implementation**: Start simple (runtime only), optimize later (compile-time detection)

**Trade-offs Accepted**:

- Type instability (returns `ARef` vs `Either[XLError, ARef]`) for unified API
- `$$` escape may require user education (mitigated with clear documentation)
- More complex implementation than separate `refDyn` macro (justified by ergonomics)

**Next Steps**:

1. Review this plan with stakeholders
2. Prioritize phases (defer to post-1.0 or implement in P7)
3. Create GitHub issues for each phase
4. Implement Phase 1 (basic runtime support) as prototype
5. Gather user feedback before committing to full implementation

---

**Document Metadata**:
- **Author**: Generated by Claude Code (Anthropic)
- **Date**: 2025-11-14
- **Version**: 1.0 (Initial draft)
- **Status**: Awaiting review and approval
