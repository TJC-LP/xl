# Testing & Laws — Property Suites, Round-Trips, and Coverage

**Current Status**: 636/636 tests passing across 4 modules

## Test Infrastructure

### Frameworks
- **MUnit** - Primary test framework
- **MUnit ScalaCheck** - Property-based testing integration
- **ScalaCheck Generators** - Custom generators in `Generators.scala`

### Test Location
```
xl-core/test/src/com/tjclp/xl/
xl-ooxml/test/src/com/tjclp/xl/ooxml/
xl-cats-effect/test/src/com/tjclp/xl/io/
```

## Test Coverage by Module

### xl-core: 221 tests ✅

#### Addressing Laws (17 tests)
- **Column/Row round-trips**: `from0` → `index0` identity
- **ARef packing**: 64-bit (row, col) ↔ Long bijection
- **A1 notation**: `parse` → `toA1` round-trip for all valid refs
- **CellRange normalization**: Constructor always produces `start ≤ end`
- **Range contains**: Point-in-rectangle tests
- **Property-based**: Generated Column (0-16383), Row (0-1048575)

#### Patch Laws (21 tests)
- **Monoid associativity**: `(p1 |+| p2) |+| p3 == p1 |+| (p2 |+| p3)`
- **Monoid identity**: `p |+| Patch.Empty == p`
- **Idempotence**: `Put(ref, v1) |+| Put(ref, v2)` keeps v2
- **Application semantics**: `applyPatch` updates sheet correctly
- **Error handling**: Invalid patches return `Left[XLError]`
- **Batch composition**: `Patch.Batch` flattens correctly

#### Style System (60 tests)
- **Unit conversions**: `Pt ↔ Px ↔ Emu` bidirectional laws
- **Color parsing**: Hex, RGB, ARGB, theme color parsing
- **Style canonicalization**: `canonicalKey` idempotence
- **StylePatch Monoid**: Composition laws verified
- **StyleRegistry**: Per-sheet registration and lookup
- **Font/Fill/Border**: Builder patterns and deduplication
- **NumFmt**: Pre-defined format IDs and custom formats

#### DateTime (8 tests)
- **Excel serial numbers**: LocalDate ↔ Double conversion
- **1900 leap year bug**: Compatibility with Excel's quirk
- **DateTime round-trips**: LocalDateTime ↔ serial + fraction
- **Epoch correctness**: 1899-12-30 baseline

#### CellCodec (42 tests)
- **Identity laws**: `read(write(v)) == Right(Some(v))` for all 9 types
- **Type safety**: `readTyped[Wrong]` returns `Left[CodecError]`
- **Auto-formatting**: LocalDate → NumFmt.Date, BigDecimal → NumFmt.Decimal
- **Primitive codecs**: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText
- **Error cases**: Parse failures, type mismatches
- **Round-trip precision**: BigDecimal maintains precision

#### Batch Operations (16 tests)
- **Batch `Sheet.put`**: Heterogeneous updates with type-safe codecs (replaces old `putMixed`)
- **Style deduplication**: Multiple cells with same format share style
- **Given conversions**: Implicit codec resolution
- **Type-safe reading**: `readTyped[A]` with compile-time types

- **Given conversions**: `sheet.put(ref"A1", "text")` without wrappers
- **Batch put**: `sheet.put(ref"A1" -> "Name", ref"B1" -> 42)` via codecs
- **Formatted literals**: `money"$1,234.56"`, `percent"45.5%"`, `date"2025-11-10"`, `accounting"-$500.00"`
- **Macro expansion**: Compile-time parsing verified

#### Optics (34 tests)
- **Lens laws**: `get ∘ set == id`, `set(get(s), s) == s`, `set(set(s, a), b) == set(s, b)`
- **Optional laws**: Similar to Lens for partial updates
- **Focus DSL**: `sheet.focus(ref).modify(f)` correctness
- **Real-world scenarios**: Invoice updates, financial model edits
- **Compose**: Lens/Optional composition verified

#### RichText (5 tests)
- **Composition**: `run1 + run2` combines correctly
- **DSL extensions**: `.bold`, `.italic`, `.red`, `.size()` work
- **Whitespace**: Preserved correctly with `xml:space="preserve"`
- **OOXML mapping**: `TextRun` → `<r><rPr>...</rPr><t>...</t></r>`

### xl-ooxml: 24 tests ✅

#### Round-Trip Tests (24 tests)
- **Text cells**: String values preserve exactly
- **Number cells**: Numeric precision maintained
- **Boolean cells**: True/false round-trip
- **DateTime cells**: Serial number conversion correct
- **Formula cells**: Formula strings preserved (not evaluated)
- **Mixed workbooks**: All cell types in one sheet
- **Multi-sheet**: Multiple sheets with relationships
- **Shared Strings Table**: Deduplication verified
- **Styles**: Style indices match after round-trip
- **RichText**: Multi-run formatted text preserved
- **XML determinism**: Same input → same byte output

### xl-cats-effect: 18 tests ✅

#### Streaming I/O (18 tests)
- **writeStream / writeStreamsSeq**: Event-based ZIP write via fs2-data-xml
- **readStream / readSheetStream / readStreamByIndex**: Event-based worksheet reads with fs2-data-xml + fs2.io.readInputStream
- **Constant memory**: O(1) memory usage verified (100k rows @ ~50MB)
- **Large files**: 100k+ row tests pass
- **Multi-sheet streaming**: Multiple worksheets in one file (true streaming sequence)
- **Shared strings**: SST parsing integrates with streaming readers
- **Style integration**: Minimal styles + default formatting preserved in streaming writes
- **Error handling**: Invalid XML returns `Left[XLError]`
- **Performance**: Benchmarked at 4.5x faster than Apache POI

## Property-Based Testing Patterns

### Generators (xl-core/test/src/com/tjclp/xl/Generators.scala)

All domain types have `Arbitrary` instances:
- `Column` - Valid range [0, 16383]
- `Row` - Valid range [0, 1048575]
- `ARef` - All valid (col, row) pairs
- `CellRange` - Normalized ranges
- `CellValue` - Text, Number, Bool, DateTime, Formula, Error
- `CellStyle` - Valid font/fill/border combinations
- `Patch` - All Patch enum cases

### Test Patterns

#### Round-Trip Laws
```scala
property("parse . print = id") {
  forAll { (ref: ARef) =>
    ARef.parse(ref.toA1) == Right(ref)
  }
}
```

#### Monoid Laws
```scala
property("associativity") {
  forAll { (p1: Patch, p2: Patch, p3: Patch) =>
    ((p1 |+| p2) |+| p3) == (p1 |+| (p2 |+| p3))
  }
}

property("identity") {
  forAll { (p: Patch) =>
    (p |+| Patch.Empty) == p
  }
}
```

#### Lens Laws
```scala
property("get-set") {
  forAll { (s: Sheet, ref: ARef, v: CellValue) =>
    val lens = Optics.valueLens(ref)
    lens.get(lens.set(s, Some(v))) == Some(v)
  }
}
```

## Golden File Tests (Future - P11)

### Planned Infrastructure
- Curated `.xlsx` corpus covering:
  - Edge cases (empty cells, large numbers, special characters)
  - Excel compatibility (2007, 2010, 2013, 2016, 2019, M365)
  - Feature coverage (all cell types, styles, multi-sheet)
- Deterministic XML diff:
  - Normalized attribute ordering
  - Whitespace normalization
  - Stable sort for elements
- Version control:
  - Check in `.xlsx` files with LFS
  - Store expected XML separately

### Not Yet Implemented
- Golden file test framework
- Compatibility test suite
- Visual regression tests (for charts)

## Test Execution

### Run All Tests
```bash
./mill __.test               # All modules (263 tests)
./mill xl-core.test          # Core only (221 tests)
./mill xl-ooxml.test         # OOXML only (24 tests)
./mill xl-cats-effect.test   # Streaming only (18 tests)
```

### Run Specific Test
```bash
./mill xl-core.test.testOnly com.tjclp.xl.AddressingSpec
./mill xl-core.test.testOnly com.tjclp.xl.PatchSpec
./mill xl-ooxml.test.testOnly com.tjclp.xl.ooxml.RoundTripSpec
```

### CI Integration
GitHub Actions runs:
1. `./mill __.checkFormat` (Scalafmt verification)
2. `./mill __.compile` (Compilation check)
3. `./mill __.test` (All 263 tests)

## Coverage Goals

| Module | Lines | Coverage | Status |
|--------|-------|----------|--------|
| xl-core | ~3500 | ~90% | ✅ Excellent |
| xl-macros | ~400 | ~95% | ✅ Excellent |
| xl-ooxml | ~1800 | ~85% | ✅ Good |
| xl-cats-effect | ~500 | ~80% | ✅ Good |
| **Total** | **~6200** | **~88%** | **✅ Excellent** |

## Test Quality Metrics

- **All tests pass**: 263/263 ✅
- **Zero flaky tests**: Deterministic, reproducible
- **Fast execution**: <5 seconds for full suite
- **Property-based**: 60%+ of tests use ScalaCheck
- **Law coverage**: All algebras (Monoid, Lens, Optional) verified
- **Edge cases**: Boundary values, error paths tested
- **Round-trips**: All serialization paths verified

## Future Testing Work (P11)

1. **Golden file framework** - Curated .xlsx corpus with stable diffs
2. **Benchmark suite** - Performance regression tests
3. **Compatibility matrix** - Test across Excel versions
4. **Stress tests** - 1M+ row files, deeply nested formulas
5. **Mutation testing** - Verify test thoroughness with PIT
6. **Visual regression** - For charts/drawings (P8/P9)
