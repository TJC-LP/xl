# Testing & Laws — Property Suites, Round-Trips, and Coverage

**Current Status**: CI runs the full Mill test graph across the library, evaluator, CLI, and support modules — **4,085 tests** as of 0.12.6. Use `./mill __.test` as the authoritative count.

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
xl-evaluator/test/src/com/tjclp/xl/formula/
xl-cli/test/src/com/tjclp/xl/cli/
xl-agent/test/src/
xl/test/src/xlprelude/          (external-consumer probes for the scripting prelude)
```

## Test Coverage by Module

### xl-core: domain, style, codec, optics, and law suites ✅

#### Addressing Laws
- **Column/Row round-trips**: `from0` → `index0` identity
- **ARef packing**: 64-bit (row, col) ↔ Long bijection
- **A1 notation**: `parse` → `toA1` round-trip for all valid refs
- **CellRange normalization**: Constructor always produces `start ≤ end`
- **Range contains**: Point-in-rectangle tests
- **Property-based**: Generated Column (0-16383), Row (0-1048575)

#### Patch Laws
- **Monoid associativity**: `(p1 |+| p2) |+| p3 == p1 |+| (p2 |+| p3)`
- **Monoid identity**: `p |+| Patch.Empty == p`
- **Idempotence**: `Put(ref, v1) |+| Put(ref, v2)` keeps v2
- **Application semantics**: `applyPatch` updates sheet correctly
- **Error handling**: Invalid patches return `Left[XLError]`
- **Batch composition**: `Patch.Batch` flattens correctly

#### Style System
- **Unit conversions**: `Pt ↔ Px ↔ Emu` bidirectional laws
- **Color parsing**: Hex, RGB, ARGB, theme color parsing
- **Style canonicalization**: `canonicalKey` idempotence
- **StylePatch Monoid**: Composition laws verified
- **StyleRegistry**: Per-sheet registration and lookup
- **Font/Fill/Border**: Builder patterns and deduplication
- **NumFmt**: Pre-defined format IDs and custom formats

#### DateTime
- **Excel serial numbers**: LocalDate ↔ Double conversion
- **1900 leap year bug**: Compatibility with Excel's quirk
- **DateTime round-trips**: LocalDateTime ↔ serial + fraction
- **Epoch correctness**: 1899-12-30 baseline

#### CellCodec
- **Identity laws**: `read(write(v)) == Right(Some(v))` for all 9 types
- **Type safety**: `readTyped[Wrong]` returns `Left[CodecError]`
- **Auto-formatting**: LocalDate → NumFmt.Date, BigDecimal → NumFmt.Decimal
- **Primitive codecs**: String, Int, Long, Double, BigDecimal, Boolean, LocalDate, LocalDateTime, RichText
- **Error cases**: Parse failures, type mismatches
- **Round-trip precision**: BigDecimal maintains precision

#### Batch Operations
- **Batch `Sheet.put`**: Heterogeneous updates with type-safe codecs (replaces old `putMixed`)
- **Style deduplication**: Multiple cells with same format share style
- **Given conversions**: Implicit codec resolution
- **Type-safe reading**: `readTyped[A]` with compile-time types

- **Given conversions**: `sheet.put(ref"A1", "text")` without wrappers
- **Batch put**: `sheet.put(ref"A1" -> "Name", ref"B1" -> 42)` via codecs
- **Formatted literals**: `money"$1,234.56"`, `percent"45.5%"`, `date"2025-11-10"`, `accounting"-$500.00"`
- **Macro expansion**: Compile-time parsing verified

#### Optics
- **Lens laws**: `get ∘ set == id`, `set(get(s), s) == s`, `set(set(s, a), b) == set(s, b)`
- **Optional laws**: Similar to Lens for partial updates
- **Focus DSL**: `sheet.focus(ref).modify(f)` correctness
- **Real-world scenarios**: Invoice updates, financial model edits
- **Compose**: Lens/Optional composition verified

#### RichText
- **Composition**: `run1 + run2` combines correctly
- **DSL extensions**: `.bold`, `.italic`, `.red`, `.size()` work
- **Whitespace**: Preserved correctly with `xml:space="preserve"`
- **OOXML mapping**: `TextRun` → `<r><rPr>...</rPr><t>...</t></r>`

### xl-ooxml: OOXML round-trip, surgical write, metadata, table, security, and performance suites ✅

#### Round-Trip and Regression Tests
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

### xl-cats-effect: streaming and effectful IO suites ✅

#### Streaming I/O
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

## Golden File Tests

The golden corpus that exists pins the CLI's output, not workbook bytes: see "CLI contract
goldens" below. A curated `.xlsx` corpus with normalized XML diffs remains future work (P11).

## CLI contract goldens

`xl-cli/test/resources/golden/<case>.golden` pins what an agent sees from the `xl` binary — exit
code, stdout and stderr, separately — for representative invocations: `--help`/`--version`, parse
errors (and the acceptance of a global flag after the verb), the read verbs in every text format,
write verbs with and without `-o`, `batch -` fed through stdin, `diff` and `lint` exit codes,
`-i recalc`, and `--stream`. `GoldenSpec` (`com.tjclp.xl.cli.contract`) runs each case through `CliHarness`, an
in-process harness that drives the real parser and handlers (`Cli.run(args, io)`) with the
production sinks, redirected JVM streams and an injected stdin, and compares the result with the
file after normalization: the per-run fixture directory becomes `<DIR>`, any other temp path
`<TMP>`, the build version `<VERSION>`, trailing whitespace is trimmed. Fixtures are built
in-process by `TestFixtures` on every run — nothing binary is checked in.

File format, one `## ` header per section: `args` (one argv token per line), optional `stdin`,
`exit`, `stdout`, `stderr`.

```bash
./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.GoldenSpec       # verify
XL_UPDATE_GOLDEN=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.GoldenSpec   # (re)record
XL_GOLDEN_KEEP=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.GoldenSpec     # keep fixtures, print the dir
```

To add a case, write `<name>.golden` with its `## args` (and `## stdin`) and record. A mismatch
fails with a unified diff in the assertion message.

**Review discipline**: a golden diff is an agent-visible contract change, never noise. Read the
diff, decide whether the change is intended, and only then re-record — and add a CHANGELOG line
under `## [Unreleased]` saying what agents now see differently. Re-recording to turn a red build
green without that line is how a contract drifts unnoticed.

## Test Execution

### Run All Tests
```bash
./mill __.test                 # All test modules
./mill xl-core.test            # Core only
./mill xl-ooxml.test           # OOXML only
./mill xl-cats-effect.test     # Streaming/IO only
./mill xl-evaluator.test       # Formula parser/evaluator only
./mill xl-cli.test             # CLI only
```

### Run Specific Test
```bash
./mill xl-core.test.testOnly com.tjclp.xl.AddressingSpec
./mill xl-core.test.testOnly com.tjclp.xl.PatchSpec
./mill xl-ooxml.test.testOnly com.tjclp.xl.ooxml.OoxmlRoundTripSpec
```

### CI Integration
GitHub Actions runs:
1. `./mill mill.scalalib.scalafmt.ScalafmtModule/checkFormatAll __.sources` (all-source Scalafmt verification)
2. `./mill __.compile` (Compilation check)
3. `./mill __.test` (All test modules)

## Test Counts by Module

As of the #606 blind-reader fix (2026-09-08), from the full-suite JUnit reports; macros are part of xl-core. One existing style-performance comparison is ignored; four subprocess smokes skip when openpyxl is absent or the sandbox runs as root.

| Module | Tests |
|--------|-------|
| xl-evaluator | 2374 |
| xl-core | 1533 |
| xl-ooxml | 1119 |
| xl-cli | 1264 |
| xl-cats-effect | 167 |
| xl-agent | 145 |
| xl (prelude probes, `xlprelude.ScriptingPreludeTest`) | 50 |
| **Total** | **6,652** |

## Test Quality Metrics

- **All tests pass**: enforced by `./mill __.test` in CI ✅
- **Zero flaky tests**: Deterministic, reproducible
- **Fast execution**: focused suites run quickly; full-suite time is tracked by CI
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
