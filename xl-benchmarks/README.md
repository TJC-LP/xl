# XL Benchmarks

JMH-based performance benchmarking suite for the XL Excel library.

## Overview

This module provides comprehensive benchmarks to:
- Validate performance claims (4.5x faster than POI, 16x less memory)
- Enable performance regression detection in CI
- Establish baseline metrics for optimization work
- Compare XL vs Apache POI on equivalent operations

## Running Benchmarks

### All Benchmarks
```bash
./mill xl-benchmarks.runJmh
```

### List Available Benchmarks
```bash
./mill xl-benchmarks.listJmhBenchmarks
```

### Run Specific Benchmark Suite
```bash
./mill xl-benchmarks.runJmh ".*ReadWrite.*"    # Read/Write operations
./mill xl-benchmarks.runJmh ".*Patch.*"        # Patch operations
./mill xl-benchmarks.runJmh ".*Style.*"        # Style operations
./mill xl-benchmarks.runJmh ".*Poi.*"          # XL vs POI comparison
```

### Run Single Benchmark
```bash
./mill xl-benchmarks.runJmh "singleCellUpdate"
```

### JMH Options
```bash
# List all JMH options
./mill xl-benchmarks.runJmh -h

# Custom warmup/measurement
./mill xl-benchmarks.runJmh -wi 5 -i 10 ".*ReadWrite.*"

# Enable GC profiler (memory allocation tracking)
./mill xl-benchmarks.runJmh -prof gc ".*ReadWrite.*"
```

## Benchmark Suites

### 1. ReadWriteBenchmark
Tests read/write throughput for varying row counts (1k, 10k, 100k).

**Benchmarks**:
- `writeWorkbook` - Write workbook to disk
- `readWorkbook` - Read workbook from disk
- `roundTrip` - Full write + read cycle

**Parameters**: `rows` = 1000, 10000, 100000

### 2. PatchBenchmark
Microbenchmarks for Patch operations and Monoid composition.

**Benchmarks**:
- `singleCellUpdate` - Single cell update overhead
- `rowUpdate` - Update 1000 cells in a row
- `columnUpdate` - Update 10000 cells in a column
- `patchComposition` - Patch combining (Monoid operations)

### 3. StyleBenchmark
Tests style system performance (deduplication, canonicalKey, application).

**Benchmarks**:
- `canonicalKeyComputation` - CellStyle.canonicalKey timing
- `registerUniqueStyle` - Register new style (deduplication miss)
- `registerDuplicateStyle` - Register existing style (deduplication hit)
- `registerManyUniqueStyles` - 1000 unique styles
- `registerManyDuplicateStyles` - 1000 duplicate styles (perfect deduplication)
- `styleIndexLookup` - StyleRegistry.indexOf performance
- `applyStyleToRange` - Apply style to 2600 cells (A1:Z100)

### 4. PoiComparisonBenchmark
XL vs Apache POI comparison on equivalent operations.

**Benchmarks**:
- `xlWrite` - XL write performance
- `poiWrite` - POI write performance
- `xlRead` - XL read performance
- `poiRead` - POI read performance

**Parameters**: `rows` = 1000, 10000

## Interpreting Results

JMH outputs results in the following format:

```
Benchmark                                 (rows)  Mode  Cnt   Score    Error  Units
ReadWriteBenchmark.writeWorkbook            1000  avgt    5  45.123 ±  2.456  ms/op
ReadWriteBenchmark.writeWorkbook           10000  avgt    5 123.456 ±  8.901  ms/op
```

- **Score**: Average time per operation (lower is better)
- **Error**: 99.9% confidence interval
- **Units**: Typically ms/op (milliseconds per operation), μs/op (microseconds), or ns/op (nanoseconds)

### Actual Performance Results

Benchmarked on Apple Silicon (M-series), JDK 25:

#### Streaming Reads (SAX Parser - Production Recommendation)
| Rows | POI | XL | Result |
|------|-----|----|--------|
| **1,000** | 1.357 ± 0.076 ms | **0.887 ± 0.060 ms** | ✨ **XL 35% faster** |
| **10,000** | 7.773 ± 0.590 ms | 8.408 ± 0.153 ms | Competitive (XL within 8%) |

#### In-Memory Reads (For Modification Workflows)
| Rows | POI | XL | Result |
|------|-----|----|--------|
| **1,000** | 1.650 ± 0.055 ms | **1.225 ± 0.086 ms** | ✨ **XL 26% faster** |
| **10,000** | 13.784 ± 0.377 ms | 14.115 ± 1.250 ms | Competitive (XL within 2%) |

#### Writes
| Rows | POI | XL | Result |
|------|-----|----|--------|
| **1,000** | 1.280 ± 0.041 ms | 1.906 ± 0.245 ms | POI 49% faster |
| **10,000** | 10.228 ± 0.417 ms | 15.248 ± 1.315 ms | POI 49% faster |

**Validated Methodology**:
- **Fair comparison**: Both libraries read identical shared file
- **Verifiable data**: Arithmetic series (A{i} = i) enables sum validation
- **No JIT artifacts**: Benchmarks return computed sum (500,500 for 1k rows, 50,005,000 for 10k rows)
- **Write cost isolated**: Files pre-created in @Setup, not measured in read benchmarks

**Key Findings**:
- ✨ **XL is fastest for small-medium files** (< 5k rows): 35% faster streaming, 26% faster in-memory
- ✅ **XL competitive on large files**: Within 8% of POI on 10k row streaming reads
- 🔧 **Write optimization**: Future work (Phase 3) - POI currently 49% faster
- 💾 **Constant memory**: Streaming uses O(1) memory regardless of file size
- ⚡ **SAX parser**: 3.8x speedup vs previous fs2-data-xml implementation

**Recommendation**: Use `ExcelIO.readStream()` for production workloads (fastest for <5k rows, constant memory). Reserve `ExcelIO.read()` for random access + modification scenarios.

## CI Integration

*(To be added in WI-16)*

Benchmarks will run on PRs targeting main branch with:
- Baseline comparison against main branch
- Regression detection (fail if >10% slowdown)
- Performance trend tracking

## Development

### Adding New Benchmarks

1. Create new benchmark class in `xl-benchmarks/src/com/tjclp/xl/benchmarks/`
2. Extend with JMH annotations (`@Benchmark`, `@State`, `@Param`)
3. Suppress WartRemover warnings: `@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Null"))`
4. Use `scala.compiletime.uninitialized` for JMH fields
5. Run `./mill xl-benchmarks.listJmhBenchmarks` to verify detection

### Example Benchmark

```scala
package com.tjclp.xl.benchmarks

import org.openjdk.jmh.annotations.*
import java.util.concurrent.TimeUnit
import scala.compiletime.uninitialized

@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.AverageTime))
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1, timeUnit = TimeUnit.SECONDS)
@Measurement(iterations = 5, time = 2, timeUnit = TimeUnit.SECONDS)
@Fork(value = 1)
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Null"))
class MyBenchmark {

  var data: MyData = uninitialized

  @Setup(Level.Trial)
  def setup(): Unit = {
    data = generateTestData()
  }

  @Benchmark
  def myOperation(): Result = {
    performOperation(data)
  }
}
```

## Notes

- Benchmarks use fixed seeds (`Random(42)`) for reproducibility
- WartRemover warnings are acceptable for benchmark code (performance-critical)
- JMH handles statistical analysis (warm-up, measurement iterations, forking)
- Use `uninitialized` for JMH-managed fields (Scala 3 requirement)

## References

- [JMH Documentation](https://github.com/openjdk/jmh)
- [Mill JMH Module](https://mill-build.org/mill/contrib/jmh.html)
- [XL Performance Guide](../docs/reference/performance-guide.md)
