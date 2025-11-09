
# Appendix — Benchmarks Plan

## JMH benches
- Parse 100k‑row sheet (no styles) → measure allocs/op and throughput.
- Write 100k‑row sheet with SST hot cache.
- Patch application microbench (single cell, full row, full column).

## Comparisons
- POI streaming (SXSSF) baseline for read/write, comparable features only.

## Metrics
- Throughput (rows/s), p50/p95 latencies, max RSS, GC pause counts.
