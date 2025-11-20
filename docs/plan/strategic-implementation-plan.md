# Strategic Implementation Plan: Path to Best-in-Class Excel Library

**Status**: Strategic Vision & Execution Framework
**Last Updated**: 2025-11-20
**Purpose**: High-level roadmap with parallelization strategies and dependency management

> **Note**: This document provides strategic vision and execution framework. For granular phase tracking and completion status, see [roadmap.md](roadmap.md) which uses P0-P13 numbering.

---

## Executive Summary

This document outlines the complete path from current state (~87% complete, 680 tests passing) to "best Excel library in the world" with:

- **7 strategic phases** organized by dependencies and value delivery
- **Explicit parallelization opportunities** for concurrent development streams
- **Clear dependencies** to avoid painting ourselves into corners
- **Differentiator-first approach** leveraging streaming, purity, and type safety

---

## 1. The Universe of Remaining Work

From existing plans and limitations, remaining work clusters into major systems:

### A. Core Polish & Streaming

* **P6.5**: Performance & quality polish (error-path tests, round-trip tests, optimizations)
* **Streaming improvements**: SST in streaming write, two-phase writer (P7.5), column/row properties, theme colors
* **Benchmarks**: JMH harness, POI comparisons, regression tests

**Existing plans**: `future-improvements.md`, `streaming-improvements.md`, `benchmarks.md`

### B. AST Completeness & OOXML Coverage

Beyond workbook/worksheet/styles/SST (already complete), remaining OOXML parts:

* **Already complete**: Comments ✅ (full OOXML round-trip, VML drawings)
* Hyperlinks
* Named ranges, calcChain, sheet relationships
* Document properties (core/app)
* Print settings, page setup
* Conditional formatting, data validation
* Tables & pivot tables
* Drawings (images, shapes)
* Charts (including chart relationships & data bindings)

**Note**: Surgical modification already preserves unknown parts; missing piece is typed AST for first-class manipulation.

**Existing plans**: `drawings.md`, `charts.md`, `tables-and-pivots.md`

### C. Developer Experience & API Ergonomics

* Full codec derivation (P6b) / named tuples
* Advanced macros: path macro, style literal
* Type-class-based Easy Mode API (✅ Complete - PR #20)
* Query API extensions (column-name awareness, aggregations, DSL)

**Existing plans**: Design docs in `docs/design/`

### D. Formula System & Dependency Graph

* **Formula evaluator** (P9+): Execution engine, function library, circular reference detection, error propagation
* **Formula AST**: Already exists via parser, but formulas are effectively "strings with a contract"
* **Dependency graph**: Build graph of cells/names/tables + calcChain; support graph queries
  - "All precedents of X"
  - "All dependents of this table"
  - SCCs, topological sort
* **Integration**: Hybrid "row + formula" workflows with Query API

**Existing plans**: `formula-system.md`

### E. Advanced Features (User-Visible Excel Goodies)

* Drawings: images, shapes, anchors, deduplication
* Charts: core chart types, data binding, styling
* Tables & pivots: table XML, structured refs, pivotCaches
* Conditional formatting & data validation

**Existing plans**: `drawings.md`, `charts.md`, `tables-and-pivots.md`

### F. Security, Safety & Robustness

* ZIP bomb detection, XXE prevention, file size limits
* Macro preservation (XLSM, never execution)
* Formula-injection guards
* Golden-file test corpus + compatibility matrix across Excel versions

**Existing plans**: `security.md`, `error-model-and-safety.md`

### G. Architectural Evolution

* Lazy evaluation: logical plan DSL, query optimizer
* `LazySheet` as default representation
* Tighter integration with fs2

**Status**: Deferred (archived in `docs/archive/plan/deferred/lazy-evaluation.md`)

---

## 2. High-Level Principles for Ordering

Optimization priorities:

1. **Never lose data** → IO correctness & AST coverage before fancy APIs
2. **Don't paint into corners** → Formalize AST + relationships before evaluator & graph search
3. **Exploit differentiators early** → Streaming, query, dependency graph, safety
4. **Leave architectural rewrites until stable** → Lazy eval post-1.0

---

## 3. Strategic Phases with Parallelization

### Phase 1 – Hardening & Observability (SHORT, UNBLOCKS EVERYTHING)

**Goal**: Lock in current core as rock-solid with strong tests and clear perf profile.

**Work Items**:
* Finish P6.5: Quality/polish from PR feedback (whitespace utilities, error-path tests, full XLSX round-trip)
* Streaming improvements (low-risk): Compression defaults & configs
* Benchmarks infrastructure (first pass): JMH harness, basic POI vs XL benchmarks, simple perf regression CI

**Parallelization**:
* **Stream A**: P6.5 + streaming-improvements glue
* **Stream B**: Benchmarks/CI
* Low merge conflict risk (different modules: xl-ooxml, xl-cats-effect, benchmarks)

**Maps to existing phases**: P6.5, P6.7 (compression)

---

### Phase 2 – OOXML AST Coverage Spine

**Goal**: Move from "we preserve unknown XML" to "we understand almost everything Excel cares about."

**Work Items**:

1. **Small, high-leverage mappings** (warm-up):
   * Hyperlinks
   * Named ranges, calcChain, sheet relationships
   * Document properties (core/app), theme1.xml if needed

2. **Comments & related structures**:
   * ✅ **Already complete**: Comments AST, VML drawings for comment shapes, full OOXML round-trip
   * No work needed (verified complete in codebase)

3. **Print settings & page setup**:
   * Map current in-memory representation back to XML

**Why Now?**:
* Formula graph and evaluator will want named ranges, calcChain, robust sheet relationships
* Comments were key user feature (now complete)
* Small, isolated changes that build foundation

**Parallelization**:
* **Stream A**: Hyperlinks + document properties + print/page setup
* **Stream B**: Named ranges + calcChain + relationships
* **Stream C**: (Reserved, comments already complete)

Minimal overlap, all depend on existing OOXML mapping & surgical-preservation infrastructure.

**Maps to existing phases**: Extensions of P4 OOXML work

---

### Phase 3 – Developer Experience & Data Model Ergonomics

**Goal**: Make 80% workflows ridiculously pleasant before diving into evaluator & heavy features.

**Work Items**:
* P6b: Full RowCodec / case class derivation + named tuple integration
* ✅ **Already complete**: Easy-mode `put()` API consolidation via type classes (PR #20)
* Advanced macros:
  * `path` macro → named paths to cells/ranges built on relationships + named ranges
  * `style` literal for inline style DSL
* Query API extensions:
  * Column-name aware filters
  * Aggregation helpers & mini DSL for row queries

**Why Now?**:
* Turns XL from "fast, pure, correct" into "fast, pure, correct, and *fun*"
* Clean, expressive surface that formula/graph work can integrate with
* Type-safe references, named paths ready for evaluator

**Parallelization**:
* **Stream A**: P6b + remaining type-class API work
* **Stream B**: Macros & Query API extras

Mostly touch xl-core, xl-macros, xl-cats-effect with minimal OOXML changes.

**Maps to existing phases**: P6b (codecs), P9 (macros)

---

### Phase 4 – Formula Engine & Dependency Graph (GRAPH SEARCH VIA RELATIONSHIPS)

**Goal**: Turn formulas from opaque strings into proper, queryable, evaluatable graph. **This is the core differentiator.**

**Work Items**:

1. **Formula AST & evaluation core**:
   * Solidify formula AST around existing parser
   * Implement evaluator with first tranche of functions (arithmetic, logical, basic date/time, lookup)
   * Integrate with XLError so evaluation failures are principled

2. **Dependency graph & relationships layer**:
   * Build `WorkbookGraph` abstraction: nodes = cells, names, tables; edges = references
   * Derive from formula AST + named ranges + table references + calcChain
   * Expose queries:
     - `precedents(ref)`, `dependents(ref)`
     - Reachability queries (all cells influenced by this input)
     - SCC detection (circular refs)
     - Topological sort for recalculation order

3. **Graph search API**:
   * Pure API (xl-core or dedicated `xl-graph` module)
   * Operations:
     - `graph.findPaths(from, to)`
     - `graph.subgraph(pred)` (e.g., "all formulas using volatile functions")
   * Optional: integrate with Query API for hybrid "row + formula" workflows

**Dependencies**:
* **Requires**: Phase 2 named ranges + calcChain + relationship mapping
* **Doesn't require**: Drawings/charts/tables fully implemented
  - Can start with cell ↔ cell / name graph and extend later

**Parallelization**:
* **Engine team**: Formula evaluation, function library, circular detection
* **Graph team**: Graph builder + query API (uses engine's AST, not necessarily evaluation logic)

**Maps to existing phases**: P9+ (formula-system.md)

---

### Phase 5 – Advanced Features: Tables, Charts, Drawings, Rules

**Goal**: Reach POI-level (or better) feature breadth.

**Work Items**:

1. **Tables & pivots + rules**:
   * Excel tables (structured references)
   * Pivot tables (preservation + basic creation)
   * Conditional formatting rules
   * Data validation
   * All plug into formula/graph system (tables/pivots heavily referenced by formulas)

2. **Drawings & charts**:
   * **Drawings**: images, shapes, anchors, textboxes
   * **Charts**: bar/line/pie/scatter; data bindings to cells or tables; style knobs

**Dependencies**:
* **Tables/pivots**: Easier once formulas, named ranges, graph exist (structured references wire into dependency graph)
* **Charts**: Conceptually depend on tables/pivots & ranges, but can start with simple "chart over cell range"

**Parallelization**:
* **Stream A**: Tables + data validation + conditional formatting (share XML + rule logic)
* **Stream B**: Drawings (images, shapes)
* **Stream C**: Charts (start from simple bar/line anchored to ranges)

Coordination needed around OOXML relationship plumbing, but each sub-area has own module.

**Maps to existing phases**: P10 (drawings.md), P11 (charts.md), P12 (tables-and-pivots.md)

---

### Phase 6 – Security, Safety & Golden Tests

**Goal**: Make XL safe for untrusted inputs and "boringly reliable."

**Work Items**:

* **P13 security & safety**:
  * ZIP bomb detection, size/ratio limits
  * XXE prevention across all XML entry points
  * Formula injection guards (opt-in safe mode for untrusted data)
  * XLSM macro preservation (never execute, never drop)

* **Golden test corpus + compatibility matrix**:
  * Curated `.xlsx` corpus covering feature combinations
  * Stable XML diff tooling
  * Versioned across Excel versions (2007 → M365)

* **Stress & mutation tests**:
  * 1M+ rows, deeply nested formulas, complex graphs
  * Mutation testing to verify test thoroughness

**Why After Advanced Features?**:
* Golden files & compatibility tests most useful once feature set more complete
* Don't want to constantly rebuild corpus as core capabilities land
* Security can be parallelized earlier, but finishing here gives full surface to protect

**Parallelization**:
* **Security stream**: ZIP/XXE/limits/macros/injection
* **Testing stream**: Golden files, compatibility matrix, stress tests

**Maps to existing phases**: P13 (security.md, error-model-and-safety.md)

---

### Phase 7 – Architectural Evolution (Lazy Evaluation / Logical Plans)

**Goal**: Upgrade internal execution model without disturbing users.

**Work Items**:
* Implement `LazySheet` + logical plan DSL
* Rewrite high-level operations into plan
* Optimize plan (projection/predicate pushdown, dedup, etc.)
* Lower to streaming execution with fs2, leveraging established I/O & evaluator

**Dependencies**:
* **Beneficial to have**:
  * Stable domain model & AST (Phases 2-5)
  * Evaluator & graph system (Phase 4) - easier to reason about laziness + dependency ordering
  * Streaming infrastructure & perf harness (Phases 1 & 6) for validation

**Parallelization**:
* Essentially own "epic" touching everything
* Best treated as focused, time-boxed project after rest settles

**Status**: **Deferred post-1.0** (see `docs/archive/plan/deferred/lazy-evaluation.md`)

---

## 4. Where Key Features Sit

### Comments: **Phase 2** (✅ Already Complete)

* Already implemented: typed AST + simple APIs (get/set comment on A1)
* OOXML round-trip working
* Preserved surgically
* No dependencies, localized, user-facing
* **No work needed** - marked complete

### Graph Search via Relationships: **Phase 4** (Core Differentiator)

* First: part-level graph (workbook ⇄ sheets ⇄ tables ⇄ charts) from `Relationships` + named ranges + calcChain
* Then: cell-level dependency graph from formula AST
* Combined: answer questions like "Which charts depend on this input cell?" or "Which pivot tables fed by formulas reference external links?"

**This is the killer feature** that distinguishes XL from POI.

---

## 5. Suggested Parallel Development Streams

For 3-4 developers over time:

### 1. **Core & IO Stream**
* P6.5, streaming-improvements, benchmarks
* Later: lazy-eval + perf tuning
* **Modules**: xl-cats-effect, xl-benchmarks

### 2. **OOXML & AST Stream**
* Phase 2 mappings, tables/pivots, charts/drawings
* Comments/hyperlinks, print/docProps
* **Modules**: xl-ooxml

### 3. **Semantics & Graph Stream**
* Formula evaluator, dependency graph
* Security (formula injection)
* Graph query API
* **Modules**: xl-core, xl-evaluator (new), xl-graph (new)

### 4. **DX & Docs Stream**
* P6b, macros, easy-mode APIs
* Query API extras
* Cookbook/docs, examples
* **Modules**: xl-macros, docs

These streams align with existing module structure and minimize merge conflicts.

---

## 6. Mapping to Existing Granular Phases

This strategic framework maps to existing P0-P13 numbering in [roadmap.md](roadmap.md):

| Strategic Phase | Existing Phases | Status |
|----------------|----------------|--------|
| Phase 1 (Hardening) | P6.5, P6.7 | P6.7 ✅, P6.5 ⬜ |
| Phase 2 (OOXML AST) | P4 extensions | Partially ✅ (comments done) |
| Phase 3 (DX) | P6b, P9 | P9 ⬜, P6b ⬜ |
| Phase 4 (Formula/Graph) | P9+ (formula-system.md) | ⬜ |
| Phase 5 (Advanced) | P10, P11, P12 | ⬜ |
| Phase 6 (Security) | P13 | ⬜ |
| Phase 7 (Lazy Eval) | Deferred | ⏸ |

See [roadmap.md](roadmap.md) for detailed completion status, test counts, and commit history.

---

## 7. Implementation Resources

For concrete code scaffolds and implementation patterns, see:
* **[docs/reference/implementation-scaffolds.md](../reference/implementation-scaffolds.md)** - Scala 3 code examples for all major features
* **Individual plan files** - Detailed designs for specific features (formula-system.md, drawings.md, etc.)

---

## 8. Success Criteria

**Phase 1 Complete**: When all tests pass with <5% regression variance, JMH benchmarks in CI
**Phase 2 Complete**: When all common OOXML parts have typed ASTs with round-trip tests
**Phase 3 Complete**: When case class → row mapping is automatic and path macro works
**Phase 4 Complete**: When dependency graph answers reachability queries in <10ms for 10k cell workbooks
**Phase 5 Complete**: When we match or exceed POI feature set
**Phase 6 Complete**: When we pass fuzzing tests and handle all security.md scenarios
**Phase 7 Complete**: When LazySheet is default with <10% overhead vs eager

---

**Last Updated**: 2025-11-20
**Maintained By**: XL Core Team
**Next Review**: After each major phase completion
