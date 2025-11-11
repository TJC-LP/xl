# Active Plans

This directory contains **active future work** only. Completed phases are archived in `docs/archive/plan/`.

## Current Status

- **Phases P0-P6 + P31**: ✅ Complete (85%, 263/263 tests passing)
- **Production Ready**: Core XLSX read/write, streaming, codecs, optics, RichText
- **Future Work**: Documented below

## Active Plans (Unimplemented Features)

### Core Future Work

1. **[future-improvements.md](future-improvements.md)** - P6.5 Performance & Quality Polish
   - Status: ⬜ Not started
   - Source: PR #4 review feedback (medium-priority enhancements)
   - Scope: O(n²) → O(1) optimization, whitespace utility, error path tests, integration tests
   - Priority: Medium (8-10 hours)

2. **[formula-system.md](formula-system.md)** - P7+ Formula Evaluator
   - Status: ⬜ Not started
   - Scope: AST evaluation, function library, circular reference detection
   - Note: Formula parsing/serialization already complete

3. **[error-model-and-safety.md](error-model-and-safety.md)** - P11 Security Hardening
   - Status: Error model ✅ complete, security ⬜ not started
   - Scope: ZIP bomb detection, formula injection guards, XXE prevention
   - Priority: High (required for production use)

4. **[security.md](security.md)** - P11 Additional Security Features
   - Status: ⬜ Not started
   - Scope: File size limits, macro handling, sanitization
   - Note: Consider merging with error-model-and-safety.md

### Architectural Evolution

5. **[lazy-evaluation.md](lazy-evaluation.md)** - Spark-Style Lazy Evaluation
   - Status: ⬜ Not started (design complete)
   - Scope: Logical plan DSL, query optimizer (4 passes), streaming execution with fs2
   - Breaking Change: Sheet → LazySheet (eager → lazy by default)
   - Benefits: 35% faster writes, O(n) → O(1) memory, 20-40% operation reduction
   - Timeline: 4-5 weeks (6 phases)
   - Note: Major architectural rewrite, deferred until post-1.0

### Advanced Features

6. **[drawings.md](drawings.md)** - P8 Images & Shapes
   - Status: ⬜ Not started
   - Scope: PNG/JPEG embedding, anchors, text boxes

7. **[charts.md](charts.md)** - P9 Chart Generation
   - Status: ⬜ Not started
   - Scope: Bar, line, pie, scatter charts with data binding

8. **[tables-and-pivots.md](tables-and-pivots.md)** - P10 Structured Data
   - Status: ⬜ Not started
   - Scope: Excel tables, conditional formatting, data validation

### Infrastructure

9. **[benchmarks.md](benchmarks.md)** - Performance Testing
   - Status: ⬜ Not started
   - Scope: JMH benchmarks, POI comparisons, regression tests

10. **[roadmap.md](roadmap.md)** - Master Status Tracker
    - Status: ✅ Maintained (living document)
    - Scope: Overall progress, phase definitions, completion criteria

## Related Documentation

- **Design Docs**: `docs/design/` - Architectural decisions (purity charter, domain model, ADRs)
- **Reference**: `docs/reference/` - Examples, glossary, OOXML research, testing guide
- **Archived Plans**: `docs/archive/plan/` - Completed P0-P6, P31 implementation plans
- **Status**: `docs/STATUS.md` - Detailed current state (263 tests, performance, limitations)

## Quick Links

- **[CLAUDE.md](../../CLAUDE.md)** - AI assistant context for working with the codebase
- **[README.md](../../README.md)** - User-facing documentation
- **[CONTRIBUTING.md](../CONTRIBUTING.md)** - Contribution guidelines
- **[FAQ.md](../FAQ.md)** - Frequently asked questions

## Bugfix & Integration Guidance

### High-Priority Quality Issues (Active)

_No active quality issues! All critical defects resolved._

### Recently Completed Quality Improvements

9. ✅ **[ooxml-quality.md](ooxml-quality.md)** - P4.5 OOXML Spec Compliance
   - Status: ✅ Complete (commits b22832e, 4dd98f5)
   - Completed: 2025-11-10
   - Scope: Fixed all 6 critical issues (gray125, SharedStrings count, whitespace, alignment, Scala version, xml:space)
   - Result: Zero spec violations, zero data loss, full round-trip fidelity
   - Tests: +22 new tests across 4 test classes

### Recently Fixed Defects

- ✅ **Styles lost on import** - Fixed in commit 7293ccc
  - *Solution*: Extended XlsxReader to parse `xl/styles.xml` and populate StyleRegistry per sheet

- ✅ **Worksheet lookup ignores `r:id` relationships** - Fixed in commit efd0204
  - *Solution*: Added relationship-based worksheet resolution using `xl/_rels/workbook.xml.rels`

- ✅ **Runtime cell parser rejects lowercase columns** - Fixed in commit 7293ccc
  - *Solution*: Added `toUpperCase(Locale.ROOT)` normalization to ARef.parse and Column.fromLetter

- ✅ **`Lens.modify` does not update structures** - Fixed in commit 7293ccc
  - *Solution*: Changed return type to S and delegated to `set(f(get(s)), s)`

### Implementation Order Best Practices

1. **Data-loss fixes first** – prioritize IO defects (styles + relationships) before API ergonomics, because they affect workbook integrity and any downstream plan milestone relying on round-tripping data.
2. **Shared infrastructure next** – when two bugs share plumbing (style handling touches both reader and writer), design those changes together to avoid churn in later phases like P8–P10.
3. **API consistency pass** – once IO stability is restored, address library-surface issues (`Column.fromLetter`, `Lens.modify`) so future feature work (e.g., formula system) assumes correct primitives.
4. **Document & test as you go** – each fix should add regression tests under the owning module (`xl-ooxml/test/...`, `xl-core/test/...`) and update this plan README (status links) so contributors know where the work landed.
5. **Align with roadmap phases** – tag each fix with the relevant plan phase (e.g., P31 for IO, P0 for core API) to keep `docs/plan/roadmap.md` authoritative and help future triage slot work into sprint buckets.

This guidance should be kept in sync with `docs/plan/roadmap.md` whenever priorities shift or additional blocking bugs are discovered.
