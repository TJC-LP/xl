# Active Plans

This directory contains **active future work** only. Completed phases are archived in `docs/archive/plan/`.

## Current Status

- **Phases P0-P6 + P31**: ✅ Complete (85%, 263/263 tests passing)
- **Production Ready**: Core XLSX read/write, streaming, codecs, optics, RichText
- **Future Work**: Documented below

## Active Plans (Unimplemented Features)

### Core Future Work

1. **[formula-system.md](formula-system.md)** - P7+ Formula Evaluator
   - Status: ⬜ Not started
   - Scope: AST evaluation, function library, circular reference detection
   - Note: Formula parsing/serialization already complete

2. **[error-model-and-safety.md](error-model-and-safety.md)** - P11 Security Hardening
   - Status: Error model ✅ complete, security ⬜ not started
   - Scope: ZIP bomb detection, formula injection guards, XXE prevention
   - Priority: High (required for production use)

3. **[security.md](security.md)** - P11 Additional Security Features
   - Status: ⬜ Not started
   - Scope: File size limits, macro handling, sanitization
   - Note: Consider merging with error-model-and-safety.md

### Advanced Features

4. **[drawings.md](drawings.md)** - P8 Images & Shapes
   - Status: ⬜ Not started
   - Scope: PNG/JPEG embedding, anchors, text boxes

5. **[charts.md](charts.md)** - P9 Chart Generation
   - Status: ⬜ Not started
   - Scope: Bar, line, pie, scatter charts with data binding

6. **[tables-and-pivots.md](tables-and-pivots.md)** - P10 Structured Data
   - Status: ⬜ Not started
   - Scope: Excel tables, conditional formatting, data validation

### Infrastructure

7. **[benchmarks.md](benchmarks.md)** - Performance Testing
   - Status: ⬜ Not started
   - Scope: JMH benchmarks, POI comparisons, regression tests

8. **[roadmap.md](roadmap.md)** - Master Status Tracker
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

### High-Priority Defects (Active)

- **Styles lost on import (`XlsxReader.convertToDomainSheet`)**
  - *Symptom*: `Cell.styleId` values refer to workbook-level indices that are never registered in the returned `Sheet`, so any subsequent write or inspection drops formatting.
  - *Fix plan*: extend the reader pipeline to parse `xl/styles.xml`, hydrate a `StyleRegistry` per sheet, and remap workbook `cellXf` indices into sheet-local `StyleId`s. Update roadmap item P31 (IO parity) once the registry is preserved through the round trip. Tests should cover reading a styled workbook and asserting `sheet.getCellStyle` matches the source fonts/fills.

- **Worksheet lookup ignores `r:id` relationships**
  - *Symptom*: `sheetId` is treated as `sheetN.xml`, which fails when Excel leaves gaps or renames targets.
  - *Fix plan*: load `xl/_rels/workbook.xml.rels`, build a map from `relationshipId` → part path, and have `parseSheets` read using the resolved relationship target. Cover with a test that reorders sheets so that `sheetId` ≠ file suffix.

- **Runtime cell parser rejects lowercase columns**
  - *Symptom*: APIs such as `"a1".asCell` or `ARef.parse("b2")` return `Invalid column letter` even though Excel treats references case-insensitively and the compile-time macro accepts lowercase.
  - *Fix plan*: normalize input to uppercase (Locale.ROOT) in `Column.fromLetter` before validation; add tests proving mixed-case references round-trip.

- **`Lens.modify` does not update structures**
  - *Symptom*: Method returns the transformed field instead of writing it back via `set`, diverging from the documented behavior and standard lens laws.
  - *Fix plan*: change `modify` to return `S` and delegate to `set(f(get(s)), s)` (or rename/remove if a pure value transformer is desired). Update `Lens` tests in `OpticsSpec` to exercise `modify` and verify lawfulness.

### Implementation Order Best Practices

1. **Data-loss fixes first** – prioritize IO defects (styles + relationships) before API ergonomics, because they affect workbook integrity and any downstream plan milestone relying on round-tripping data.
2. **Shared infrastructure next** – when two bugs share plumbing (style handling touches both reader and writer), design those changes together to avoid churn in later phases like P8–P10.
3. **API consistency pass** – once IO stability is restored, address library-surface issues (`Column.fromLetter`, `Lens.modify`) so future feature work (e.g., formula system) assumes correct primitives.
4. **Document & test as you go** – each fix should add regression tests under the owning module (`xl-ooxml/test/...`, `xl-core/test/...`) and update this plan README (status links) so contributors know where the work landed.
5. **Align with roadmap phases** – tag each fix with the relevant plan phase (e.g., P31 for IO, P0 for core API) to keep `docs/plan/roadmap.md` authoritative and help future triage slot work into sprint buckets.

This guidance should be kept in sync with `docs/plan/roadmap.md` whenever priorities shift or additional blocking bugs are discovered.
