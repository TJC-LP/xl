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
