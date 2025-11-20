# Archived Documentation

This file tracks documentation that has been moved to `docs/archive/` with reasons and timestamps.

## Archive Structure

```
docs/archive/
├── plan/
│   ├── p0-bootstrap/          # Phase 0 plans (complete)
│   ├── p1-addressing/         # Phase 1 plans (complete)
│   ├── p2-patches/            # Phase 2 plans (complete)
│   ├── p3-styles/             # Phase 3 plans (complete)
│   ├── p4-ooxml/              # Phase 4 plans (complete)
│   ├── p45-ooxml-quality.md   # Phase 4.5 (complete)
│   ├── p5-streaming/          # Phase 5 plans (complete)
│   ├── p6-codecs/             # Phase 6 plans (complete)
│   ├── p31-refactor/          # Phase 31 refactoring (complete)
│   ├── p68-surgical-modification/  # Phase 6.8 surgical modification (complete)
│   ├── string-interpolation/  # P7-P8 string interpolation (complete)
│   ├── unified-put-api.md     # Unified put API consolidation (complete)
│   ├── completed-post-p8/     # Other completed plans after P8
│   └── deferred/              # Deferred/rejected plans
└── (future: reference/, design/ subdirectories as needed)
```

## Recently Archived (2025-11-20)

### Session 1: Plan Cleanup

| Original Path | New Path | Reason | Date |
|---------------|----------|--------|------|
| `docs/plan/type-class-put.md` | `docs/archive/plan/completed-post-p8/type-class-put.md` | ✅ Completed in PR #20 (type class consolidation for Easy Mode put()) | 2025-11-20 |
| `docs/plan/numfmt-preservation.md` | `docs/archive/plan/p68-surgical-modification/numfmt-preservation.md` | ✅ Completed in P6.8 (surgical modification); duplicate of numfmt-id-preservation.md | 2025-11-20 |
| `docs/plan/numfmt-id-preservation.md` | `docs/archive/plan/p68-surgical-modification/numfmt-id-preservation.md` | ✅ Completed in P6.8 (CellStyle.numFmtId field for byte-perfect preservation) | 2025-11-20 |
| `docs/plan/lazy-evaluation.md` | `docs/archive/plan/deferred/lazy-evaluation.md` | ⏸ Deferred indefinitely (Spark-style optimizer deemed overkill; streaming-improvements.md prioritized instead) | 2025-11-20 |

### Session 2: Design Docs Archive

| Original Path | New Path | Reason | Date |
|---------------|----------|--------|------|
| `docs/design/easy-mode-api.md` | `docs/archive/design/easy-mode-api.md` | ✅ Completed (PR #20); design proposal fully implemented | 2025-11-20 |
| `docs/design/unified-ref-literal.md` | `docs/archive/design/unified-ref-literal.md` | ✅ Implemented (ref"..." literal standard; cell"..." deprecated) | 2025-11-20 |

### Session 2: File Consolidations

| Action | Files | Result | Reason | Date |
|--------|-------|--------|--------|------|
| Merge | `drawings.md`, `charts.md`, `tables-and-pivots.md`, `benchmarks.md` → `advanced-features.md` | -4 files | Consolidated stub plans (<20 lines each) into comprehensive P10-P12 plan | 2025-11-20 |
| Merge | `security.md` → `error-model-and-safety.md` | -1 file | Eliminated duplicate P13 security coverage | 2025-11-20 |
| Merge | `style-guide.md` → `CONTRIBUTING.md` | -1 file | Style guide naturally part of contribution guidelines | 2025-11-20 |
| Merge | `glossary.md` + `FAQ.md` → `FAQ-AND-GLOSSARY.md` | -2 files (+1 new) | Natural pairing of questions and terminology | 2025-11-20 |
| Merge | `ooxml-cheatsheet.md` → `ooxml-research.md` | -1 file | Cheatsheet as "Quick Reference" section | 2025-11-20 |
| Merge | `executive-summary.md` → Deleted | -1 file | Content redundant with README.md and project overview | 2025-11-20 |
| Merge | `quick-wins.md` → `future-improvements.md` | -1 file | Both track low-priority enhancements | 2025-11-20 |
| Merge | `plan/README.md` → `roadmap.md` | -1 file | Plan index integrated into roadmap "How to Use" section | 2025-11-20 |

**Total Consolidation**: 41 files → 28 files (**-13 files**, 32% reduction)

## Previously Archived (Historical)

### Phase 0-6 + P31 Completion

All implementation plans for phases P0 through P6, P31, and P4.5 were archived upon completion:

- **P0** (Bootstrap & CI): Build system, modules, formatting, GitHub Actions
- **P1** (Addressing & Literals): Opaque types, ARef packing, macros (`ref"A1"`, `ref"A1:B10"`)
- **P2** (Core & Patches): Immutable domain model, Patch Monoid, XLError ADT
- **P3** (Styles): CellStyle system, StylePatch, StyleRegistry, unit conversions
- **P4** (OOXML MVP): Full XLSX read/write, SST, Styles.xml, ZIP assembly
- **P4.5** (OOXML Quality): Spec compliance, round-trip fidelity, whitespace preservation
- **P5** (Streaming): Excel[F] algebra, true streaming read/write with fs2-data-xml
- **P6** (Codecs): CellCodec primitives (9 types), batch operations, auto-formatting
- **P31** (Refactoring): Optics, RichText, HTML export, ergonomic enhancements

### String Interpolation (P7-P8)

- **P7**: Runtime string interpolation for all macros (`ref"$var"`, `money"$var"`, etc.)
- **P8**: Compile-time optimization when all interpolated values are literals

### API Consolidation

- **Unified Put API**: Consolidation of batch put operations with auto-inferred formatting
- **Type Class Put**: CellWriter type class for extensible, type-safe put() methods

### Surgical Modification (P6.8)

- **P6.8**: Surgical modification with unknown part preservation, hybrid write strategy
- **NumFmt Preservation**: CellStyle.numFmtId for byte-perfect format ID preservation

## Archival Policy

Documents are archived when:

1. **Completed**: Implementation finished, tested, and merged (moved to `completed-*` or phase-specific folder)
2. **Deferred**: Explicitly decided not to implement (moved to `deferred/`)
3. **Superseded**: Replaced by newer approach or integrated into another plan
4. **Duplicate**: Redundant with another document

Before archiving:
- Add completion banner at top: `> **Status**: ✅ Completed / ⏸ Deferred`
- Include reason, commit/PR references, and archive date
- Update `docs/plan/roadmap.md` to reflect new status
- Add entry to this ARCHIVE_LIST.md

## Finding Archived Documentation

Use this tracking file or `git log` to find moved documentation:

```bash
# Find when a file was moved
git log --follow --all -- docs/plan/type-class-put.md

# Search archived plans by keyword
grep -r "keyword" docs/archive/plan/

# List all archived plans
find docs/archive/plan/ -name "*.md" -type f
```

---

**Maintained by**: Claude Code cleanup sessions
**Last Updated**: 2025-11-20
