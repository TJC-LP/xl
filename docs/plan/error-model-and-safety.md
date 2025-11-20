# Error Model & Security Hardening

**Status**: Error model ✅ complete, security features ⬜ not started
**Priority**: High (for production use)
**Estimated Effort**: 2-3 weeks
**Last Updated**: 2025-11-20

---

## Metadata

| Field | Value |
|-------|-------|
| **Owner Modules** | `xl-ooxml` (security checks), `xl-core` (error ADT — already complete) |
| **Touches Files** | `XlsxReader.scala`, `XlsxWriter.scala`, `SecurityValidator.scala` (new) |
| **Dependencies** | P0-P8 complete (P6.5-03 XXE fix already done) |
| **Enables** | Production-ready 1.0 release |
| **Parallelizable With** | WI-07 (formula), WI-10 (tables), WI-15 (benchmarks) — different concerns |
| **Merge Risk** | Medium (modifies reader/writer, but additive validation only) |

---

## Work Items

| ID | Description | Type | Files | Status | PR |
|----|-------------|------|-------|--------|----|
| `WI-30` | ZIP Bomb Detection | Security | `SecurityValidator.scala`, `XlsxReader.scala` | ⏳ Not Started | - |
| `WI-31` | Formula Injection Guards | Security | `XlsxWriter.scala`, `CellValue.scala` | ⏳ Not Started | - |
| `WI-32` | File Size Limits | Security | `XlsxReader.scala`, config | ⏳ Not Started | - |
| `WI-33` | XLSM Macro Handling | Security | `XlsxReader.scala`, `ContentTypes.scala` | ⏳ Not Started | - |
| `WI-34` | XXE Hardening (explicit) | Security | `XlsxReader.scala` | ✅ Done (P6.5-03) | (79b3269) |

---

## Dependencies

### Prerequisites
- ✅ P0-P8: Foundation complete
- ✅ P6.5-03: XXE vulnerability fixed (basic protection in place)

### Enables
- Production 1.0 release (security requirements met)
- Enterprise adoption (security compliance)

### File Conflicts
- **Medium risk**: WI-30/WI-32 modify `XlsxReader.scala` (coordinate if reader changes active)
- **Low risk**: WI-31 modifies writer (isolated)
- **Low risk**: WI-33 is additive (new XLSM handling)

### Safe Parallelization
- ✅ WI-07/WI-08 (Formula) — different module
- ✅ WI-10/WI-11 (Tables/Charts) — different concerns
- ✅ WI-15 (Benchmarks) — infrastructure work

---

## Worktree Strategy

**Branch naming**: `security` or `WI-30-zip-bomb-detection`

**Merge order**:
1. WI-30 (ZIP Bomb) — most critical, reader-side
2. WI-32 (File Limits) — pairs with WI-30
3. WI-33 (XLSM) — independent
4. WI-31 (Formula Injection) — writer-side, can parallel with reader work

**Conflict resolution**:
- If reader changes active, coordinate WI-30/WI-32 merge order
- WI-31 and WI-33 are independent

---

## Execution Algorithm

### WI-30: ZIP Bomb Detection
```
1. Create worktree: `gtr create WI-30-zip-bomb`
2. Create SecurityValidator.scala in xl-ooxml:
   - validateCompressionRatio(entry: ZipEntry)
   - validateEntryCount(zipFile: ZipFile)
   - validateUncompressedSize(zipFile: ZipFile)
3. Integrate into XlsxReader.read():
   - Check before extracting ZIP entries
   - Return XLError.SecurityViolation on detection
4. Add tests:
   - Generate pathological ZIP files
   - Verify rejection with appropriate errors
5. Run tests: `./mill xl-ooxml.test`
6. Create PR: "feat(security): add ZIP bomb detection"
7. Update roadmap: WI-30 → ✅ Complete
```

### WI-31: Formula Injection Guards
```
1. Create worktree: `gtr create WI-31-formula-injection`
2. Add escaping logic to CellValue.Text writing:
   - Detect dangerous prefixes (= + - @)
   - Prefix with single quote when detected
   - Make configurable via WriterConfig
3. Add tests:
   - Verify "=CMD" escapes to "'=CMD"
   - Verify legitimate formulas unchanged
4. Run tests: `./mill xl-ooxml.test`
5. Create PR: "feat(security): add formula injection guards"
6. Update roadmap: WI-31 → ✅ Complete
```

### WI-32: File Size Limits
```
1. Create worktree: `gtr create WI-32-file-limits`
2. Add ReaderConfig with limits:
   - maxFileSize (default 100MB)
   - maxCellCount (default 10M)
   - maxStringLength (default 32KB)
3. Integrate checks into XlsxReader
4. Add tests for limit enforcement
5. Run tests: `./mill xl-ooxml.test`
6. Create PR: "feat(security): add configurable file size limits"
7. Update roadmap: WI-32 → ✅ Complete
```

### WI-33: XLSM Macro Handling
```
1. Create worktree: `gtr create WI-33-xlsm`
2. Update ContentTypes to recognize XLSM
3. Add vbaProject.bin preservation (opaque bytes)
4. Add stripMacros() API to remove vbaProject
5. Add tests for XLSM round-trip (no execution)
6. Run tests: `./mill xl-ooxml.test`
7. Create PR: "feat(security): add XLSM macro preservation (no execution)"
8. Update roadmap: WI-33 → ✅ Complete
```

---

## Design

### Completed: Error Model ✅

Error handling is centralized in `com.tjclp.xl.error.XLError` (see `xl-core/src/com/tjclp/xl/error/XLError.scala`). Key cases:

- **Addressing/validation**: `InvalidCellRef`, `InvalidRange`, `InvalidReference`, `InvalidSheetName`, `InvalidColumn`, `InvalidRow`, `OutOfBounds`
- **Workbook structure**: `SheetNotFound`, `DuplicateSheet`, `InvalidWorkbook`, `ValueCountMismatch`
- **Typing/formatting**: `TypeMismatch`, `FormulaError`, `StyleError`, `NumberFormatError`, `MoneyFormatError`, `PercentFormatError`, `DateFormatError`, `AccountingFormatError`, `ColorError`, `UnsupportedType`
- **IO/parse surface**: `IOError`, `ParseError`, `Other`

**Properties**:
- ✅ Public APIs are pure: return `XLResult[A] = Either[XLError, A]`
- ✅ Exceptions only via explicit `.unsafe` helpers
- ✅ Platform exceptions wrapped in IO/parse layers
- ✅ All validation is total (never throws)

### Not Started: Security Features (P13) ⬜

#### Threat Model
**Untrusted inputs**: Adversarial `.xlsx` files, malformed XML, formula injection vectors

**Attack vectors**:
- ZIP bombs (compression ratio attacks)
- XML entity expansion
- Formula injection (CSV injection via Excel)
- Path traversal in ZIP entries
- Resource exhaustion (memory, disk, CPU)

#### ZIP Bomb Detection
- **Entry count limits**: Prevent archives with excessive file counts
- **Uncompressed size limits**: Detect compression ratio attacks (cap size/ratio)
- **Nested archive detection**: Prevent recursive bombs
- **Path traversal prevention**: Reject `..` or absolute paths in ZIP entries

#### Formula Injection Guards
- **Untrusted text escaping**: Prefix with `'` when starting with `= + - @`
- **Configurable policy**: Strict (always escape) vs permissive (warn only)
- **CSV export safety**: Ensure formula injection protection on export

#### XLSM Macro Handling
- **Opaque preservation**: Read/write XLSM parts without execution
- **Explicit stripping**: API to remove vbaProject.bin
- **Never execute**: No macro evaluation, ever

#### File Size Limits
- **Max file size**: Configurable limit (default 100MB)
- **Max cell count**: Prevent pathological sheets (default 10M cells)
- **Max string length**: Prevent memory exhaustion (default 32KB per cell)

#### XML External Entity (XXE) Prevention
- ✅ **Currently safe**: P6.5-03 disabled DTDs and external entities
- **Explicit hardening**: XML parser limits (entity expansion, depth, attribute count)

### Security Guidance
- **Never evaluate macros**: Do not embed active content (XLSM vbaProject.bin never executed)
- **Prefer typed formulas**: Use formula AST over string manipulation where possible
- **Formula text sanitizer**: Opt-in allowlist for untrusted text writes

---

## Definition of Done

### Functional
- [ ] ZIP bomb detection prevents malicious archives
- [ ] Formula injection guards protect against CSV injection
- [ ] File size limits prevent resource exhaustion
- [ ] XLSM macros preserved but never executed
- [ ] XXE hardening explicitly configured

### Code Quality
- [ ] Zero WartRemover errors
- [ ] Security tests cover all attack vectors
- [ ] 30+ tests added (10 per major feature)
- [ ] Scalafmt applied

### Documentation
- [ ] CLAUDE.md updated with security guidance
- [ ] STATUS.md updated with security capabilities
- [ ] Examples in docs/reference/examples.md
- [ ] Security audit documentation

### Integration
- [ ] Roadmap.md: WI-30 through WI-34 marked Complete
- [ ] No regressions (all existing tests pass)
- [ ] Performance impact < 5% (security checks are lightweight)

---

## Module Ownership

**Primary**: `xl-ooxml` (SecurityValidator.scala, XlsxReader.scala, XlsxWriter.scala)

**Secondary**: `xl-core` (config types for limits/policies)

**Test Files**: `xl-ooxml/test` (SecuritySpec extensions, ZipBombSpec, FormulaInjectionSpec)

---

## Merge Risk Assessment

**Risk Level**: Medium

**Rationale**:
- Modifies reader/writer (shared across many features)
- Changes are additive (validation only, no refactoring)
- Performance impact minimal (<5%)

**Mitigation**:
- Small PRs (one work item per PR)
- Comprehensive tests before merge
- Feature flags if needed for gradual rollout

---

## Related Documentation

- **Roadmap**: `docs/plan/roadmap.md` (WI-30 through WI-34)
- **Design**: `docs/design/purity-charter.md` (error handling philosophy)
- **Strategic**: `docs/plan/strategic-implementation-plan.md` (Phase 6: Security & Observability)
- **Tests**: `xl-ooxml/test/src/com/tjclp/xl/ooxml/SecuritySpec.scala`

---

## Notes

- **XXE protection** already in place (P6.5-03) — this plan adds explicit configuration and additional hardening
- **Priority**: High for production 1.0 release (security table stakes)
- **Effort**: 2-3 weeks total (can parallelize work items)
- **Testing**: Security features require adversarial test cases (malicious ZIPs, formula injection attempts)
