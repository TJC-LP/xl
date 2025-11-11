# Style Review - Status: Documented ✓

**Review Date**: Pre-implementation
**Documentation Date**: 2025-11-10
**Status**: All issues documented in `docs/plan/ooxml-quality.md` for future implementation

> **Note**: This review was conducted before implementation. All identified issues have been triaged and documented in the active plan docs. See [ooxml-quality.md](../plan/ooxml-quality.md) for the implementation plan.

---

## Summary

Overall, the codebase is cleanly structured, consistent, and well-documented. It follows your stated principles: pure core, deterministic XML, law-governed algebras, and zero-overhead opaque types. The OOXML layer has clear separation of concerns and good deterministic helpers (XmlUtil). Tests are meaningful and comprehensive, especially for addressing, styles, patches, codecs, and streaming.

For AI/Human readability, you’ve already done a lot right: explicit types on public APIs, descriptive ADTs, predictable naming (from*, to*, parse), and rich Scaladoc. There are a few correctness gaps in the OOXML layer that should be fixed, and several style tweaks that would further improve readability and AI-friendliness.

Critical issues to address

1) OOXML styles default fills are incorrect
- Problem: Excel requires two default fills at indices 0 and 1: patternType="none" and patternType="gray125". Current code uses:
  val defaultFills = Vector(Fill.None, Fill.Solid(Color.Rgb(0x00000000)))
  This can produce styles.xml that is non-conformant and potentially triggers odd behavior in Excel.
- Fix: Use Gray125 for the second default fill.
  - Add Fill.Pattern(..., PatternType.Gray125) (with suitable fg/bg). Or better, represent the “default” two fills explicitly independent of your Fill ADT to guarantee spec compliance.
  - Update serializer and tests to assert that the second fill is gray125.

2) SharedStrings count attribute is wrong
- Problem: In SharedStrings.toXml, both count and uniqueCount are set to strings.size. The OOXML spec expects:
  - count = total number of string instances across cells
  - uniqueCount = number of unique strings
- Fix: Track totalCount alongside unique strings (change SharedStrings to carry totalCount: Int, unique: Vector[String], indexMap). Modify fromWorkbook to compute totalCount from all text cells and uniqueCount from the deduped set. Update toXml to emit both correctly (or omit count if unavailable).

3) Inline string whitespace preservation gaps
- Problem: For inline strings (<is><t>text</t></is>), leading/trailing/multiple spaces must be preserved via xml:space="preserve".
  - SharedStrings handles this, but:
    - OoxmlCell for CellValue.Text with inline strings doesn’t add xml:space="preserve".
    - StreamingXmlWriter.cellToEvents for CellValue.Text writes <t>text</t> without xml:space when needed.
- Fix: Add xml:space="preserve" for inline <t> when:
  - startsWith(" "), endsWith(" "), or contains("  ").
  - In Scala XML, prefer PrefixedAttribute("xml", "space", "preserve", ...) or emit the xml namespace-always-available prefix.
  - In fs2-data-xml, add Attr(QName(Some("xml"), "space"), List(XmlString("preserve", false))).

4) Alignment is parsed but not serialized into styles.xml
- Problem: OoxmlStyles writes cellXfs but never emits <alignment .../>, so alignment in CellStyle is lost when writing (tests currently don’t check this).
- Fix: For styles where alignment != default, include a child <alignment> under <xf> with horizontal, vertical, wrapText, indent, plus ensure excel recognizes it (you may need applyAlignment="1").

5) Scala version mismatch in project files
- build.mill uses Scala 3.7.3, project.scala uses 3.7.4, generated .scala-build shows 3.7.4. This is confusing for humans and tooling.
- Fix: Pick one version everywhere (recommended 3.7.3 if that’s what CI uses) and align README/CLAUDE docs.

6) Determinism/namespace nit: xml:space attribute construction
- You sometimes use UnprefixedAttribute("xml:space", "preserve", ...) which may still serialize to xml:space, but the idiomatic and unambiguous approach is PrefixedAttribute("xml", "space", "preserve", ...). Prefer the latter for clarity.

Minor improvements

AI/Human readability and consistency
- Make all public methods include explicit result types consistently (most are already typed; ensure consistency in all modules).
- Avoid wildcard imports in core/OOXML (e.g., import XmlUtil.*) where clarity matters; explicit imports help static analyzers and LLMs reason precisely.
- Centralize Excel serial number conversion:
  - You already have CellValue.dateTimeToExcelSerial/excelSerialToDateTime; ensure all writers (OoxmlWorksheet and StreamingXmlWriter) delegate to these (they do, but keep it strictly single-source-of-truth).
- Ensure identical behavior across streaming vs non-streaming writers:
  - Inline strings whitespace handling (see critical item 3).
  - Style/formatting parity: streaming doesn’t emit style references (by design) but document this difference prominently in Scaladoc.
- BigDecimal rendering:
  - Use n.bigDecimal.toPlainString for XML to avoid scientific notation surprises from Double → BigDecimal → toString paths.
- XlsxWriter configuration:
  - Consider config flag to control compression method (STORED vs DEFLATED). For production, default to DEFLATED; use STORED only when explicitly requested for debugging byte stability.

OOXML reader/writer edge cases
- OoxmlWorksheet parsing: type "s" with missing/invalid SST index currently maps to CellValue.Empty or Error. Ensure behaviors are documented, and consider reducing silent fallbacks in favor of explicit errors in reader (or return Empty only when spec allows).
- Relationships ordering and indexing: deterministic ordering is good; tests should assert workbook.xml’s r:id to rel mapping is consistent for non-sequential sheet indices (you partly cover this in streaming tests; consider adding a dedicated OOXML-level test).

Documentation and “AI-first” readability enhancements
- Strengthen Scaladoc contracts for all public functions:
  - Pre- and post-conditions, determinism guarantees, invariants, and error semantics as bullet lists.
  - “Notable edge cases” section per method (e.g., ARef.parse: empty strings, lower-case input).
- Add “AI contracts” blocks:
  - Short structured tags LLMs can latch on to, for example:
    - REQUIRES:
    - ENSURES:
    - DETERMINISTIC:
    - ERROR CASES:
  - Keep them consistent so prompts and agents can reliably parse expectations.
- Provide a “Public API index” document mapping each user entry point and its guarantees, with links. This vastly improves AI navigability.
- Prefer stable names for internal helpers used across modules; avoid overloading and optional parameters in core protocols.

Test coverage improvements
- Add tests for:
  - styles.xml having gray125 as default fill index 1.
  - Inline string whitespace preservation for both OOXML writer and streaming writer.
  - Alignment survives round-trip (verify <alignment> is emitted and read back).
  - SharedStrings count vs uniqueCount correctness.
  - Non-sequential sheet indices in OOXML (workbook.xml + rels mapping).

What’s done particularly well

- Strong type design with opaque types and enums.
- Deterministic XML helpers and sorted attributes.
- Clear separation of core vs OOXML vs streaming IO.
- Comprehensive property-based tests for the algebraic parts (addressing, patches, styles, codecs).
- Streaming writer and reader are thoughtfully designed and well documented.
- Style registry and deduplication strategy are clean and idiomatic.

Actionable next steps

1) Fix styles default fills:
- Replace the second default fill with gray125 and add a test to lock this in.

2) Correct SharedStrings:
- Change SharedStrings to track totalCount vs uniqueCount; update fromWorkbook and toXml accordingly.

3) Preserve whitespace for inline strings:
- Add xml:space="preserve" where needed in both OoxmlCell.toXml and StreamingXmlWriter.cellToEvents.

4) Serialize alignment:
- Emit <alignment> with appropriate attributes for non-default Align, and consider applyAlignment="1" if needed.

5) Unify Scala version across build files and docs.

6) Add “AI contracts” to Scaladoc and a Public API Index doc:
- Short REQUIRES/ENSURES/ERRORS sections for each public API.
- A top-level map of the public surface including intended determinism and side-effectlessness.

7) Consider writer config flags:
- compression: DEFLATED | STORED
- prettyPrint already exists; keep defaults documented for deterministic output.

8) Tighten XML namespace usage:
- Use PrefixedAttribute for xml:space consistently.

By addressing the few correctness gaps and adding structured “AI contracts” to your docs and Scaladoc, you’ll further improve both human and AI readability, make behavior easier to reason about, and reduce drift between modules.
