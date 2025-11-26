# Row/Column Operations — Full POI Parity

**Status**: 🔵 Ready to Start
**Priority**: High (POI parity, common use case)
**Estimated Effort**: 1-1.5 days
**Last Updated**: 2025-11-25

---

## Metadata

| Field | Value |
|-------|-------|
| **Owner Modules** | `xl-core` (domain model, DSL), `xl-ooxml` (serialization) |
| **Touches Files** | `RowProperties.scala`, `ColumnProperties.scala`, `Worksheet.scala`, `syntax.scala` |
| **Dependencies** | P0-P8 complete (foundation) |
| **Enables** | Full row/column formatting in output (widths, heights, hidden, grouping) |
| **Parallelizable With** | WI-07/WI-08 (Formula), WI-11 (Charts) — different modules |
| **Merge Risk** | Low (additive changes, mostly new code) |

---

## Work Items

| ID | Description | Type | Files | Status | PR |
|----|-------------|------|-------|--------|----|
| `WI-19a` | Serialize column properties to `<cols>` XML | Feature | `Worksheet.scala` | ⏳ Not Started | - |
| `WI-19b` | Serialize row properties to `<row>` attributes | Feature | `Worksheet.scala` | ⏳ Not Started | - |
| `WI-19c` | Add grouping fields to domain model | Feature | `RowProperties.scala`, `ColumnProperties.scala` | ⏳ Not Started | - |
| `WI-19d` | Column string interpolator macro | Feature | `ColumnLiteral.scala` (new) | ⏳ Not Started | - |
| `WI-19e` | Row/Column builder DSL | Feature | `RowColumnDsl.scala` (new) | ⏳ Not Started | - |
| `WI-19f` | Round-trip tests | Tests | `RowColumnPropertiesSpec.scala` (new) | ⏳ Not Started | - |

---

## Dependencies

### Prerequisites (Complete)
- ✅ P0-P8: Foundation complete
- ✅ Domain model: `RowProperties`, `ColumnProperties` exist
- ✅ Patches: `SetRowProperties`, `SetColumnProperties` work
- ✅ Sheet API: `setRowProperties()`, `setColumnProperties()` exist

### Enables
- Full row/column formatting preserved in output files
- Grouping/outlining for collapsible sections
- Ergonomic DSL for row/column operations
- POI feature parity for formatting

### File Conflicts
- **Low risk**: `Worksheet.scala` (additive — new helper functions)
- **Low risk**: `RowProperties.scala`, `ColumnProperties.scala` (add fields)
- **None**: All DSL files are new

### Safe Parallelization
- ✅ WI-07/WI-08/WI-09 (Formula) — different module
- ✅ WI-11 (Charts) — different OOXML parts
- ✅ WI-18 (Merged Cells) — same file but different sections, sequence after
- ⚠️ WI-16 (Two-Phase Writer) — coordinate if active

---

## Worktree Strategy

**Branch naming**: `feat/row-column-operations` or `WI-19-row-col-ops`

**Merge order**:
1. WI-19c (Domain model) — foundation for others
2. WI-19a/WI-19b (Serialization) — can parallel after domain
3. WI-19d/WI-19e (DSL) — after serialization complete
4. WI-19f (Tests) — continuous throughout

**Conflict resolution**: Minimal — mostly new files and additive changes

---

## Execution Algorithm

### Phase 1: Domain Model (WI-19c) — 30 minutes
```
1. Create worktree: `gtr create feat/row-column-operations`
2. Edit xl-core/src/com/tjclp/xl/sheets/RowProperties.scala:
   - Add outlineLevel: Option[Int] = None
   - Add collapsed: Boolean = false
   - Add require() for outline level 0-7 validation
3. Edit xl-core/src/com/tjclp/xl/sheets/ColumnProperties.scala:
   - Same fields as RowProperties
4. Update xl-core/test/src/com/tjclp/xl/Generators.scala:
   - Add generators for new fields
5. Run tests: `./mill xl-core.test`
6. Commit: "feat(core): add outlineLevel and collapsed to row/column properties"
```

### Phase 2: Column Serialization (WI-19a) — 2 hours
```
1. Edit xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala:
2. Add helper function groupConsecutiveColumns():
   - Input: Map[Column, ColumnProperties]
   - Output: Seq[(minCol, maxCol, properties)]
   - Merge adjacent columns with identical properties
3. Add helper function buildColsElement():
   - Generate <cols><col min="1" max="1" width="15.5" .../></cols>
   - Attributes: min, max, width, customWidth, hidden, style, outlineLevel, collapsed
4. Update OoxmlWorksheet.fromDomainWithMetadata():
   - Call buildColsElement() with sheet.columnProperties
   - Insert <cols> before <sheetData> in XML
5. Add test: column width round-trip
6. Add test: hidden column serialization
7. Add test: consecutive columns merged into spans
8. Run tests: `./mill xl-ooxml.test`
9. Commit: "feat(ooxml): serialize column properties to <cols> element"
```

### Phase 3: Row Serialization (WI-19b) — 1.5 hours
```
1. Edit xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala:
2. In row creation logic, merge domain RowProperties:
   - Apply height from sheet.getRowProperties(row)
   - Apply hidden flag
   - Apply outlineLevel and collapsed
   - Set customHeight="1" when height specified
3. Handle empty rows with properties:
   - Generate <row> elements for rows in rowProperties with no cells
   - Merge with existing row list, sort by index
4. Add test: row height round-trip
5. Add test: hidden row serialization
6. Add test: empty row with properties
7. Run tests: `./mill xl-ooxml.test`
8. Commit: "feat(ooxml): serialize row properties to <row> attributes"
```

### Phase 4: Column Macro (WI-19d) — 1 hour
```
1. Create xl-core/src/com/tjclp/xl/macros/ColumnLiteral.scala:
2. Implement col"A" string interpolator:
   - Parse column letter at compile time
   - Validate range (A-XFD)
   - Return Column opaque type
3. Export in xl-core/src/com/tjclp/xl/syntax.scala
4. Add test: col"A" compiles to Column(0)
5. Add test: col"XFD" compiles to Column(16383)
6. Add test: col"XFE" fails at compile time
7. Run tests: `./mill xl-core.test`
8. Commit: "feat(core): add col\"A\" compile-time column macro"
```

### Phase 5: Builder DSL (WI-19e) — 2 hours
```
1. Create xl-core/src/com/tjclp/xl/dsl/RowColumnDsl.scala:
2. Implement RowBuilder class:
   - Immutable builder with copy-on-modify
   - Methods: height(h), hidden, visible, outlineLevel(n), collapsed
   - Terminal: toPatch returning Patch.SetRowProperties
3. Implement ColumnBuilder class:
   - Methods: width(w), hidden, visible, outlineLevel(n), collapsed
   - Terminal: toPatch returning Patch.SetColumnProperties
4. Add entry points:
   - def row(index: Int): RowBuilder
   - def rows(range: Range): Seq[RowBuilder]
   - Extension on Row: .height(), .hidden
   - Extension on Column: .width(), .hidden
5. Export in xl-core/src/com/tjclp/xl/dsl/syntax.scala
6. Create xl-core/test/src/com/tjclp/xl/dsl/RowColumnDslSpec.scala:
7. Add tests for builder composition
8. Add tests for integration with existing patch DSL
9. Run tests: `./mill xl-core.test`
10. Commit: "feat(dsl): add row/column builder DSL"
```

### Phase 6: Final Integration (WI-19f) — 1 hour
```
1. Create xl-ooxml/test/src/com/tjclp/xl/ooxml/RowColumnPropertiesSpec.scala:
2. Add comprehensive round-trip tests:
   - Write sheet with column widths → read back → verify
   - Write sheet with row heights → read back → verify
   - Write sheet with hidden rows/cols → read back → verify
   - Write sheet with outlining → read back → verify
3. Add property-based tests:
   - Arbitrary ColumnProperties round-trip
   - Arbitrary RowProperties round-trip
4. Run full test suite: `./mill __.test`
5. Update docs/LIMITATIONS.md: remove row/column serialization limitation
6. Commit: "test(ooxml): add row/column properties round-trip tests"
7. Create PR: "feat: comprehensive row/column operations with DSL"
8. Update roadmap: WI-19 → ✅ Complete
```

---

## Design

### Current State (90% Complete)

**Domain Model** (✅ Working):
```scala
// xl-core/src/com/tjclp/xl/sheets/RowProperties.scala
final case class RowProperties(
  height: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None
)

// xl-core/src/com/tjclp/xl/sheets/ColumnProperties.scala
final case class ColumnProperties(
  width: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None
)
```

**Sheet Storage** (✅ Working):
```scala
// xl-core/src/com/tjclp/xl/sheets/Sheet.scala
final case class Sheet(
  // ...
  columnProperties: Map[Column, ColumnProperties] = Map.empty,
  rowProperties: Map[Row, RowProperties] = Map.empty,
  defaultColumnWidth: Option[Double] = None,
  defaultRowHeight: Option[Double] = None
)
```

**Patches** (✅ Working):
```scala
// xl-core/src/com/tjclp/xl/patch/Patch.scala
enum Patch:
  case SetColumnProperties(col: Column, props: ColumnProperties)
  case SetRowProperties(row: Row, props: RowProperties)
```

**OOXML Parsing** (✅ Working):
```scala
// xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala
case class OoxmlRow(
  rowIndex: Int,
  height: Option[Double],
  customHeight: Boolean,
  hidden: Boolean,
  outlineLevel: Option[Int],
  collapsed: Boolean,
  // ... all attributes parsed from XML
)
```

**Gap**: Domain properties NOT serialized to output XML

### Proposed Changes

#### 1. Extended Domain Model
```scala
// xl-core/src/com/tjclp/xl/sheets/RowProperties.scala
final case class RowProperties(
  height: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None,
  outlineLevel: Option[Int] = None,  // NEW: 0-7 for grouping depth
  collapsed: Boolean = false          // NEW: outline collapsed state
):
  require(outlineLevel.forall(l => l >= 0 && l <= 7),
    "Outline level must be 0-7")

// xl-core/src/com/tjclp/xl/sheets/ColumnProperties.scala
final case class ColumnProperties(
  width: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None,
  outlineLevel: Option[Int] = None,  // NEW
  collapsed: Boolean = false          // NEW
):
  require(outlineLevel.forall(l => l >= 0 && l <= 7),
    "Outline level must be 0-7")
```

#### 2. Column Serialization
```scala
// xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala

private def groupConsecutiveColumns(
  props: Map[Column, ColumnProperties]
): Seq[(Column, Column, ColumnProperties)] =
  // Group adjacent columns with identical properties into spans
  // (minCol, maxCol, properties)
  props.toSeq.sortBy(_._1.toIndex).foldLeft(Vector.empty[(Column, Column, ColumnProperties)]) {
    case (acc, (col, p)) if acc.nonEmpty &&
         acc.last._2.toIndex + 1 == col.toIndex &&
         acc.last._3 == p =>
      acc.init :+ (acc.last._1, col, p)
    case (acc, (col, p)) =>
      acc :+ (col, col, p)
  }

private def buildColsElement(sheet: Sheet): Option[Elem] =
  val props = sheet.columnProperties
  if props.isEmpty then None
  else
    val spans = groupConsecutiveColumns(props)
    val colElems = spans.map { case (minCol, maxCol, p) =>
      XmlUtil.elem("col",
        "min" -> (minCol.toIndex + 1).toString,
        "max" -> (maxCol.toIndex + 1).toString,
        "width" -> p.width.map(_.toString).orNull,
        "customWidth" -> (if p.width.isDefined then "1" else null),
        "hidden" -> (if p.hidden then "1" else null),
        "outlineLevel" -> p.outlineLevel.map(_.toString).orNull,
        "collapsed" -> (if p.collapsed then "1" else null)
      )()
    }
    Some(XmlUtil.elem("cols")(colElems*))
```

#### 3. Row Serialization
```scala
// In OoxmlWorksheet.fromDomainWithMetadata()

// Merge domain RowProperties with existing OoxmlRow attributes
def applyDomainRowProps(row: OoxmlRow, sheet: Sheet): OoxmlRow =
  val domainProps = sheet.getRowProperties(Row.fromOneBased(row.rowIndex))
  row.copy(
    height = domainProps.height.orElse(row.height),
    customHeight = domainProps.height.isDefined || row.customHeight,
    hidden = domainProps.hidden || row.hidden,
    outlineLevel = domainProps.outlineLevel.orElse(row.outlineLevel),
    collapsed = domainProps.collapsed || row.collapsed
  )

// Generate rows for properties-only rows (no cells)
val propsOnlyRows = sheet.rowProperties.keys
  .filterNot(existingRowIndices.contains)
  .map(r => OoxmlRow.empty(r.toOneBased).applyProps(sheet.getRowProperties(r)))
```

#### 4. Column Macro
```scala
// xl-core/src/com/tjclp/xl/macros/ColumnLiteral.scala
package com.tjclp.xl.macros

import scala.quoted.*
import com.tjclp.xl.addressing.Column

object ColumnLiteral:
  inline def col(inline s: String): Column = ${ colImpl('s) }

  private def colImpl(s: Expr[String])(using Quotes): Expr[Column] =
    import quotes.reflect.*
    s.value match
      case Some(str) =>
        Column.parse(str) match
          case Right(col) => '{ Column.fromZeroBased(${ Expr(col.toIndex) }) }
          case Left(err) => report.errorAndAbort(s"Invalid column: $err")
      case None =>
        report.errorAndAbort("Column literal must be a constant string")

// String interpolator
extension (inline sc: StringContext)
  inline def col(inline args: Any*): Column =
    ${ colInterpolatorImpl('sc, 'args) }
```

#### 5. Builder DSL
```scala
// xl-core/src/com/tjclp/xl/dsl/RowColumnDsl.scala
package com.tjclp.xl.dsl

import com.tjclp.xl.addressing.{Column, Row}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}

final class RowBuilder private[dsl] (
  private val row: Row,
  private val props: RowProperties = RowProperties()
):
  def height(h: Double): RowBuilder =
    RowBuilder(row, props.copy(height = Some(h)))
  def hidden: RowBuilder =
    RowBuilder(row, props.copy(hidden = true))
  def visible: RowBuilder =
    RowBuilder(row, props.copy(hidden = false))
  def outlineLevel(level: Int): RowBuilder =
    RowBuilder(row, props.copy(outlineLevel = Some(level)))
  def collapsed: RowBuilder =
    RowBuilder(row, props.copy(collapsed = true))
  def toPatch: Patch =
    Patch.SetRowProperties(row, props)

final class ColumnBuilder private[dsl] (
  private val col: Column,
  private val props: ColumnProperties = ColumnProperties()
):
  def width(w: Double): ColumnBuilder =
    ColumnBuilder(col, props.copy(width = Some(w)))
  def hidden: ColumnBuilder =
    ColumnBuilder(col, props.copy(hidden = true))
  def visible: ColumnBuilder =
    ColumnBuilder(col, props.copy(hidden = false))
  def outlineLevel(level: Int): ColumnBuilder =
    ColumnBuilder(col, props.copy(outlineLevel = Some(level)))
  def collapsed: ColumnBuilder =
    ColumnBuilder(col, props.copy(collapsed = true))
  def toPatch: Patch =
    Patch.SetColumnProperties(col, props)

object RowColumnDsl:
  def row(index: Int): RowBuilder =
    RowBuilder(Row.fromZeroBased(index))
  def rows(range: Range): Seq[RowBuilder] =
    range.map(i => RowBuilder(Row.fromZeroBased(i)))

  extension (r: Row)
    def height(h: Double): RowBuilder = RowBuilder(r).height(h)
    def hidden: RowBuilder = RowBuilder(r).hidden
    def outlineLevel(level: Int): RowBuilder = RowBuilder(r).outlineLevel(level)

  extension (c: Column)
    def width(w: Double): ColumnBuilder = ColumnBuilder(c).width(w)
    def hidden: ColumnBuilder = ColumnBuilder(c).hidden
    def outlineLevel(level: Int): ColumnBuilder = ColumnBuilder(c).outlineLevel(level)

export RowColumnDsl.*
```

#### 6. Usage Examples
```scala
import com.tjclp.xl.{*, given}
import com.tjclp.xl.dsl.*

// Using builder DSL with patches
val patch =
  (ref"A1" := "Revenue") ++
  (ref"B1" := "Q1") ++
  row(0).height(30.0).toPatch ++        // Header row height
  col"A".width(25.0).toPatch ++          // First column width
  col"B".width(15.0).hidden.toPatch      // Hidden column

val sheet = Sheet("Report").applyPatch(patch).unsafe

// Using extension methods directly
val sheet2 = Sheet("Data")
  .setRowProperties(Row.fromZeroBased(0), RowProperties(height = Some(25.0)))
  .setColumnProperties(Column.fromZeroBased(0), ColumnProperties(width = Some(20.0)))

// Grouping/outlining
val groupedPatch =
  rows(1 to 5).map(_.outlineLevel(1).toPatch).reduce(_ ++ _) ++
  row(0).collapsed.toPatch  // Summary row
```

---

## Definition of Done

### Phase 1: Domain Model
- [ ] `outlineLevel: Option[Int]` added to RowProperties
- [ ] `collapsed: Boolean` added to RowProperties
- [ ] Same fields added to ColumnProperties
- [ ] Validation: outline level 0-7 enforced
- [ ] Generators updated for property tests
- [ ] Existing tests pass

### Phase 2: Serialization
- [ ] `<cols>` element generated from Sheet.columnProperties
- [ ] Adjacent columns with same props merged into spans
- [ ] Row `ht`, `hidden`, `outlineLevel` attributes emitted
- [ ] Empty rows with properties serialized
- [ ] 8+ serialization tests pass
- [ ] Round-trip property tests pass

### Phase 3: DSL
- [ ] `col"A"` macro validates at compile time
- [ ] RowBuilder and ColumnBuilder implemented
- [ ] Extensions on Row and Column types
- [ ] DSL exported from `com.tjclp.xl.dsl.*`
- [ ] 6+ DSL tests pass

### Final
- [ ] All tests pass: `./mill __.test`
- [ ] Format check passes: `./mill __.checkFormat`
- [ ] LIMITATIONS.md updated (remove row/col serialization)
- [ ] PR created and merged

---

## Files Summary

### New Files
| File | Purpose | LOC |
|------|---------|-----|
| `xl-core/src/.../macros/ColumnLiteral.scala` | `col"A"` compile-time macro | ~40 |
| `xl-core/src/.../dsl/RowColumnDsl.scala` | Builder DSL classes | ~80 |
| `xl-core/test/.../dsl/RowColumnDslSpec.scala` | DSL unit tests | ~60 |
| `xl-ooxml/test/.../RowColumnPropertiesSpec.scala` | Serialization tests | ~120 |

### Modified Files
| File | Changes | LOC |
|------|---------|-----|
| `xl-core/src/.../sheets/RowProperties.scala` | Add outlineLevel, collapsed | ~8 |
| `xl-core/src/.../sheets/ColumnProperties.scala` | Add outlineLevel, collapsed | ~8 |
| `xl-ooxml/src/.../ooxml/Worksheet.scala` | Serialization helpers | ~80 |
| `xl-core/test/.../Generators.scala` | Update generators | ~8 |
| `xl-core/src/.../dsl/syntax.scala` | Export new DSL | ~3 |

**Total**: ~400 LOC new/modified

---

## Module Ownership

**Primary**: `xl-ooxml` (Worksheet.scala serialization)

**Secondary**: `xl-core` (domain model, DSL)

**Test Files**:
- `xl-ooxml/test` (RowColumnPropertiesSpec)
- `xl-core/test` (RowColumnDslSpec, updated Generators)

---

## Merge Risk Assessment

**Risk Level**: Low

**Rationale**:
- Domain model changes are additive (new fields with defaults)
- Serialization is new code paths (doesn't break existing)
- DSL is entirely new files
- No breaking API changes
- All changes are opt-in (existing code unaffected)

**Potential Conflicts**:
- WI-18 (Merged Cells) touches same file — sequence after if active
- WI-16 (Two-Phase Writer) — coordinate on Worksheet.scala changes

---

## Related Documentation

- **Roadmap**: `docs/plan/roadmap.md` (WI-19)
- **Streaming**: `docs/plan/streaming-improvements.md` (WI-19 mentioned)
- **Limitations**: `docs/LIMITATIONS.md` (row/col serialization limitation to remove)
- **STATUS**: `docs/STATUS.md` (update after completion)

---

## Deferred: autoSizeColumn (Follow-up PR)

**Reason**: Requires font metrics strategy decision

**Options for follow-up**:
1. Pure approximation (character width table) — fast, no deps, ~90% accurate
2. Java AWT fonts — precise, requires AWT availability
3. External library — most accurate, adds dependency

**When to implement**: After WI-19 complete, as separate PR

---

## Notes

- **Quick win opportunity**: WI-19a/WI-19b (serialization only) can ship first
- **DSL optional**: Can ship serialization without DSL if time-constrained
- **POI parity**: This brings XL to ~70% POI feature parity for formatting
- **Test strategy**: Property-based round-trip tests are highest value
