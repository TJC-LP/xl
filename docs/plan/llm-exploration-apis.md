# LLM-Friendly Exploration & Output APIs

**Status**: Design Spec
**Target**: xl-core (primary), xl-evaluator (formula flow)
**Priority**: High - Core ergonomics for LLM use cases

---

## Executive Summary

This spec defines APIs for LLM-friendly spreadsheet exploration and rendering. The design assumes **adversarial unstructured data** (financial models, LBO models, DCF spreadsheets) where traditional tabular assumptions fail.

**Core insight**: Financial models are NOT tabular. They have scattered labels, merged cells, multiple logical sections, and spatial relationships that matter more than column structure.

**Design philosophy**:
- No header detection
- No column type inference
- Spatial-first exploration
- Always include cell references in output
- Composable query → view → render pattern

---

## API Overview

```scala
import com.tjclp.xl.{*, given}

// Core pattern: select a range, then render or explore
sheet(ref"A1:D20").toMarkdown
sheet(ref"A1:D20").toHtml
sheet(ref"A1:D20").toText
sheet(ref"A1:D20").labels
sheet(ref"A1:D20").regions

// Consistent with existing single-cell access
sheet(ref"A1")      // Returns Cell (existing API)
sheet(ref"A1:D20")  // Returns RangeView (new)

// Sheet-level operations
sheet.describe
sheet.formulaFlow  // (xl-evaluator module)
```

---

## 1. RangeView - The Core Type

### 1.1 Definition

```scala
package com.tjclp.xl.explore

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.Cell
import com.tjclp.xl.sheets.Sheet

/**
 * An immutable view into a rectangular region of a sheet.
 *
 * RangeView is the primary type for LLM-friendly exploration and rendering.
 * It captures a snapshot of cells within a range and provides methods for:
 * - Rendering (toMarkdown, toHtml, toText)
 * - Exploration (labels, regions)
 * - Metadata (cellCount, hasFormulas, etc.)
 *
 * Design note: RangeView is intentionally simple - it holds the cells
 * at construction time rather than being lazy. For very large ranges,
 * consider using streaming APIs instead.
 */
final case class RangeView(
  sheet: Sheet,
  range: CellRange,
  cells: Map[ARef, Cell]
):
  // === RENDERING ===

  /** Render as Markdown table with row/column references */
  def toMarkdown: String

  /** Render as HTML table with row/column references */
  def toHtml: String

  /** Render as HTML with options */
  def toHtml(includeStyles: Boolean = true, includeRefs: Boolean = true): String

  /** Render as plain text grid */
  def toText: String

  /** Render as plain text with column width limit */
  def toText(maxWidth: Int): String

  // === EXPLORATION ===

  /** Find label-value pairs (text cell adjacent to number/formula cell) */
  def labels: Vector[LabeledValue]

  /** Find contiguous data regions (separated by empty rows/columns) */
  def regions: Vector[DataRegion]

  // === METADATA ===

  /** Total cells in range (including empty) */
  def cellCount: Int = range.cellCount

  /** Non-empty cells in range */
  def nonEmptyCount: Int = cells.size

  /** Whether any cell contains a formula */
  def hasFormulas: Boolean = cells.values.exists(_.value.isFormula)

  /** Whether range overlaps any merged regions */
  def hasMergedCells: Boolean =
    sheet.mergedRanges.exists(mr => range.intersects(mr))
```

### 1.2 Construction via sheet(range) apply syntax

```scala
package com.tjclp.xl.sheets

// Add to Sheet companion or as extension
extension (sheet: Sheet)
  /**
   * Select a range of cells, returning a RangeView for exploration/rendering.
   *
   * This extends the existing `sheet(ref)` pattern (which returns Cell) to
   * support ranges. The `ref"..."` macro distinguishes types at compile time:
   * - `ref"A1"` → ARef → Cell
   * - `ref"A1:D20"` → CellRange → RangeView
   *
   * Example:
   * {{{
   * sheet(ref"A1:D20").toMarkdown
   * sheet(ref"B5:E10").labels
   * }}}
   */
  def apply(range: CellRange): RangeView =
    val cells = range.cells
      .flatMap(ref => sheet.cells.get(ref).map(ref -> _))
      .toMap
    RangeView(sheet, range, cells)
```

**Note**: The existing `sheet(ref: ARef): Cell` method remains unchanged. The two overloads coexist naturally due to distinct parameter types.

---

## 2. Rendering APIs

### 2.1 toMarkdown

**Purpose**: Render range as Markdown table for LLM consumption.

**Critical design decisions**:
- **Always include row numbers and column letters** - LLMs need spatial references
- **No header detection** - First row is just first row, not assumed to be headers
- **No alignment inference** - All cells left-aligned
- **Show formulas as-is** - Display `=SUM(A1:A10)` not evaluated result

**Output format**:
```markdown
|   | A | B | C |
|---|---|---|---|
| 1 | Revenue | | $1,000 |
| 2 | COGS | | $400 |
| 3 | | | |
| 4 | Gross Profit | | =C1-C2 |
```

**Implementation notes**:
```scala
def toMarkdown: String =
  val sb = new StringBuilder
  val startCol = range.start.col.index0
  val endCol = range.end.col.index0
  val startRow = range.start.row.index0
  val endRow = range.end.row.index0

  // Header row with column letters
  sb.append("|   |")
  for col <- startCol to endCol do
    sb.append(s" ${Column.from0(col).toA1} |")
  sb.append("\n")

  // Separator row
  sb.append("|---|")
  for _ <- startCol to endCol do
    sb.append("---|")
  sb.append("\n")

  // Data rows with row numbers
  for row <- startRow to endRow do
    sb.append(s"| ${row + 1} |")  // 1-based row number
    for col <- startCol to endCol do
      val ref = ARef.from0(col, row)
      val value = cells.get(ref).map(formatCellValue).getOrElse("")
      sb.append(s" ${escapeMarkdown(value)} |")
    sb.append("\n")

  sb.toString

private def escapeMarkdown(s: String): String =
  s.replace("|", "\\|").replace("\n", " ")

private def formatCellValue(cell: Cell): String =
  cell.value match
    case CellValue.Formula(expr, _) => expr  // Show raw formula
    case CellValue.Text(s) => s
    case CellValue.Number(n) => formatNumber(n, cell.styleId)
    case CellValue.Bool(b) => if b then "TRUE" else "FALSE"
    case CellValue.DateTime(dt) => dt.toLocalDate.toString
    case CellValue.Error(err) => err.toExcel
    case CellValue.RichText(rt) => rt.toPlainText
    case CellValue.Empty => ""
```

### 2.2 toHtml

**Purpose**: Render range as HTML table with optional styling.

**Output format**:
```html
<table>
  <thead>
    <tr><th></th><th>A</th><th>B</th><th>C</th></tr>
  </thead>
  <tbody>
    <tr><th>1</th><td>Revenue</td><td></td><td style="...">$1,000</td></tr>
    <tr><th>2</th><td>COGS</td><td></td><td style="...">$400</td></tr>
  </tbody>
</table>
```

**Parameters**:
- `includeStyles: Boolean = true` - Include inline CSS from cell styles
- `includeRefs: Boolean = true` - Include row/column headers (should almost always be true)

**Integration**: Builds on existing `HtmlRenderer` but adds ref headers.

### 2.3 toText

**Purpose**: Plain text grid for terminal output.

**Output format**:
```
     A            B            C
1    Revenue                   $1,000
2    COGS                      $400
3
4    Gross Profit              =C1-C2
```

**Parameters**:
- `maxWidth: Int` - Optional column width limit (truncates with `...`)

**Implementation notes**:
- Calculate column widths based on content (capped at maxWidth if provided)
- Right-pad text values, left-pad numeric values
- Show row numbers left-aligned

---

## 3. Exploration APIs

### 3.1 labels - Label-Value Pair Detection

**Purpose**: Find "Revenue: $1,000" patterns - text cells adjacent to number/formula cells.

**Types**:
```scala
/**
 * A detected label-value pair in a spreadsheet.
 *
 * Labels are text cells that appear to describe adjacent numeric values.
 * Common patterns in financial models:
 * - Label in column A, value in column B (LeftOf)
 * - Label in row above value (Above)
 */
final case class LabeledValue(
  label: LabelCell,
  value: ValueCell,
  relationship: LabelRelationship
)

final case class LabelCell(
  ref: ARef,
  text: String
)

final case class ValueCell(
  ref: ARef,
  value: CellValue,
  formatted: String  // Display-formatted value
)

enum LabelRelationship:
  case LeftOf   // Label immediately left of value (most common)
  case Above    // Label immediately above value
```

**Algorithm**:
```scala
def labels: Vector[LabeledValue] =
  cells.collect {
    case (ref, cell) if isValueCell(cell) =>
      findAdjacentLabel(ref, cell)
  }.flatten.toVector

private def isValueCell(cell: Cell): Boolean =
  cell.value match
    case CellValue.Number(_) | CellValue.Formula(_, _) |
         CellValue.DateTime(_) => true
    case _ => false

private def findAdjacentLabel(valueRef: ARef, valueCell: Cell): Option[LabeledValue] =
  // Check left first (most common in financial models)
  val leftRef = valueRef.offset(col = -1, row = 0)
  cells.get(leftRef).filter(isTextCell).map { labelCell =>
    LabeledValue(
      LabelCell(leftRef, extractText(labelCell)),
      ValueCell(valueRef, valueCell.value, formatCellValue(valueCell)),
      LabelRelationship.LeftOf
    )
  }.orElse {
    // Check above
    val aboveRef = valueRef.offset(col = 0, row = -1)
    cells.get(aboveRef).filter(isTextCell).map { labelCell =>
      LabeledValue(
        LabelCell(aboveRef, extractText(labelCell)),
        ValueCell(valueRef, valueCell.value, formatCellValue(valueCell)),
        LabelRelationship.Above
      )
    }
  }
```

**Use case**:
```scala
sheet(ref"B5:E25").labels.foreach { lv =>
  println(s"${lv.label.text}: ${lv.value.formatted}")
}
// Output:
// "Entry EV: $500M"
// "Exit Multiple: 8.0x"
// "IRR: 22.5%"
```

### 3.2 regions - Contiguous Data Cluster Detection

**Purpose**: Find distinct data sections within a range, separated by empty rows/columns.

**Types**:
```scala
/**
 * A contiguous region of non-empty cells.
 *
 * Financial models often have multiple logical sections on one sheet:
 * - Assumptions section
 * - Calculations section
 * - Output/summary section
 *
 * Regions are detected by finding connected components of non-empty cells,
 * where cells are "connected" if they share an edge (not diagonal).
 */
final case class DataRegion(
  bounds: CellRange,       // Bounding box of the region
  cellCount: Int,          // Number of non-empty cells
  density: Double,         // cellCount / (bounds area) - how "full" is the region
  hasFormulas: Boolean     // Whether region contains any formulas
)
```

**Algorithm** (flood-fill / connected components):
```scala
def regions: Vector[DataRegion] =
  val nonEmpty = cells.keySet
  val visited = mutable.Set.empty[ARef]
  val result = Vector.newBuilder[DataRegion]

  for ref <- nonEmpty if !visited(ref) do
    val regionCells = floodFill(ref, nonEmpty, visited)
    result += DataRegion.fromCells(regionCells, cells)

  result.result().sortBy(r => (r.bounds.start.row.index0, r.bounds.start.col.index0))

private def floodFill(
  start: ARef,
  nonEmpty: Set[ARef],
  visited: mutable.Set[ARef]
): Set[ARef] =
  val queue = mutable.Queue(start)
  val region = mutable.Set.empty[ARef]

  while queue.nonEmpty do
    val ref = queue.dequeue()
    if nonEmpty.contains(ref) && !visited(ref) then
      visited += ref
      region += ref
      // Add 4-connected neighbors (not diagonal)
      queue ++= Seq(
        ref.offset(-1, 0), ref.offset(1, 0),
        ref.offset(0, -1), ref.offset(0, 1)
      ).filter(nonEmpty.contains).filterNot(visited.contains)

  region.toSet
```

**Use case**:
```scala
sheet(ref"A1:Z100").regions.foreach { r =>
  println(s"Region at ${r.bounds}: ${r.cellCount} cells, ${r.density * 100}% dense")
}
// Output:
// "Region at A1:D15: 42 cells, 70% dense"
// "Region at A20:F35: 84 cells, 88% dense"
// "Region at A40:C45: 12 cells, 67% dense"
```

---

## 4. Sheet-Level APIs

### 4.1 sheet.describe - Spatial Summary

**Purpose**: Quick one-line summary for LLM context injection.

**Types**:
```scala
/**
 * Summary description of a sheet's contents.
 *
 * Designed for efficient LLM context injection - provides
 * a compact overview without requiring full sheet traversal.
 */
final case class SheetDescription(
  name: String,
  usedRange: Option[CellRange],
  cellCount: Int,
  nonEmptyCount: Int,
  formulaCount: Int,
  regionCount: Int,
  topLabels: Vector[String]  // First N notable labels (e.g., "Revenue", "EBITDA")
):
  /** Compact single-line summary */
  override def toString: String =
    val rangeStr = usedRange.map(_.toA1).getOrElse("empty")
    val labelsStr = if topLabels.isEmpty then ""
                    else s" | Labels: ${topLabels.take(5).mkString(", ")}"
    s"$name | $rangeStr | $nonEmptyCount cells | $formulaCount formulas$labelsStr"

  /** Multi-line detailed breakdown */
  def detailed: String =
    s"""Sheet: $name
       |Used Range: ${usedRange.map(_.toA1).getOrElse("(empty)")}
       |Cells: $nonEmptyCount non-empty / $cellCount total
       |Formulas: $formulaCount
       |Regions: $regionCount distinct data regions
       |Top Labels: ${topLabels.mkString(", ")}""".stripMargin
```

**Extension**:
```scala
extension (sheet: Sheet)
  def describe: SheetDescription =
    val used = sheet.usedRange
    val view = used.map(r => sheet(r))
    SheetDescription(
      name = sheet.name.value,
      usedRange = used,
      cellCount = used.map(_.cellCount).getOrElse(0),
      nonEmptyCount = sheet.cells.size,
      formulaCount = sheet.cells.values.count(_.value.isFormula),
      regionCount = view.map(_.regions.size).getOrElse(0),
      topLabels = view.map(_.labels.take(10).map(_.label.text)).getOrElse(Vector.empty)
    )
```

**Use case**:
```scala
println(sheet.describe)
// "Model | A1:Z200 | 847 cells | 142 formulas | Labels: Revenue, EBITDA, Exit Multiple"
```

### 4.2 sheet.formulaFlow (xl-evaluator module)

**Purpose**: Understand calculation graph - what feeds what.

**Types**:
```scala
package com.tjclp.xl.formula

/**
 * Analysis of formula dependencies in a sheet.
 *
 * Classifies cells into three categories:
 * - Inputs: Raw values that formulas depend on (the "assumptions")
 * - Calculations: Intermediate formula cells
 * - Outputs: Formula cells that nothing else depends on (the "results")
 *
 * This is critical for understanding financial models where:
 * - Inputs are often assumptions (growth rate, discount rate, etc.)
 * - Outputs are often KPIs (IRR, NPV, exit value, etc.)
 */
final case class FormulaFlow(
  inputs: Vector[InputCell],
  calculations: Vector[FormulaCell],
  outputs: Vector[OutputCell],
  hasCycles: Boolean
)

final case class InputCell(
  ref: ARef,
  label: Option[String],  // Adjacent label if found
  value: CellValue
)

final case class FormulaCell(
  ref: ARef,
  formula: String,
  dependsOn: Vector[ARef]
)

final case class OutputCell(
  ref: ARef,
  formula: Option[String],
  value: CellValue
)
```

**Extension** (in xl-evaluator):
```scala
extension (sheet: Sheet)
  def formulaFlow: FormulaFlow =
    val graph = DependencyGraph.fromSheet(sheet)

    // Inputs: cells that formulas depend on but are not formulas themselves
    val formulaCells = sheet.cells.filter((_, c) => c.value.isFormula).keySet
    val allDeps = graph.dependencies.values.flatten.toSet
    val inputs = (allDeps -- formulaCells).toVector.map { ref =>
      val label = findAdjacentLabel(sheet, ref)
      InputCell(ref, label, sheet(ref).value)
    }

    // Outputs: formula cells that nothing depends on
    val dependedOn = graph.dependents.filter((_, deps) => deps.nonEmpty).keySet
    val outputs = (formulaCells -- dependedOn).toVector.map { ref =>
      val cell = sheet(ref)
      OutputCell(ref, Some(cell.value.asFormula.get), cell.value)
    }

    // Calculations: formula cells that are depended on
    val calculations = (formulaCells & dependedOn).toVector.map { ref =>
      val cell = sheet(ref)
      FormulaCell(ref, cell.value.asFormula.get, graph.dependencies(ref).toVector)
    }

    FormulaFlow(
      inputs = inputs.sortBy(_.ref),
      calculations = calculations.sortBy(_.ref),
      outputs = outputs.sortBy(_.ref),
      hasCycles = graph.hasCycles
    )
```

**Use case**:
```scala
val flow = sheet.formulaFlow
println(s"Inputs (assumptions): ${flow.inputs.size}")
flow.inputs.foreach { i =>
  println(s"  ${i.ref.toA1}: ${i.label.getOrElse("?")} = ${i.value}")
}
println(s"Outputs (results): ${flow.outputs.size}")
flow.outputs.foreach { o =>
  println(s"  ${o.ref.toA1}: ${o.formula.getOrElse("value")}")
}
if flow.hasCycles then
  println("WARNING: Circular references detected!")
```

---

## 5. File Structure

### New Files

```
xl-core/src/com/tjclp/xl/explore/
├── RangeView.scala        # Core RangeView type
├── LabelDetector.scala    # LabeledValue, LabelCell, ValueCell, LabelRelationship
├── RegionDetector.scala   # DataRegion and flood-fill algorithm
├── SheetDescribe.scala    # SheetDescription
├── MarkdownRenderer.scala # toMarkdown implementation
├── TextRenderer.scala     # toText implementation
├── HtmlRenderer.scala     # Enhanced toHtml (or modify existing)
└── syntax.scala           # Extension methods (sheet.apply(range), sheet.describe)

xl-evaluator/src/com/tjclp/xl/formula/
└── FormulaFlow.scala      # FormulaFlow, InputCell, OutputCell, sheet.formulaFlow
```

### Modified Files

| File | Changes |
|------|---------|
| `xl-core/src/com/tjclp/xl/syntax.scala` | Export `explore.*` |
| `xl-core/src/com/tjclp/xl/api.scala` | Export `RangeView`, `LabeledValue`, etc. |
| `xl-evaluator/src/com/tjclp/xl/exports.scala` | Export `FormulaFlow` |

---

## 6. Testing Strategy

### Property Tests

```scala
// Region coverage - no cells left behind
property("regions covers all non-empty cells") {
  forAll(genSheetWithData) { sheet =>
    val view = sheet(sheet.usedRange.get)
    val regionCells = view.regions.flatMap(r => r.bounds.cells).toSet
    val nonEmptyCells = view.cells.keySet
    assert(nonEmptyCells.subsetOf(regionCells))
  }
}

// Regions are non-overlapping
property("regions are disjoint") {
  forAll(genSheetWithData) { sheet =>
    val view = sheet(sheet.usedRange.get)
    val regions = view.regions
    for
      r1 <- regions
      r2 <- regions if r1 != r2
    do assert(!r1.bounds.intersects(r2.bounds))
  }
}

// toMarkdown preserves cell count
property("toMarkdown has correct dimensions") {
  forAll(genRangeView) { view =>
    val lines = view.toMarkdown.split("\n")
    val dataLines = lines.drop(2)  // Skip header and separator
    val expectedRows = view.range.rowCount
    assertEquals(dataLines.length, expectedRows)
  }
}

// Labels only match text -> numeric patterns
property("labels have text label and numeric value") {
  forAll(genRangeView) { view =>
    view.labels.foreach { lv =>
      assert(isTextValue(view.cells(lv.label.ref).value))
      assert(isNumericValue(view.cells(lv.value.ref).value))
    }
  }
}
```

### Unit Tests

```scala
test("toMarkdown escapes pipe characters") {
  val sheet = Sheet("Test").put(ref"A1", "A|B").unsafe
  val md = sheet(ref"A1:A1").toMarkdown
  assert(md.contains("A\\|B"))
}

test("labels finds left-of pattern") {
  val sheet = Sheet("Test")
    .put(ref"A1", "Revenue")
    .put(ref"B1", 1000)
    .unsafe
  val labels = sheet(ref"A1:B1").labels
  assertEquals(labels.size, 1)
  assertEquals(labels.head.label.text, "Revenue")
  assertEquals(labels.head.relationship, LabelRelationship.LeftOf)
}

test("regions detects gap-separated clusters") {
  val sheet = Sheet("Test")
    .put(ref"A1", "X").put(ref"A2", "Y")  // Region 1
    .put(ref"A5", "Z").put(ref"A6", "W")  // Region 2 (gap at rows 3-4)
    .unsafe
  val regions = sheet(ref"A1:A6").regions
  assertEquals(regions.size, 2)
}

test("describe produces compact summary") {
  val sheet = Sheet("Model")
    .put(ref"A1", "Revenue").put(ref"B1", 1000)
    .put(ref"A2", "=B1*2")
    .unsafe
  val desc = sheet.describe
  assert(desc.toString.contains("Model"))
  assert(desc.toString.contains("1 formulas"))
}
```

---

## 7. Open Design Questions

### 7.1 Method Naming: ✅ RESOLVED → `apply` syntax

**Decision**: Use `sheet(range)` apply syntax, consistent with existing `sheet(ref)` for single cells.

```scala
sheet(ref"A1")      // Returns Cell (existing)
sheet(ref"A1:D20")  // Returns RangeView (new)
```

**Rationale**:
- Type-safe dispatch via distinct parameter types (`ARef` vs `CellRange`)
- Extends existing mental model rather than introducing new method
- More ergonomic: `sheet(ref"A1:D20").toMarkdown`
- Scala idiomatic: `collection(index)` pattern

### 7.2 Whole-Sheet Shorthand

Should `sheet(sheet.usedRange.get).toMarkdown` have a shorthand?

```scala
// Option A: No shorthand (explicit is good)
sheet(sheet.usedRange.get).toMarkdown

// Option B: Convenience method
sheet.toMarkdown  // Implies usedRange

// Option C: Default parameter (not possible with apply)
```

**Recommendation**: Option A - explicit is better for LLM use cases where ranges matter.

### 7.3 Large Range Handling

Should `toMarkdown` truncate at some limit?

**Recommendation**: No automatic truncation. Caller controls range size. Could add `toMarkdown(maxRows: Int)` overload if needed.

### 7.4 Formula Display

Show raw formula, evaluated value, or both?

```scala
// Option A: Raw formula
| 4 | =SUM(A1:A3) |

// Option B: Evaluated value
| 4 | 1500 |

// Option C: Both
| 4 | 1500 [=SUM(A1:A3)] |
```

**Recommendation**: Option A (raw formula) by default. LLMs need to understand the calculation structure. Could add `toMarkdown(evaluateFormulas: Boolean)` option.

---

## 8. Implementation Order

### Phase 1: Core Rendering (Highest Priority)
1. `RangeView` type definition
2. `sheet(range)` apply extension (add to Sheet)
3. `toMarkdown` implementation
4. Tests for markdown rendering

### Phase 2: Additional Renderers
1. `toText` implementation
2. `toHtml` enhancement (add includeRefs)
3. Tests

### Phase 3: Exploration APIs
1. `LabeledValue` types
2. `labels` detection algorithm
3. `DataRegion` types
4. `regions` flood-fill algorithm
5. Tests

### Phase 4: Sheet-Level APIs
1. `SheetDescription` type
2. `sheet.describe` implementation
3. Tests

### Phase 5: Formula Flow (xl-evaluator)
1. `FormulaFlow` types
2. `sheet.formulaFlow` implementation
3. Integration tests with real financial models

---

## 9. Future Considerations

### 9.1 Streaming for Large Sheets

For very large sheets, consider a streaming variant:
```scala
def toMarkdownStream: fs2.Stream[F, String]  // Line by line
```

### 9.2 Additional Exploration

Potential future APIs:
```scala
view.numbers      // All numeric cells
view.formulas     // All formula cells
view.errors       // All error cells (#REF!, #DIV/0!, etc.)
view.timeline     // Detect time series (dates in sequence)
```

### 9.3 Smart Truncation

For LLM context window management:
```scala
view.toMarkdown(maxTokens: Int)  // Estimate and truncate
view.summarize(targetSize: Int)  // Intelligent sampling
```
