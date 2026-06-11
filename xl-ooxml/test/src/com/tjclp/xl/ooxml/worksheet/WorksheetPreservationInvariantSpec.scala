package com.tjclp.xl.ooxml.worksheet

import munit.FunSuite

/**
 * Guards the surgical-modify preservation invariant (GH-232).
 *
 * Every inline worksheet element the reader recognizes (`worksheetKnownElements`) is excluded from
 * the `otherElements` verbatim catch-all, so each MUST be either modeled by a dedicated
 * `OoxmlWorksheet` field / regenerated, OR captured in `preservedKnown`. An element that is neither
 * is silently dropped on any sheet modification — the exact "xl ate my dropdowns" failure GH-232
 * fixed.
 *
 * If this fails after you add a label to `worksheetKnownElements`, either give it a dedicated field
 * (and list it in `modeledOrRegenerated` below) or add it to a `preservedAfter*` group in
 * WorksheetHelpers.
 */
class WorksheetPreservationInvariantSpec extends FunSuite:

  // Labels backed by a dedicated OoxmlWorksheet field, or regenerated from the domain model
  // (sheetData from cells, mergeCells from mergedRanges).
  private val modeledOrRegenerated: Set[String] = Set(
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "mergeCells",
    "conditionalFormatting",
    "printOptions",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "drawing",
    "legacyDrawing",
    "picture",
    "oleObjects",
    "controls",
    "tableParts",
    "extLst"
  )

  test("every known worksheet element is modeled or preserved (no tier-3 silent data loss)") {
    val orphans = worksheetKnownElements -- modeledOrRegenerated -- preservedKnownLabels
    assert(
      orphans.isEmpty,
      "These knownElements are excluded from verbatim copy but neither modeled nor preserved, so " +
        s"they are silently dropped on modify (GH-232 class): ${orphans.toSeq.sorted.mkString(", ")}"
    )
  }

  test("preservedKnown labels are disjoint from modeled labels (no double emit)") {
    val overlap = modeledOrRegenerated.intersect(preservedKnownLabels)
    assert(overlap.isEmpty, s"labels both modeled and preserved would emit twice: $overlap")
  }

  test("worksheetKnownElements is exactly modeled + preserved (no stragglers either way)") {
    assertEquals(worksheetKnownElements, modeledOrRegenerated ++ preservedKnownLabels)
  }
