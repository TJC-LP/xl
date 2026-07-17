package com.tjclp.xl.ooxml

import scala.xml.Elem

import munit.FunSuite

/**
 * Guards the workbook-level preservation invariant (GH-397), mirroring
 * WorksheetPreservationInvariantSpec (GH-232).
 *
 * Every CT_Workbook child label must be either modeled by a dedicated `OoxmlWorkbook` field /
 * regenerated (`sheets`), OR position-slotted via the `preservedAfter*` groups. A label that is
 * neither rides the `otherElements` catch-all and is re-emitted AFTER extLst — a CT_Workbook
 * child-order violation Excel repairs loudly, produced on every surgical write (workbook.xml always
 * regenerates).
 *
 * If this fails after you add a member to `canonicalChildOrder`, either give it a dedicated field
 * (and list it in `modeledOrRegenerated` below) or add it to a `preservedAfter*` group in
 * OoxmlWorkbook at its canonical position.
 */
class WorkbookPreservationInvariantSpec extends FunSuite:

  // Labels backed by a dedicated OoxmlWorkbook field, or regenerated from the domain model
  // (sheets from workbook.sheets). mc:AlternateContent / xr:revisionPtr are also dedicated
  // fields but are markup-compatibility members, not CT_Workbook sequence members.
  private val modeledOrRegenerated: Set[String] = Set(
    "fileVersion",
    "workbookPr",
    "bookViews",
    "sheets",
    "definedNames",
    "calcPr",
    "extLst"
  )

  test("every CT_Workbook member is modeled or position-slotted (no after-extLst emission)") {
    val orphans =
      OoxmlWorkbook.canonicalChildOrder.toSet -- modeledOrRegenerated -- OoxmlWorkbook.slottedLabels
    assert(
      orphans.isEmpty,
      "These CT_Workbook members are neither modeled nor slotted, so they are re-emitted after " +
        s"extLst — the GH-397 child-order violation: ${orphans.toSeq.sorted.mkString(", ")}"
    )
  }

  test("slotted labels are disjoint from modeled labels (no double emit)") {
    val overlap = modeledOrRegenerated.intersect(OoxmlWorkbook.slottedLabels)
    assert(overlap.isEmpty, s"labels both modeled and slotted would emit twice: $overlap")
  }

  test("canonicalChildOrder is exactly modeled + slotted (no stragglers either way)") {
    assertEquals(
      OoxmlWorkbook.canonicalChildOrder.toSet,
      modeledOrRegenerated ++ OoxmlWorkbook.slottedLabels
    )
  }

  test("canonicalChildOrder has no duplicate labels") {
    assertEquals(OoxmlWorkbook.canonicalChildOrder.distinct, OoxmlWorkbook.canonicalChildOrder)
  }

  test("the emission sequence (fields + slot groups) IS the canonical CT_Workbook order") {
    // This pins the writer's slot positions to the schema: field emission points interleaved with
    // the preservedAfter* groups must reproduce canonicalChildOrder exactly. If a group moves or a
    // member lands in the wrong group, this fails before any fixture does.
    val emissionOrder =
      Vector("fileVersion") ++ OoxmlWorkbook.preservedAfterFileVersion ++
        Vector("workbookPr") ++ OoxmlWorkbook.preservedAfterWorkbookPr ++
        Vector("bookViews", "sheets") ++ OoxmlWorkbook.preservedAfterSheets ++
        Vector("definedNames", "calcPr") ++ OoxmlWorkbook.preservedAfterCalcPr ++
        Vector("extLst")
    assertEquals(emissionOrder, OoxmlWorkbook.canonicalChildOrder)
  }

  test("toXml slots preserved otherElements into canonical positions (behavioral)") {
    // otherElements deliberately shuffled: emission must still be schema-ordered, with unknown
    // labels (no canonical slot) kept last.
    def raw(label: String): Elem = XmlUtil.elem(label)()
    val wb = OoxmlWorkbook
      .minimal()
      .copy(
        definedNames = Some(raw("definedNames")),
        extLst = Some(raw("extLst")),
        otherElements = Seq(
          raw("pivotCaches"),
          raw("externalReferences"),
          raw("workbookProtection"),
          raw("vendorUnknown")
        )
      )
    val labels = wb.toXml.child.collect { case e: Elem => e.label }.toVector
    assertEquals(
      labels,
      Vector(
        "workbookProtection",
        "sheets",
        "externalReferences",
        "definedNames",
        "pivotCaches",
        "extLst",
        "vendorUnknown"
      )
    )
  }
