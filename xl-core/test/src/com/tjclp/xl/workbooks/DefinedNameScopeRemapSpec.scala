package com.tjclp.xl.workbooks

import com.tjclp.xl.Workbook
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.sheets.Sheet
import munit.FunSuite

/**
 * GH-434: sheet-scoped defined names (`localSheetId`) are POSITIONAL — every sheet-order mutation
 * (removeAt / insertAt / reorder) must remap the stored indices or the names silently attach to
 * whatever sheet now occupies their old position. Modeled print names are immune (PrintNames
 * re-derives them from live sheet position on write); these tests pin the generic names and the
 * deliberately-unmodelable verbatim print names that live in `metadata.definedNames`.
 */
class DefinedNameScopeRemapSpec extends FunSuite:

  private def name(s: String): SheetName = SheetName.unsafe(s)

  /** Three sheets [Alpha, Beta, Gamma] with one sheet-scoped name each plus a global name. */
  private def workbook: Workbook =
    val base = Workbook(Vector(Sheet(name("Alpha")), Sheet(name("Beta")), Sheet(name("Gamma"))))
    base.copy(metadata =
      base.metadata.copy(definedNames =
        Vector(
          DefinedName("AlphaLocal", "Alpha!$A$1", localSheetId = Some(0)),
          DefinedName("BetaLocal", "Beta!$B$2", localSheetId = Some(1)),
          DefinedName("GammaLocal", "Gamma!$C$3", localSheetId = Some(2)),
          DefinedName("Global", "0.08", localSheetId = None)
        )
      )
    )

  private def scopeOf(wb: Workbook, dn: String): Option[Int] =
    wb.metadata.definedNames.find(_.name == dn).flatMap(_.localSheetId)

  test("GH-434: removeAt(0) shifts higher scopes down and drops the removed sheet's names") {
    val wb = workbook.removeAt(0).fold(e => fail(s"removeAt failed: $e"), identity)
    assertEquals(wb.sheetNames.map(_.value), Vector("Beta", "Gamma"))
    assertEquals(
      wb.metadata.definedNames.find(_.name == "AlphaLocal"),
      None,
      "names scoped to the removed sheet must be dropped"
    )
    assertEquals(scopeOf(wb, "BetaLocal"), Some(0), "Beta moved 1 -> 0")
    assertEquals(scopeOf(wb, "GammaLocal"), Some(1), "Gamma moved 2 -> 1")
    assertEquals(scopeOf(wb, "Global"), None, "workbook-scoped names untouched")
  }

  test("GH-434: removeAt(middle) leaves lower scopes alone") {
    val wb = workbook.removeAt(1).fold(e => fail(s"removeAt failed: $e"), identity)
    assertEquals(scopeOf(wb, "AlphaLocal"), Some(0))
    assertEquals(wb.metadata.definedNames.find(_.name == "BetaLocal"), None)
    assertEquals(scopeOf(wb, "GammaLocal"), Some(1))
  }

  test("GH-434: remove by name routes through the same remap") {
    val wb = workbook.remove(name("Alpha")).fold(e => fail(s"remove failed: $e"), identity)
    assertEquals(scopeOf(wb, "BetaLocal"), Some(0))
    assertEquals(scopeOf(wb, "GammaLocal"), Some(1))
  }

  test("GH-434: insertAt(middle) shifts scopes at and above the insertion point up") {
    val wb = workbook
      .insertAt(1, Sheet(name("Inserted")))
      .fold(e => fail(s"insertAt failed: $e"), identity)
    assertEquals(wb.sheetNames.map(_.value), Vector("Alpha", "Inserted", "Beta", "Gamma"))
    assertEquals(scopeOf(wb, "AlphaLocal"), Some(0), "below insertion point: unchanged")
    assertEquals(scopeOf(wb, "BetaLocal"), Some(2), "Beta moved 1 -> 2")
    assertEquals(scopeOf(wb, "GammaLocal"), Some(3), "Gamma moved 2 -> 3")
    assertEquals(scopeOf(wb, "Global"), None)
  }

  test("GH-434: insertAt(0) shifts every scope up") {
    val wb = workbook
      .insertAt(0, Sheet(name("First")))
      .fold(e => fail(s"insertAt failed: $e"), identity)
    assertEquals(scopeOf(wb, "AlphaLocal"), Some(1))
    assertEquals(scopeOf(wb, "BetaLocal"), Some(2))
    assertEquals(scopeOf(wb, "GammaLocal"), Some(3))
  }

  test("GH-434: reorder applies the full permutation to every sheet-scoped name") {
    val wb = workbook
      .reorder(Vector(name("Gamma"), name("Alpha"), name("Beta")))
      .fold(e => fail(s"reorder failed: $e"), identity)
    assertEquals(wb.sheetNames.map(_.value), Vector("Gamma", "Alpha", "Beta"))
    assertEquals(scopeOf(wb, "AlphaLocal"), Some(1), "Alpha moved 0 -> 1")
    assertEquals(scopeOf(wb, "BetaLocal"), Some(2), "Beta moved 1 -> 2")
    assertEquals(scopeOf(wb, "GammaLocal"), Some(0), "Gamma moved 2 -> 0")
    assertEquals(scopeOf(wb, "Global"), None)
  }

  test("GH-434: identity reorder keeps every name where it is") {
    val wb = workbook
      .reorder(Vector(name("Alpha"), name("Beta"), name("Gamma")))
      .fold(e => fail(s"reorder failed: $e"), identity)
    assertEquals(scopeOf(wb, "AlphaLocal"), Some(0))
    assertEquals(scopeOf(wb, "BetaLocal"), Some(1))
    assertEquals(scopeOf(wb, "GammaLocal"), Some(2))
  }

  test("GH-434: verbatim print names (unmodelable shapes) remap like any sheet-scoped name") {
    // Multi-range Print_Area and column-span Print_Titles never ride the PageSetup lift
    // (PrintNames.scala keeps them in metadata verbatim) — the remap must cover them too.
    val base = Workbook(Vector(Sheet(name("Alpha")), Sheet(name("Beta"))))
    val wb0 = base.copy(metadata =
      base.metadata.copy(definedNames =
        Vector(
          DefinedName(
            "_xlnm.Print_Area",
            "Beta!$A$1:$B$2,Beta!$D$1:$E$2",
            localSheetId = Some(1)
          ),
          DefinedName("_xlnm.Print_Titles", "Beta!$A:$B", localSheetId = Some(1))
        )
      )
    )
    val wb = wb0.removeAt(0).fold(e => fail(s"removeAt failed: $e"), identity)
    assertEquals(
      wb.metadata.definedNames.map(dn => (dn.name, dn.localSheetId)),
      Vector(
        ("_xlnm.Print_Area", Some(0)),
        ("_xlnm.Print_Titles", Some(0))
      ),
      "verbatim print names must follow their sheet across a removal"
    )
    // The formulas themselves are untouched (name-based, not positional)
    assert(wb.metadata.definedNames.forall(_.formula.startsWith("Beta!")))
  }
