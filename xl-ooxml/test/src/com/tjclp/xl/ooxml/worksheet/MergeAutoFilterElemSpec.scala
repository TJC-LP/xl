package com.tjclp.xl.ooxml.worksheet

import scala.xml.Elem

import munit.FunSuite

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.XmlSecurity
import com.tjclp.xl.sheets.AutoFilterState

/** GH-429: the autoFilter lift-and-overlay helper (ref-only rewrite, children verbatim). */
class MergeAutoFilterElemSpec extends FunSuite:

  private def xml(s: String): Elem =
    XmlSecurity.parseSafe(s, "test").fold(e => fail(s"parse failed: ${e.message}"), identity)

  private val withChildren = xml(
    """<autoFilter ref="A1:C6"><filterColumn colId="0"><filters><filter val="North"/></filters></filterColumn><sortState ref="A2:C6"/></autoFilter>"""
  )

  test("None passes the source element through untouched") {
    assertEquals(mergeAutoFilterElem(Some(withChildren), None), Some(withChildren))
    assertEquals(mergeAutoFilterElem(None, None), None)
  }

  test("Ranged replaces ONLY the ref attribute; filterColumn/sortState children ride verbatim") {
    val merged = mergeAutoFilterElem(
      Some(withChildren),
      Some(AutoFilterState.Ranged(ref"A3:C8"))
    ).getOrElse(fail("expected an element"))
    assertEquals(merged \@ "ref", "A3:C8")
    assertEquals(merged.child.toList, withChildren.child.toList) // same Nodes, byte-verbatim
    // sortState@ref rides stale by design (documented limitation)
    assertEquals((merged \ "sortState").headOption.map(_ \@ "ref"), Some("A2:C6"))
  }

  test("identity fast-path: a Ranged equal to the source ref returns the source element itself") {
    val merged = mergeAutoFilterElem(
      Some(withChildren),
      Some(AutoFilterState.Ranged(ref"A1:C6"))
    )
    assert(merged.exists(_ eq withChildren), "must be the SAME Elem instance (zero churn)")
  }

  test("Ranged with no source materializes a fresh minimal element; Remove deletes") {
    val fresh = mergeAutoFilterElem(None, Some(AutoFilterState.Ranged(ref"B2:D9")))
      .getOrElse(fail("expected an element"))
    assertEquals(fresh.label, "autoFilter")
    assertEquals(fresh \@ "ref", "B2:D9")
    assertEquals(fresh.child.isEmpty, true)
    assertEquals(mergeAutoFilterElem(Some(withChildren), Some(AutoFilterState.Remove)), None)
    assertEquals(mergeAutoFilterElem(None, Some(AutoFilterState.Remove)), None)
  }

  test("a single-cell source ref parses as a 1x1 range (identity holds for it)") {
    val single = xml("""<autoFilter ref="A1"/>""")
    assertEquals(parseAutoFilterRef(single), Some(CellRange(ref"A1", ref"A1")))
    val merged = mergeAutoFilterElem(
      Some(single),
      Some(AutoFilterState.Ranged(CellRange(ref"A1", ref"A1")))
    )
    assert(merged.exists(_ eq single), "1x1 promotion must hit the identity fast-path")
  }
