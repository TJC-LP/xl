// Deliberately OUTSIDE com.tjclp.xl (see ScriptingPreludeTest): `Hyperlink` is an opaque type, so
// its alias + companion val pair and the companion extension must work exactly as a script or an
// external library sees them — through each of the two public imports, with no import of
// `com.tjclp.xl.cells`.
package xlprelude

import munit.FunSuite

/**
 * Gate test for `Cell.hyperlink: Option[Hyperlink]` (#891) through the scripting prelude: the type,
 * the factory, `target` chained on a factory result (the export landmine for opaque companions),
 * and the point of the type — a `String` in `Cell`'s fourth slot, where the removed `Cell.comment`
 * used to be, is a compile error.
 */
class HyperlinkPreludeProbe extends FunSuite:
  import com.tjclp.xl.scripting.{*, given}

  test("Hyperlink type, factory and target resolve through the prelude"):
    val link: Hyperlink = Hyperlink("https://example.com")
    assertEquals(link.target, "https://example.com")
    assertEquals(Hyperlink("Sheet1!A1").target, "Sheet1!A1") // chained on the factory result
    val inferred = Hyperlink("mailto:a@example.com")
    assertEquals(inferred.target, "mailto:a@example.com")

  test("Cell carries an Option[Hyperlink]; both withHyperlink overloads and named construction"):
    val viaString: Cell = Cell(ref"A1", CellValue.Text("Docs")).withHyperlink("#Sheet1!B2")
    val read: Option[Hyperlink] = viaString.hyperlink
    assertEquals(read.map(_.target), Some("Sheet1!B2"))
    val viaLink = Cell(ref"A1", CellValue.Text("Docs")).withHyperlink(Hyperlink("Sheet1!B2"))
    assertEquals(viaLink, viaString)
    val named = Cell(ref"A1", CellValue.Text("Docs"), hyperlink = Some(Hyperlink("Sheet1!B2")))
    assertEquals(named, viaString)
    val sheet = Sheet("Sheet1").put(named)
    assertEquals(sheet(ref"A1").hyperlink.map(_.target), Some("Sheet1!B2"))
    assertEquals(sheet(ref"A1").clearHyperlink.hyperlink, None)

  test("#931/#932: range hyperlinks set, read and remove through the prelude"):
    val range: CellRange = ref"A1:A5"
    val sheet = Sheet("Sheet1").setHyperlink(range, Hyperlink("https://example.com"))
    val ranged: RangeHyperlinks = sheet.rangeHyperlinks
    assertEquals(ranged, RangeHyperlinks(RangeHyperlink(range, Hyperlink("https://example.com"))))
    assertEquals(ranged.toVector.map(_.range), Vector(range))
    assertEquals(ranged.at(ref"A4").map(_.link.target), Some("https://example.com"))
    assertEquals(RangeHyperlinks.empty.isEmpty, true)
    val built = RangeHyperlinks.from(Vector(RangeHyperlink(range, Hyperlink("Sheet1!A1"))))
    assertEquals(
      Sheet("S").copy(rangeHyperlinks = built)(ref"A2").hyperlink.map(_.target),
      Some("Sheet1!A1")
    )
    assertEquals(sheet(ref"A3").hyperlink.map(_.target), Some("https://example.com"))
    assertEquals(sheet.hyperlinkAt(ref"A5").map(_.target), Some("https://example.com"))
    assertEquals(sheet.rangeHyperlinkAt(ref"A2").map(_.range), Some(range))
    val removed = sheet.removeHyperlinks(ref"A2:A3")
    assertEquals(removed.hyperlinkAt(ref"A2"), None)
    assertEquals(removed.hyperlinkAt(ref"A4").map(_.target), Some("https://example.com"))

  test("a String in the old comment slot does not compile through the prelude"):
    val errors = compileErrors("""Cell(ref"A1", CellValue.Text("Q3"), None, Some("review note"))""")
    assert(errors.contains("Hyperlink"), errors)
    assert(errors.contains("Found") && errors.contains("Required"), errors)
    assertEquals(
      compileErrors("""Cell(ref"A1", CellValue.Text("Q3"), None, Some(Hyperlink("https://x")))"""),
      ""
    )
    assertEquals(compileErrors("""Cell(ref"A1", CellValue.Text("Q3"), None)"""), "")

/** The same surface through the pure import, `com.tjclp.xl.{*, given}`. */
class HyperlinkBaseImportProbe extends FunSuite:
  import com.tjclp.xl.{*, given}

  test("Hyperlink type, factory and target resolve through the base import"):
    val link: Hyperlink = Hyperlink("https://example.com")
    assertEquals(link.target, "https://example.com")
    assertEquals(Hyperlink("Sheet1!A1").target, "Sheet1!A1")
    val cell = Cell(ref"A1", CellValue.Text("Docs"), hyperlink = Some(link))
    assertEquals(cell.hyperlink.map(_.target), Some("https://example.com"))
    assertEquals(cell.withHyperlink("#Sheet1!B2").hyperlink, Some(Hyperlink("Sheet1!B2")))

  test("#931/#932: range hyperlinks resolve through the base import"):
    val sheet = Sheet("Sheet1").setHyperlink(ref"B2:C3", Hyperlink("Sheet1!A1"))
    val ranged: RangeHyperlinks = sheet.rangeHyperlinks
    assertEquals(ranged.toVector.map(_.link.target), Vector("Sheet1!A1"))
    assertEquals(RangeHyperlinks.from(ranged.iterator), ranged)
    assertEquals(sheet(ref"C3").hyperlink, Some(Hyperlink("Sheet1!A1")))

  test("a String in the old comment slot does not compile through the base import"):
    val errors = compileErrors("""Cell(ref"A1", CellValue.Text("Q3"), None, Some("review note"))""")
    assert(errors.contains("Hyperlink"), errors)
    assert(errors.contains("Found") && errors.contains("Required"), errors)
    assertEquals(
      compileErrors("""Cell(ref"A1", CellValue.Text("Q3"), None, Some(Hyperlink("https://x")))"""),
      ""
    )
