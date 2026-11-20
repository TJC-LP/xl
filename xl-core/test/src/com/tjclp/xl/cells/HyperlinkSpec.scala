package com.tjclp.xl.cells

import munit.FunSuite

import com.tjclp.xl.addressing.ARef

/**
 * `Cell.hyperlink` is an `Option[Hyperlink]`, not an `Option[String]` (#891).
 *
 * `Cell`'s fourth parameter was the cell comment until `Cell.comment` was removed, so
 * `Cell(ref, value, None, Some("review note"))` compiled (with a deprecation warning) into a
 * comment. With a `String` hyperlink in that slot the same call would keep compiling and write a
 * link instead. These tests pin that it is now a type error, and that every intended spelling still
 * compiles.
 */
class HyperlinkSpec extends FunSuite:

  private val ref: ARef = ARef.from0(0, 0)
  private val value: CellValue = CellValue.Text("Q3")

  test("a positional String in the old comment slot does not compile"):
    val errors = compileErrors("""Cell(ref, value, None, Some("review note"))""")
    assert(errors.contains("Hyperlink"), errors)
    assert(errors.contains("Found") && errors.contains("Required"), errors)

  test("a named String hyperlink does not compile either"):
    val errors = compileErrors("""Cell(ref, value, hyperlink = Some("https://example.com"))""")
    assert(errors.contains("Hyperlink"), errors)

  test("positive controls: the same calls compile with a Hyperlink, and without the argument"):
    assertEquals(
      compileErrors("""Cell(ref, value, None, Some(Hyperlink("https://example.com")))"""),
      ""
    )
    assertEquals(
      compileErrors("""Cell(ref, value, hyperlink = Some(Hyperlink("https://example.com")))"""),
      ""
    )
    assertEquals(compileErrors("""Cell(ref, value, None)"""), "")
    assertEquals(compileErrors("""Cell(ref, value)"""), "")
    assertEquals(compileErrors("""Cell(ref, value).withHyperlink("https://example.com")"""), "")

  test("the 3-arg and named-arg forms build the cells they say"):
    assertEquals(Cell(ref, value, None).hyperlink, None)
    val named = Cell(ref, value, hyperlink = Some(Hyperlink("https://example.com")))
    assertEquals(named.hyperlink.map(_.target), Some("https://example.com"))
    assertEquals(named, Cell(ref, value, None, Some(Hyperlink("https://example.com"))))

  test("Hyperlink holds its target verbatim; withHyperlink(String) is where '#' is stripped"):
    assertEquals(Hyperlink("#Sheet1!B2").target, "#Sheet1!B2")
    assertEquals(
      Cell(ref, value).withHyperlink("#Sheet1!B2").hyperlink,
      Some(Hyperlink("Sheet1!B2"))
    )
    assertEquals(
      Cell(ref, value).withHyperlink(Hyperlink("#Sheet1!B2")).hyperlink,
      Some(Hyperlink("#Sheet1!B2"))
    )
    // only a LEADING '#' is stripped: URL fragments survive
    assertEquals(
      Cell(ref, value).withHyperlink("https://x/page#s").hyperlink.map(_.target),
      Some("https://x/page#s")
    )

  test("withHyperlink overloads agree on a target with no leading '#'; clearHyperlink clears"):
    val viaString = Cell(ref, value).withHyperlink("https://example.com")
    val viaLink = Cell(ref, value).withHyperlink(Hyperlink("https://example.com"))
    assertEquals(viaString, viaLink)
    assertEquals(viaLink.clearHyperlink.hyperlink, None)
