package com.tjclp.xl.workbooks

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

import com.tjclp.xl.addressing.{CellRange, SheetName}
import com.tjclp.xl.error.XLError
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.{PageSetup, Sheet}

/**
 * GH-538: the mutation API matches defined names the way resolution does — case-insensitively
 * (`String.equalsIgnoreCase`, Excel's identifier relation, the relation [[DefinedNameIndex]]
 * hashes). A write replaces the whole same-scope equivalence class and appends the caller's
 * spelling; a remove drops the whole class; both stay total. GH-462: the same rule behind the
 * sheet-scoped overloads, whose scope lookup is case-insensitive like every other sheet lookup.
 */
class DefinedNameMutationSpec extends ScalaCheckSuite:

  private val model = SheetName.unsafe("Model")
  private val other = SheetName.unsafe("Other")

  private def book: Workbook = Workbook(Vector(Sheet(model), Sheet(other)))

  private def withTable(names: DefinedName*): Workbook =
    val wb = book
    wb.copy(metadata = wb.metadata.copy(definedNames = names.toVector))

  private def table(wb: Workbook): Vector[(String, String, Option[Int])] =
    wb.metadata.definedNames.map(dn => (dn.name, dn.formula, dn.localSheetId))

  // The DefinedNameIndexSpec alphabet: small, so case collisions actually occur.
  private val genName: Gen[String] =
    for
      base <- Gen.oneOf("case", "mincash", "close", "circ", "print_area", "f", "g", "ｇ", "straße")
      cased <- Gen.oneOf(base, base.toUpperCase(java.util.Locale.ROOT), base.capitalize)
      suffix <- Gen.oneOf("", "1", "_x")
    yield cased + suffix

  property("GH-538: a second write holds one entry iff the names are case-variants") {
    forAll(genName, genName) { (a, b) =>
      val wb = book.withDefinedName(a, "0.08").withDefinedName(b, "0.10")
      val entries = wb.metadata.definedNames
      ((entries.size == 1) ?= a.equalsIgnoreCase(b)) :| s"entries: $entries" &&
      (wb.metadata.definedNameIndex.resolve(b, None).map(_.formula) ?= Some("0.10")) :|
        "the second write must be the one the evaluator sees"
    }
  }

  test("GH-538: the issue's repro — defining `case` when `CASE` exists replaces it") {
    val wb = book.withDefinedName("CASE", "0.08").withDefinedName("case", "0.10")
    assertEquals(table(wb), Vector(("case", "0.10", None)))
    assertEquals(wb.metadata.definedNameIndex.resolve("case", None).map(_.formula), Some("0.10"))
    assertEquals(wb.metadata.definedNameIndex.resolve("CASE", None).map(_.formula), Some("0.10"))
  }

  test("GH-538: removeDefinedName drops the whole case class in workbook scope, nothing else") {
    val wb = withTable(
      DefinedName("CASE", "0.08"),
      DefinedName("Case", "0.09"),
      DefinedName("case", "0.10"),
      DefinedName("case1", "0.11"),
      DefinedName("case", "Model!$B$2", localSheetId = Some(0))
    )
    assertEquals(
      table(wb.removeDefinedName("case")),
      Vector(("case1", "0.11", None), ("case", "Model!$B$2", Some(0)))
    )
  }

  test("GH-538: removing an absent name is the identity on the table (total)") {
    val wb = withTable(DefinedName("Total", "Model!$A$1"))
    assertEquals(wb.removeDefinedName("Nope").metadata.definedNames, wb.metadata.definedNames)
    assertEquals(book.removeDefinedName("Nope").metadata.definedNames, Vector.empty)
  }

  test("GH-528 corpus shape: a write over case-colliding duplicates leaves exactly one entry") {
    val wb = withTable(DefinedName("g", "Sheet1!$A$1"), DefinedName("G", "Sheet1!$B$2"))
      .withDefinedName("g", "Sheet1!$C$3")
    assertEquals(table(wb), Vector(("g", "Sheet1!$C$3", None)))
    assertEquals(
      wb.metadata.definedNameIndex.resolve("G", None).map(_.formula),
      Some("Sheet1!$C$3"),
      "no shadowed survivor: the index resolves the written entry"
    )
  }

  // Black-box over arbitrary strings (ſ, ı, µ, surrogates): the mutation rule IS the resolution
  // rule, so nothing a write matches can differ from what the evaluator then resolves.
  property("GH-538: DefinedName.sameName ≡ an index hit, for arbitrary names") {
    forAll { (declared: String, query: String) =>
      DefinedName.sameName(declared, query) ?=
        DefinedNameIndex(Vector(DefinedName(declared, "0"))).resolve(query, None).isDefined
    }
  }

  test("GH-538: sameName folds supplementary-plane case pairs by code point, like the index") {
    val capitalLongI = new String(Character.toChars(0x10400))
    val smallLongI = new String(Character.toChars(0x10428))
    assert(DefinedName.sameName(capitalLongI, smallLongI))
    assertEquals(
      table(book.withDefinedName(capitalLongI, "1").withDefinedName(smallLongI, "2")),
      Vector((smallLongI, "2", None))
    )
  }

  private def right(result: Either[XLError, Workbook]): Workbook =
    result.fold(e => fail(e.message), identity)

  test("GH-462: the scoped overload sets localSheetId to the sheet's position, global untouched") {
    val wb = right(book.withDefinedName("Rate", "0.08").withDefinedName("Rate", "0.09", other))
    assertEquals(table(wb), Vector(("Rate", "0.08", None), ("Rate", "0.09", Some(1))))
    val index = wb.metadata.definedNameIndex
    assertEquals(index.resolve("rate", Some(1)).map(_.formula), Some("0.09"))
    assertEquals(index.resolve("rate", Some(0)).map(_.formula), Some("0.08"))
    assertEquals(index.resolve("rate", None).map(_.formula), Some("0.08"))
  }

  test("GH-462: an unknown scope is SheetNotFound naming the sheets, on add and on remove") {
    val nope = SheetName.unsafe("Nope")
    val expected = Left(XLError.SheetNotFound("Nope", Vector("Model", "Other")))
    assertEquals(book.withDefinedName("Rate", "0.08", nope), expected)
    assertEquals(book.removeDefinedName("Rate", nope), expected)
  }

  test("GH-462: the scope is matched case-insensitively, like every other sheet lookup") {
    val wb = right(book.withDefinedName("Local", "Other!$A$1", SheetName.unsafe("OTHER")))
    assertEquals(table(wb), Vector(("Local", "Other!$A$1", Some(1))))
  }

  test("GH-538: a scoped replace is case-insensitive within its scope") {
    val wb = right(
      book
        .withDefinedName("RATE", "0.08", other)
        .flatMap(_.withDefinedName("rate", "0.09", other))
    )
    assertEquals(table(wb), Vector(("rate", "0.09", Some(1))))
  }

  test("GH-462: the scoped remove drops only that scope's class; the global remove leaves it") {
    val wb = withTable(
      DefinedName("Rate", "0.08"),
      DefinedName("RATE", "0.09", localSheetId = Some(1)),
      DefinedName("rate", "0.10", localSheetId = Some(1)),
      DefinedName("Rate", "0.11", localSheetId = Some(0))
    )
    assertEquals(
      table(right(wb.removeDefinedName("rate", other))),
      Vector(("Rate", "0.08", None), ("Rate", "0.11", Some(0)))
    )
    assertEquals(
      table(wb.removeDefinedName("rate")),
      Vector(("RATE", "0.09", Some(1)), ("rate", "0.10", Some(1)), ("Rate", "0.11", Some(0)))
    )
    assertEquals(
      right(wb.removeDefinedName("Nope", other)).metadata.definedNames,
      wb.metadata.definedNames
    )
  }

  // GH-462: after a read a sheet's `_xlnm.Print_Area` / `_xlnm.Print_Titles` live in its PageSetup
  // (the reader lifts them out of the table), so this is the post-read shape of a book with both.
  private val lifted: Workbook = Workbook(
    Vector(
      Sheet(model).withPageSetup(
        PageSetup(printArea = Some(ref"A1:B2"), repeatRows = Some((1, 2)))
      ),
      Sheet(other)
    )
  )

  private def printFields(wb: Workbook): (Option[CellRange], Option[(Int, Int)]) =
    val setup = wb.sheets(0).pageSetup
    (setup.flatMap(_.printArea), setup.flatMap(_.repeatRows))

  test("GH-462: removeDefinedName of a lifted print name clears the PageSetup field it lives in") {
    val noArea = right(lifted.removeDefinedName("_xlnm.print_area", model))
    assertEquals(printFields(noArea), (None, Some((1, 2))))
    assertEquals(noArea.metadata.definedNames, Vector.empty)
    assertEquals(
      printFields(right(noArea.removeDefinedName("_XLNM.PRINT_TITLES", model))),
      (None, None)
    )
    // the other sheet's scope, the workbook scope and every other identifier leave it alone
    val untouched = lifted.sheets(0).pageSetup
    assertEquals(
      right(lifted.removeDefinedName("_xlnm.Print_Area", other)).sheets(0).pageSetup,
      untouched
    )
    assertEquals(lifted.removeDefinedName("_xlnm.Print_Area").sheets(0).pageSetup, untouched)
    assertEquals(right(lifted.removeDefinedName("Nope", model)).sheets(0).pageSetup, untouched)
  }

  test(
    "GH-462: withDefinedName of a print name clears the lifted field — the entry is the one truth"
  ) {
    val wb = right(lifted.withDefinedName("_xlnm.Print_Area", "Model!$C$1:$D$2", model))
    assertEquals(table(wb), Vector(("_xlnm.Print_Area", "Model!$C$1:$D$2", Some(0))))
    assertEquals(printFields(wb), (None, Some((1, 2))))
    assertEquals(
      printFields(right(lifted.withDefinedName("Local", "1", model))),
      printFields(lifted)
    )
  }

  test("define then remove is the identity on the table, in either scope and any spelling") {
    val base = withTable(DefinedName("Total", "Model!$A$1"))
    assertEquals(
      base.withDefinedName("X", "1").removeDefinedName("x").metadata.definedNames,
      base.metadata.definedNames
    )
    assertEquals(
      base
        .withDefinedName("X", "1", other)
        .flatMap(_.removeDefinedName("x", other))
        .map(_.metadata.definedNames),
      Right(base.metadata.definedNames)
    )
  }
