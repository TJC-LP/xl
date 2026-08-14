package com.tjclp.xl.workbooks

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.*

/**
 * The index must return exactly what the declaration-order linear scan it replaced returned, for
 * every table shape: duplicate names, mixed scopes, case variants (the scan used
 * `String.equalsIgnoreCase`, and Excel names are case-insensitive).
 */
class DefinedNameIndexSpec extends ScalaCheckSuite:

  /** The pre-index implementation of Evaluator.lookupDefinedName, kept as the reference. */
  private def linearReference(
    names: Vector[DefinedName],
    name: String,
    sheetIdx: Option[Int]
  ): Option[DefinedName] =
    val sheetScoped = sheetIdx.flatMap(idx =>
      names.find(dn => dn.name.equalsIgnoreCase(name) && dn.localSheetId.contains(idx))
    )
    sheetScoped.orElse(names.find(dn => dn.name.equalsIgnoreCase(name) && dn.localSheetId.isEmpty))

  // Small alphabet so duplicates and case collisions actually occur.
  private val genName: Gen[String] =
    for
      base <- Gen.oneOf("case", "mincash", "close", "circ", "print_area", "f", "g", "ｇ", "straße")
      cased <- Gen.oneOf(base, base.toUpperCase(java.util.Locale.ROOT), base.capitalize)
      suffix <- Gen.oneOf("", "1", "_x")
    yield cased + suffix

  private val genDefinedName: Gen[DefinedName] =
    for
      name <- genName
      formula <- Gen.oneOf("Sheet1!$A$1", "0.08", "Other!$B$2:$B$10")
      scope <- Gen.option(Gen.choose(0, 3))
      hidden <- Gen.oneOf(true, false)
    yield DefinedName(name, formula, scope, hidden)

  private val genTable: Gen[Vector[DefinedName]] =
    Gen.containerOf[Vector, DefinedName](genDefinedName)

  property("resolve ≡ declaration-order linear scan for every (table, query, scope)") {
    forAll(genTable, genName, Gen.option(Gen.choose(0, 4))) { (table, query, sheetIdx) =>
      val index = DefinedNameIndex(table)
      index.resolve(query, sheetIdx) ?= linearReference(table, query, sheetIdx)
    }
  }

  test("sheet-scoped entry shadows a workbook-scoped entry of the same identifier") {
    val global = DefinedName("Case", "0.08")
    val scoped = DefinedName("CASE", "Sheet1!$Z$8", localSheetId = Some(1))
    val index = DefinedNameIndex(Vector(global, scoped))
    assertEquals(index.resolve("case", Some(1)), Some(scoped))
    assertEquals(index.resolve("case", Some(0)), Some(global))
    assertEquals(index.resolve("case", None), Some(global))
  }

  test("first declared entry wins among same-name same-scope duplicates") {
    val first = DefinedName("g", "Sheet1!$A$1")
    val second = DefinedName("G", "Sheet1!$B$2")
    val index = DefinedNameIndex(Vector(first, second))
    assertEquals(index.resolve("g", None), Some(first))
  }

  test("empty table resolves nothing") {
    assertEquals(DefinedNameIndex(Vector.empty).resolve("Case", Some(0)), None)
  }
