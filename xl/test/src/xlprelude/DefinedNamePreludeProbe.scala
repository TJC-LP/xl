// Deliberately OUTSIDE com.tjclp.xl (see ScriptingPreludeTest): the defined-name surface must
// resolve through `import com.tjclp.xl.scripting.{*, given}` alone, exactly as a script sees it.
package xlprelude

import munit.FunSuite

import com.tjclp.xl.scripting.{*, given}

/**
 * Gate test for the defined-name API through the prelude (GH-462 / GH-538): the `DefinedName` type
 * and companion — `PrintArea`, `PrintTitles`, `sameName`, the `matches` extension — plus
 * `Workbook.withDefinedName` / `removeDefinedName` in both scopes, all spelled without an import of
 * `com.tjclp.xl.workbooks`. Compile success is most of the test; the assertions pin that the
 * companion relation is the one the mutation paths match by.
 */
class DefinedNamePreludeProbe extends FunSuite:

  private val Data = SheetName.unsafe("Data")

  private val book: Workbook =
    Workbook(Sheet(Data).put(ref"A1", 10).put(ref"A2", 20))

  test("DefinedName type, companion constants and sameName resolve through the prelude"):
    val dn: DefinedName = DefinedName("Revenue", "Data!$A$1:$A$2")
    assertEquals(dn.name, "Revenue")
    assertEquals(dn.localSheetId, None)
    assertEquals(DefinedName.PrintArea, "_xlnm.Print_Area")
    assertEquals(DefinedName.PrintTitles, "_xlnm.Print_Titles")
    assert(DefinedName.sameName("revenue", "REVENUE"))
    assert(!DefinedName.sameName("revenue", "revenues"))
    // the companion extension survives the export hop (a plain case class, not an opaque type)
    assert(dn.matches("REVENUE", None))
    assert(!dn.matches("REVENUE", Some(0)))

  test("Workbook.withDefinedName in both scopes, read back as Vector[DefinedName]"):
    val global = book.withDefinedName("Revenue", "Data!$A$1:$A$2")
    val names: Vector[DefinedName] = global.metadata.definedNames
    assert(names.exists(_.matches("revenue", None)))
    assert(names.exists(n => DefinedName.sameName(n.name, "REVENUE")))

    val scoped = global.withDefinedName(DefinedName.PrintArea, "Data!$A$1:$A$2", Data).unsafe
    assert(scoped.metadata.definedNames.exists(_.matches(DefinedName.PrintArea, Some(0))))

    // removal matches by the same relation: a differently-cased name removes the entry
    val removed = scoped.removeDefinedName("REVENUE")
    assert(!removed.metadata.definedNames.exists(_.matches("Revenue", None)))
    assert(removed.metadata.definedNames.exists(_.matches(DefinedName.PrintArea, Some(0))))
