package com.tjclp.xl.formula

import munit.FunSuite

import com.tjclp.xl.formula.functions.FunctionRegistry
import com.tjclp.xl.ooxml.FormulaStorage

/**
 * GH-556: the registry and the `_xlfn.` storage list must agree on which evaluable functions are
 * post-2007. A registry function Excel introduced after 2007 that is missing from
 * `FormulaStorage.FutureFunctions` would be written bare and show `#NAME?` in Excel — exactly the
 * bug GH-556 fixed — so the intersection is pinned: adding a post-2007 function to the evaluator
 * fails here until the storage list (and this pin) are updated together.
 */
class FutureFunctionRegistrySpec extends FunSuite:

  test("GH-556: registry functions that take the _xlfn. prefix are exactly the pinned set") {
    val prefixed = FunctionRegistry.allNames.filter(FormulaStorage.FutureFunctions.contains).toSet
    assertEquals(
      prefixed,
      Set(
        "FILTER",
        "IFNA",
        "IFS",
        "MAXIFS",
        "MINIFS",
        "SEQUENCE",
        "SORT",
        "SWITCH",
        "UNIQUE",
        "XLOOKUP"
      )
    )
    // LET is a parser special form, not a registry entry, and is still stored as _xlfn.LET
    assert(FormulaStorage.FutureFunctions.contains("LET"))
    assert(!FunctionRegistry.allNames.contains("LET"))
  }

  test("GH-556: every registry function parses with an _xlfn. prefix (inherited-file parity)") {
    // Only functions Excel actually prefixes appear that way in files, but the parser's rule is
    // uniform: the prefix is dropped before lookup, whatever the function.
    FunctionRegistry.allNames.foreach { name =>
      assertEquals(FormulaStorage.bareFunctionName(s"_xlfn.$name"), name, name)
    }
  }

  test("GH-556: storage laws hold for every registry function (fromStored ∘ toStored = id, ...)") {
    FunctionRegistry.allNames.foreach { name =>
      val model = s"IF($name(A1:A3,B1)>0,$name(A1:A3,B1),0)"
      val stored = FormulaStorage.toStored(model)
      assertEquals(FormulaStorage.fromStored(stored), model, name)
      assertEquals(FormulaStorage.toStored(stored), stored, name)
      assertEquals(FormulaStorage.toStored(FormulaStorage.fromStored(stored)), stored, name)
    }
  }
