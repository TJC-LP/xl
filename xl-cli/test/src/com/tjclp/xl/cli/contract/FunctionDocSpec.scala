package com.tjclp.xl.cli.contract

import munit.FunSuite

import com.tjclp.xl.formula.functions.FunctionRegistry

/**
 * [[FunctionDoc]]: the typed row `xl functions --json` prints for every registry function, plus
 * `LET` (a parser-level special form the registry does not hold). The count is the registry's plus
 * one, every row's `args` agree with its arity, and the JSON shape has exactly the eight keys.
 */
class FunctionDocSpec extends FunSuite:

  test("115 registry functions + LET, sorted by name, LET last and the only special form") {
    val docs = FunctionDoc.all
    assertEquals(docs.size, FunctionRegistry.allNames.size + 1)
    assertEquals(docs.count(_.specialForm), 1)
    assertEquals(docs.lastOption.map(_.name), Some("LET"))
    assertEquals(docs.filterNot(_.specialForm).map(_.name), FunctionRegistry.allNames.toVector)
    assertEquals(docs.map(_.name).distinct.size, docs.size, "names are unique")
    assert(docs.forall(d => d.name == d.name.toUpperCase), "names are upper-case")
  }

  test("every FunctionDoc has args that agree with its arity") {
    FunctionDoc.all.foreach { doc =>
      val clue = s"${doc.name}: minArgs=${doc.minArgs} maxArgs=${doc.maxArgs} args=${doc.args}"
      assert(doc.minArgs >= 0, clue)
      doc.maxArgs.foreach(max => assert(max >= doc.minArgs, clue))
      // A variadic spec describes its repeated slot once and has no upper bound; a fixed-arity
      // spec names every slot (optional ones included), so it lists at least minArgs of them.
      val variadic = doc.maxArgs.isEmpty
      if variadic then assert(doc.args.nonEmpty || doc.minArgs == 0, clue)
      else assert(doc.args.size >= doc.minArgs, clue)
      assert(doc.args.forall(_.trim.nonEmpty), clue)
    }
  }

  test(
    "flags come from the registry: TODAY/NOW return dates and times, INDIRECT/OFFSET are dynamic"
  ) {
    val byName = FunctionDoc.all.map(d => d.name -> d).toMap
    assertEquals(byName.get("TODAY").map(_.returnsDate), Some(true))
    assertEquals(byName.get("NOW").map(_.returnsTime), Some(true))
    assertEquals(byName.get("INDIRECT").map(_.dynamicDeps), Some(true))
    assertEquals(byName.get("OFFSET").map(_.dynamicDeps), Some(true))
    assertEquals(byName.get("SUM").map(d => (d.minArgs, d.maxArgs)), Some((1, None)))
    assertEquals(byName.get("LET").map(_.specialForm), Some(true))
    assertEquals(
      FunctionDoc.all.filter(_.dynamicDeps).map(_.name),
      FunctionRegistry.dynamicFunctionNames.toVector
    )
  }

  test("GH-588: volatile comes from FunctionFlags.volatile — TODAY, NOW, RAND, RANDBETWEEN") {
    val byName = FunctionDoc.all.map(d => d.name -> d).toMap
    Vector("TODAY", "NOW", "RAND", "RANDBETWEEN").foreach { name =>
      assertEquals(byName.get(name).map(_.volatile), Some(true), name)
    }
    assertEquals(byName.get("SUM").map(_.volatile), Some(false))
    assertEquals(byName.get("INDIRECT").map(_.volatile), Some(false), "dynamic deps ≠ volatile")
    assertEquals(byName.get("LET").map(_.volatile), Some(false))
    assertEquals(
      FunctionDoc.all.filter(_.volatile).map(_.name),
      FunctionRegistry.volatileFunctionNames.toVector
    )
  }

  test("toJson: exactly the nine keys, typed; maxArgs is null for a variadic function") {
    val keys = List(
      "name",
      "minArgs",
      "maxArgs",
      "args",
      "returnsDate",
      "returnsTime",
      "dynamicDeps",
      "volatile",
      "specialForm"
    )
    FunctionDoc.all.foreach { doc =>
      val json = doc.toJson
      assertEquals(json.value.keys.toList, keys, doc.name)
      assertEquals(json("name"), ujson.Str(doc.name))
      assertEquals(json("minArgs"), ujson.Num(doc.minArgs))
      assertEquals(json("maxArgs"), doc.maxArgs.fold[ujson.Value](ujson.Null)(n => ujson.Num(n)))
      assertEquals(json("args"), ujson.Arr.from(doc.args.map(ujson.Str.apply)))
      assertEquals(json("returnsDate"), ujson.Bool(doc.returnsDate))
      assertEquals(json("returnsTime"), ujson.Bool(doc.returnsTime))
      assertEquals(json("dynamicDeps"), ujson.Bool(doc.dynamicDeps))
      assertEquals(json("volatile"), ujson.Bool(doc.volatile))
      assertEquals(json("specialForm"), ujson.Bool(doc.specialForm))
    }
    val all = FunctionDoc.toJson(FunctionDoc.all)
    assertEquals(all.arr.size, FunctionDoc.all.size)
    assertEquals(all.arr.lastOption.map(_("maxArgs")), Some(ujson.Null), "LET is variadic")
  }
