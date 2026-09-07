package com.tjclp.xl.cli.batch

import scala.compiletime.constValueTuple
import scala.deriving.Mirror

import munit.FunSuite

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.{StreamingWriteCommands, WriteCommands}
import com.tjclp.xl.cli.helpers.BatchParser
import com.tjclp.xl.cli.helpers.BatchParser.{BatchOp, ParsedValue, StyleProps}

/**
 * The registry is THE table behind batch parsing (ADR-017 §2.6): one `OpSpec` per op name, whose
 * `cellMutating` / `streamable` flags must agree with the two hand-written classifications the
 * writers still carry, whose examples must validate against the published schema and parse back,
 * and whose names and aliases must be unique.
 */
class OpRegistrySpec extends FunSuite:

  /** The op names in the order the unknown-op error has always listed them. */
  private val historicalOrder: Vector[String] = Vector(
    "put",
    "putf",
    "style",
    "merge",
    "unmerge",
    "colwidth",
    "rowheight",
    "comment",
    "remove-comment",
    "hyperlink",
    "clear",
    "col-hide",
    "col-show",
    "row-hide",
    "row-show",
    "group-rows",
    "group-cols",
    "ungroup-rows",
    "ungroup-cols",
    "autofit",
    "add-sheet",
    "rename-sheet",
    "freeze",
    "unfreeze",
    "copy",
    "chart",
    "sheet-view",
    "tab-color",
    "autofilter",
    "page-setup",
    "header-footer",
    "cf"
  )

  /** One value per `BatchOp` case; `caseLabels` proves the vector covers the whole enum. */
  private val sample: Vector[BatchOp] = Vector(
    BatchOp.Put("A1", CellValue.Number(BigDecimal(1)), None),
    BatchOp.PutFormula("A1", "=1+1"),
    BatchOp.PutFormulaDragging("A1:A3", "=B1*2", "A1"),
    BatchOp.PutFormulas("A1:A2", Vector("=1", "=2")),
    BatchOp.PutValues(
      "A1:A2",
      Vector(
        ParsedValue(CellValue.Number(BigDecimal(1)), None),
        ParsedValue(CellValue.Number(BigDecimal(2)), None)
      )
    ),
    BatchOp.Style("A1:B2", StyleProps(bold = true)),
    BatchOp.Merge("A1:B1"),
    BatchOp.Unmerge("A1:B1"),
    BatchOp.ColWidth("A", 12.0),
    BatchOp.RowHeight(1, 20.0),
    BatchOp.AddComment("A1", "note", None),
    BatchOp.RemoveComment("A1"),
    BatchOp.Clear("A1:B2", all = true, styles = false, comments = false),
    BatchOp.ColHide("A"),
    BatchOp.ColShow("A"),
    BatchOp.RowHide(1),
    BatchOp.RowShow(1),
    BatchOp.GroupRows("2:5", 1, collapsed = false),
    BatchOp.GroupCols("B:D", 1, collapsed = false),
    BatchOp.UngroupRows("2:5"),
    BatchOp.UngroupCols("B:D"),
    BatchOp.AutoFit(None),
    BatchOp.AddSheet("New", None),
    BatchOp.RenameSheet("Old", "New"),
    BatchOp.Freeze("B2"),
    BatchOp.Unfreeze,
    BatchOp.CopyRange("A1", "B1", valuesOnly = false),
    BatchOp.Hyperlink("A1", Some("https://example.com")),
    BatchOp.AddChart("column", None, "B2:B4", None, None, None, None, None, "E2"),
    BatchOp.SetSheetView(Some(false), None, None),
    BatchOp.SetTabColor(Some("#FF0000"), clear = false),
    BatchOp.SetAutoFilter(Some("A1:C9"), clear = false),
    BatchOp.SetPageSetup(Some("landscape"), None, None, None, None),
    BatchOp.SetHeaderFooter(
      Some("&LConfidential"),
      None,
      None,
      None,
      None,
      None,
      differentOddEven = false,
      differentFirst = false
    ),
    BatchOp.AddConditionalFormat(
      "A1:A9",
      "cellIs:greaterThan:1",
      bold = true,
      italic = false,
      underline = false,
      strike = false,
      None,
      None
    )
  )

  private def label(op: BatchOp): String = op match
    case BatchOp.Unfreeze => "Unfreeze"
    case other => other.productPrefix

  private inline def labelsOf[T](using m: Mirror.SumOf[T]): List[String] =
    constValueTuple[m.MirroredElemLabels].toList.map(_.toString)

  /** Every `BatchOp` case name, straight from the compiler: a new case shows up here first. */
  private val caseLabels: Set[String] = labelsOf[BatchOp].toSet

  test("the sample covers every BatchOp case exactly once") {
    assertEquals(sample.map(label).toSet, caseLabels)
    assertEquals(sample.size, caseLabels.size)
    assertEquals(caseLabels.size, 35)
  }

  test("the registry holds the 32 op names in the historical order") {
    assertEquals(OpRegistry.all.map(_.name), historicalOrder)
  }

  test("every BatchOp case corresponds to a registered op name") {
    sample.foreach { op =>
      val name = OpRegistry.nameOf(op)
      assertEquals(OpRegistry.find(name).map(_.name), Some(name), label(op))
      assertEquals(OpRegistry.specOf(op).name, name, label(op))
    }
    assertEquals(sample.map(OpRegistry.nameOf).toSet, historicalOrder.toSet)
  }

  test("names and aliases are unique across ops, and field spellings within an op") {
    val opNames = OpRegistry.all.flatMap(s => s.name +: s.aliases)
    assertEquals(opNames.distinct, opNames)
    OpRegistry.all.foreach { spec =>
      val spellings = spec.fields.flatMap(_.spellings)
      assertEquals(spellings.distinct, spellings, s"${spec.name}: overlapping field spellings")
      assert(!spellings.contains("op"), s"${spec.name}: a field may not be spelled 'op'")
      assert(!spellings.contains("sheet"), s"${spec.name}: a field may not be spelled 'sheet'")
    }
  }

  test("cellMutating agrees with WriteCommands.isCellMutating for every op") {
    sample.foreach { op =>
      assertEquals(OpRegistry.specOf(op).cellMutating, WriteCommands.isCellMutating(op), label(op))
    }
  }

  test("a styles-only or comments-only clear is the documented exception to clear's flag") {
    val stylesOnly = BatchOp.Clear("A1", all = false, styles = true, comments = false)
    assert(OpRegistry.specOf(stylesOnly).cellMutating, "clear CAN mutate content")
    assert(!WriteCommands.isCellMutating(stylesOnly), "but a styles-only clear does not")
  }

  test("streamable agrees with the streaming writer's supported set for every op") {
    sample.foreach { op =>
      assertEquals(
        OpRegistry.specOf(op).streamable,
        StreamingWriteCommands.isStreamable(op),
        label(op)
      )
    }
  }

  test("dragging putf streams: the streaming writer shifts exactly like the in-memory path") {
    // StreamingWriteSpec pins the shifted formulas; the registry must not refuse what it honours.
    assert(StreamingWriteCommands.isStreamable(BatchOp.PutFormulaDragging("B1:B5", "=A1*2", "B1")))
    assertEquals(OpRegistry.find("putf").map(_.streamable), Some(true))
  }

  test("the 21 ops needing the whole workbook are the non-streamable ones") {
    val expected = Set(
      "comment",
      "remove-comment",
      "hyperlink",
      "clear",
      "group-rows",
      "group-cols",
      "ungroup-rows",
      "ungroup-cols",
      "autofit",
      "add-sheet",
      "rename-sheet",
      "freeze",
      "unfreeze",
      "copy",
      "chart",
      "sheet-view",
      "tab-color",
      "autofilter",
      "page-setup",
      "header-footer",
      "cf"
    )
    assertEquals(OpRegistry.all.filterNot(_.streamable).map(_.name).toSet, expected)
    assertEquals(expected.size, 21)
  }

  test("every op is sheet-scoped except add-sheet and rename-sheet") {
    assertEquals(
      OpRegistry.all.filterNot(_.sheetScoped).map(_.name).toSet,
      Set("add-sheet", "rename-sheet")
    )
  }

  test("every example validates against jsonSchema and round-trips through parseBatchJson") {
    val schema = OpRegistry.jsonSchema("test")
    OpRegistry.all.foreach { spec =>
      val problems = SchemaCheck.validate(schema, ujson.Arr(spec.example))
      assert(problems.isEmpty, s"${spec.name}: ${problems.mkString("; ")}")
      BatchParser.parseBatchJson(ujson.write(ujson.Arr(spec.example))) match
        case Right(result) =>
          assertEquals(result.ops.map(OpRegistry.nameOf), Vector(spec.name), spec.name)
          assertEquals(result.warnings, Vector.empty, s"${spec.name}: example must not warn")
          assertEquals(result.scoped.map(_.op), result.ops)
        case Left(e) => fail(s"${spec.name}: example does not parse: ${e.getMessage}")
    }
  }

  test("an unknown op and a wrong-typed field fail schema validation") {
    val schema = OpRegistry.jsonSchema("test")
    assert(SchemaCheck.validate(schema, ujson.Arr(ujson.Obj("op" -> "frobnicate"))).nonEmpty)
    val badWidth = ujson.Obj("op" -> "colwidth", "col" -> "A", "width" -> "wide")
    assert(SchemaCheck.validate(schema, ujson.Arr(badWidth)).nonEmpty)
    val missingRef = ujson.Obj("op" -> "put", "value" -> 1)
    assert(SchemaCheck.validate(schema, ujson.Arr(missingRef)).nonEmpty)
  }

  test("jsonSchema is draft 2020-12 with one branch per op carrying the x- annotations") {
    val schema = OpRegistry.jsonSchema("1.2.3")
    assertEquals(schema("$schema").str, "https://json-schema.org/draft/2020-12/schema")
    assertEquals(schema("type").str, "array")
    assert(schema("$id").str.contains("1.2.3"))
    val branches = schema("items")("oneOf").arr
    assertEquals(branches.size, 32)
    assertEquals(branches.map(_("properties")("op")("const").str).toVector, historicalOrder)
    branches.foreach { b =>
      assert(b.obj.contains("x-streamable"), b("title").str)
      assert(b.obj.contains("x-cellMutating"), b("title").str)
      assert(b.obj.contains("x-example"), b("title").str)
      assert(b.obj.contains("x-aliases"), b("title").str)
      assert(b("required").arr.map(_.str).contains("op"), b("title").str)
    }
    val putf = branches.find(_("title").str == "putf").getOrElse(fail("no putf branch"))
    assert(putf("properties").obj.contains("formula"), "the value alias is a property too")
    assert(putf("properties").obj.contains("anchor"))
    assert(putf("properties").obj.contains("sheet"))
  }

  test("find is exact, alias-aware and kebab/camel-insensitive") {
    assertEquals(OpRegistry.find("remove-comment").map(_.name), Some("remove-comment"))
    assertEquals(OpRegistry.find("removeComment").map(_.name), Some("remove-comment"))
    assertEquals(OpRegistry.find("remove_comment").map(_.name), Some("remove-comment"))
    assertEquals(OpRegistry.find("PUT").map(_.name), Some("put"))
    assertEquals(OpRegistry.find("frobnicate"), None)
    assertEquals(OpRegistry.find(""), None)
  }

  test("field spellings cover the alias and its kebab and camel forms") {
    val style = OpRegistry.find("style").getOrElse(fail("no style spec"))
    val numFormat = style.field("numFormat").getOrElse(fail("style has no numFormat"))
    assert(numFormat.spellings.contains("format"))
    assert(numFormat.spellings.contains("num-format"))
    assert(numFormat.spellings.contains("numFormat"))
    assertEquals(style.canonicalName("font-size"), Some("fontSize"))
    assertEquals(style.canonicalName("halign"), Some("align"))
    assertEquals(style.canonicalName("sheet"), Some("sheet"))
    assertEquals(style.canonicalName("op"), Some("op"))
    assertEquals(style.canonicalName("colour"), None)
    val putf = OpRegistry.find("putf").getOrElse(fail("no putf spec"))
    assertEquals(putf.canonicalName("formula"), Some("value"))
    assertEquals(putf.canonicalName("anchor"), Some("from"))
    val hyperlink = OpRegistry.find("hyperlink").getOrElse(fail("no hyperlink spec"))
    assertEquals(hyperlink.canonicalName("url"), Some("target"))
    val addSheet = OpRegistry.find("add-sheet").getOrElse(fail("no add-sheet spec"))
    assertEquals(addSheet.canonicalName("sheet"), None, "add-sheet is not sheet-scoped")
  }

  test("Spelling.camel undoes Spelling.kebab for every field name and alias") {
    // The two are only inverse when no name carries consecutive capitals (`numFmtID` would kebab
    // to `num-fmt-id` and camel back to `numFmtId`), so every name the registry accepts is checked
    val names = OpRegistry.all.flatMap(_.fields).flatMap(f => f.name +: f.aliases).distinct
    assert(names.nonEmpty)
    names.foreach { name =>
      assertEquals(Spelling.camel(Spelling.kebab(name)), name, s"round trip broke for '$name'")
      // and a kebab name is a fixpoint of kebab, so a second pass never changes a key
      assertEquals(Spelling.kebab(Spelling.kebab(name)), Spelling.kebab(name), name)
    }
    // Op names are kebab already: camel then kebab gives them back
    OpRegistry.all.flatMap(spec => spec.name +: spec.aliases).foreach { name =>
      assertEquals(Spelling.kebab(Spelling.camel(name)), name, s"op name round trip broke: '$name'")
    }
  }

  test("helpText names every op with its example") {
    val text = OpRegistry.helpText
    OpRegistry.all.foreach { spec =>
      assert(text.contains(spec.name), spec.name)
      assert(text.contains(ujson.write(spec.example)), s"${spec.name}: example missing")
    }
    assert(text.contains("\"sheet\""), "the sheet key is documented once for every op")
  }

  test("every example is a valid object of its own op with only known properties") {
    OpRegistry.all.foreach { spec =>
      assertEquals(spec.example("op").str, spec.name)
      val unknown = spec.example.value.keys.filterNot(k => spec.canonicalName(k).isDefined)
      assertEquals(unknown.toVector, Vector.empty, s"${spec.name}: unknown example keys")
    }
  }

/**
 * A structural JSON Schema checker sufficient for the batch schema: `type`, `const`, `enum`,
 * `required`, `properties`, `items`, `oneOf`, `anyOf`. Returns the problems found (empty = valid).
 */
object SchemaCheck:

  def validate(schema: ujson.Value, doc: ujson.Value, path: String = "$"): Vector[String] =
    val obj = schema.obj
    val typeProblems = obj.get("type").toList.flatMap { t =>
      val allowed = t match
        case ujson.Str(s) => Set(s)
        case ujson.Arr(xs) => xs.map(_.str).toSet
        case _ => Set.empty[String]
      val actual = typeName(doc)
      val ok = allowed.contains(actual) || (actual == "integer" && allowed.contains("number"))
      if ok then Nil
      else List(s"$path: expected type ${allowed.mkString("|")}, got $actual")
    }
    val constProblems =
      obj.get("const").filterNot(_ == doc).map(c => s"$path: expected const $c, got $doc").toList
    val enumProblems = obj
      .get("enum")
      .filterNot(_.arr.contains(doc))
      .map(_ => s"$path: $doc is not one of the enum values")
      .toList
    val requiredProblems = obj
      .get("required")
      .toList
      .flatMap(_.arr.toList.map(_.str))
      .filter(k => doc.objOpt.exists(o => !o.contains(k)))
      .map(k => s"$path: missing required '$k'")
    val propertyProblems = for
      props <- obj.get("properties").toList
      o <- doc.objOpt.toList
      (k, v) <- o.toList
      sub <- props.obj.get(k).toList
      problem <- validate(sub, v, s"$path.$k").toList
    yield problem
    val itemProblems = for
      items <- obj.get("items").toList
      arr <- doc.arrOpt.toList
      (v, i) <- arr.toList.zipWithIndex
      problem <- validate(items, v, s"$path[$i]").toList
    yield problem
    val oneOfProblems = obj.get("oneOf").toList.flatMap { alts =>
      val matching = alts.arr.count(a => validate(a, doc, path).isEmpty)
      if matching == 1 then Nil else List(s"$path: oneOf matched $matching branches")
    }
    val anyOfProblems = obj.get("anyOf").toList.flatMap { alts =>
      if alts.arr.exists(a => validate(a, doc, path).isEmpty) then Nil
      else List(s"$path: anyOf matched no branch")
    }
    (typeProblems ++ constProblems ++ enumProblems ++ requiredProblems ++ propertyProblems ++
      itemProblems ++ oneOfProblems ++ anyOfProblems).toVector

  private def typeName(v: ujson.Value): String = v match
    case _: ujson.Obj => "object"
    case _: ujson.Arr => "array"
    case _: ujson.Str => "string"
    case ujson.Num(n) => if n.isWhole then "integer" else "number"
    case ujson.True | ujson.False => "boolean"
    case ujson.Null => "null"
