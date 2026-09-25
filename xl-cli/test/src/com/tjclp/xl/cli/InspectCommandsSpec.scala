package com.tjclp.xl.cli

import cats.effect.{IO, Resource}
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cli.contract.{Argv, CliHarness, CliRun, EnvelopeSchema, TestFixtures}
import com.tjclp.xl.io.ExcelIO

/**
 * The inspection verbs of ADR-017 §2.10 through the in-process harness: `describe [--full]`,
 * `audit [--fail-on-findings]` and `deps <ref> [--direction] [--depth] [--expand]`, each typed
 * under `--json` and readable otherwise; plus `cell`'s graph lines, which list a formula's inputs
 * as declared (a range is one entry, as in `deps`). Byte-level renderings are pinned by the goldens
 * (`describe*`, `audit*`, `deps*`, `cell*`); this suite pins the shapes and the exit-code contract.
 */
class InspectCommandsSpec extends CatsEffectSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "inspect-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString

  /** The envelope a `--json` run printed, validated against the schema. */
  private def envelope(run: CliRun): ujson.Obj =
    val parsed = ujson.read(run.stdout)
    EnvelopeSchema.assertValid(parsed)
    assertEquals(parsed("exitCode"), ujson.Num(run.exit), "exitCode mirrors the process exit code")
    parsed.obj

  private def data(run: CliRun): ujson.Value =
    val e = envelope(run)
    assertEquals(e("ok"), ujson.True, run.stderr)
    e("data")

  private def names(arr: ujson.Value): Vector[String] = arr.arr.toVector.map(_.str)

  // ---------------------------------------------------------------------------------------------
  // describe
  // ---------------------------------------------------------------------------------------------

  test("describe --json: sheets with state and dimension, defined names with the hidden flag") {
    CliHarness.run("-f", file("linked.xlsx"), "--json", "describe").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val d = data(run)
      assertEquals(d.obj.keys.toList, List("sheets", "definedNames", "date1904"))
      val sheets = d("sheets").arr.toVector
      assertEquals(sheets.map(_("name").str), Vector("Sheet1", "Sheet2", "Hidden"))
      assertEquals(sheets.map(_("index").num.toInt), Vector(1, 2, 3))
      assertEquals(sheets.map(_("state").str), Vector("visible", "visible", "hidden"))
      assertEquals(sheets.lift(1).map(_("dimension")), Some(ujson.Str("A1:B1")))
      // metadata only: no counts without --full
      assert(sheets.forall(s => !s.obj.contains("cells")), "light describe carries no cell counts")
      val defined = d("definedNames").arr.toVector
      assertEquals(defined.map(_("name").str), Vector("Total", "Secret"))
      assertEquals(defined.map(_("hidden")), Vector(ujson.False, ujson.True))
      assertEquals(defined.map(_("refersTo").str), Vector("Sheet1!$A$3", "Sheet1!$A$1"))
      assertEquals(d("date1904"), ujson.False)
      assertEquals(run.stderr, "")
    }
  }

  test("--stream describe returns the same metadata-only payload as describe") {
    for
      plain <- CliHarness.run("-f", file("linked.xlsx"), "--json", "describe")
      streamed <- CliHarness.run("-f", file("linked.xlsx"), "--stream", "--json", "describe")
    yield
      assertEquals(streamed.exit, 0, streamed.stderr)
      assertEquals(streamed.stdout, plain.stdout)
      assertEquals(streamed.stderr, "")
  }

  test("describe --full adds per-sheet counts and the workbook's calcPr") {
    CliHarness.run("-f", file("linked.xlsx"), "--json", "describe", "--full").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val d = data(run)
      assertEquals(d.obj.keys.toList, List("sheets", "definedNames", "date1904", "calcPr"))
      assertEquals(d("calcPr"), ujson.Null)
      val sheets = d("sheets").arr.toVector
      val sheet1 = sheets.headOption.getOrElse(fail("no sheets"))
      assertEquals(sheet1("name").str, "Sheet1")
      assertEquals(sheet1("state").str, "visible")
      assertEquals(sheet1("dimension"), ujson.Str("A1:A3"))
      assertEquals(sheet1("cells").num.toInt, 3)
      assertEquals(sheet1("formulas").num.toInt, 1)
      assertEquals(sheet1("uncachedFormulas").num.toInt, 0)
      assertEquals(sheet1("mergedRanges").num.toInt, 1)
      assertEquals(sheet1("comments").num.toInt, 1)
      assertEquals(sheet1("hyperlinks").num.toInt, 0)
      assertEquals(sheet1("freeze"), ujson.Str("B2"))
      assertEquals(sheet1("tabColor"), ujson.Null)
      assertEquals(sheet1("autoFilter"), ujson.Null)
      assertEquals(sheet1("tables"), ujson.Arr())
      assertEquals(sheet1("charts").num.toInt, 0)
      assertEquals(sheet1("pictures").num.toInt, 0)
      assertEquals(sheet1("conditionalFormats").num.toInt, 0)
      assertEquals(sheet1("dataValidations").num.toInt, 0)
      assertEquals(sheet1("hiddenRows").num.toInt, 0)
      assertEquals(sheet1("hiddenCols").num.toInt, 0)
      val hidden = sheets.lift(2).getOrElse(fail("no Hidden sheet"))
      assertEquals(hidden("state").str, "hidden")
      assertEquals(hidden("hiddenRows").num.toInt, 1)
      assertEquals(hidden("cells").num.toInt, 1)
      val sheet2 = sheets.lift(1).getOrElse(fail("no Sheet2"))
      assertEquals(sheet2("formulas").num.toInt, 2)
    }
  }

  test("describe text mode lists sheets, names and the date system; --full adds counts") {
    for
      light <- CliHarness.run("-f", file("linked.xlsx"), "describe")
      full <- CliHarness.run("-f", file("linked.xlsx"), "describe", "--full")
    yield
      assertEquals(light.exit, 0, light.stderr)
      assert(light.stdout.startsWith("Sheets (3):"), light.stdout)
      assert(light.stdout.contains("Hidden"), light.stdout)
      assert(light.stdout.contains("hidden"), light.stdout)
      assert(light.stdout.contains("Defined names (2):"), light.stdout)
      assert(light.stdout.contains("Secret"), light.stdout)
      assert(light.stdout.contains("(hidden)"), light.stdout)
      assert(light.stdout.contains("Date system: 1900"), light.stdout)
      assert(!light.stdout.contains("cells"), s"light describe prints no counts:\n${light.stdout}")
      assertEquals(full.exit, 0, full.stderr)
      assert(full.stdout.contains("cells 3"), full.stdout)
      assert(full.stdout.contains("formulas 1"), full.stdout)
      assert(full.stdout.contains("merged 1"), full.stdout)
      assert(full.stdout.contains("freeze B2"), full.stdout)
  }

  test("--stream describe --full is refused before any read: UNSUPPORTED_IN_STREAM, exit 2") {
    CliHarness.run("-f", file("linked.xlsx"), "--stream", "--json", "describe", "--full").map {
      run =>
        assertEquals(run.exit, 2)
        val e = envelope(run)
        assertEquals(e("ok"), ujson.False)
        assertEquals(e("error")("code"), ujson.Str("UNSUPPORTED_IN_STREAM"))
        assertEquals(e("verb"), ujson.Str("describe"))
        assert(e("error")("hint").str.nonEmpty)
    }
  }

  test("describe is a workbook verb: an existing -s is accepted and ignored") {
    for
      plain <- CliHarness.run("-f", file("linked.xlsx"), "--json", "describe")
      withSheet <- CliHarness.run("-f", file("linked.xlsx"), "-s", "Sheet2", "--json", "describe")
      full <- CliHarness.run(
        "-f",
        file("linked.xlsx"),
        "-s",
        "Sheet2",
        "--json",
        "describe",
        "--full"
      )
    yield
      assertEquals(withSheet.stdout, plain.stdout)
      assertEquals(full.exit, 0, full.stderr)
  }

  test(
    "GH-667: describe --full lists the sheet-scoped print name the reader lifts into PageSetup"
  ) {
    // the dogfood's ops1.json shape: a workbook-scoped name plus `_xlnm.Print_Area` scoped to Data.
    // The full read lifts a modelable Print_Area out of the table into Sheet.pageSetup (GH-259), so
    // --full must re-derive it or the name the light card and `names` both list goes missing
    val path = fixtures().resolve("scoped-names.xlsx")
    val book = TestFixtures
      .simpleBook()
      .withDefinedName("TotalRev", "Data!$B$4")
      .withDefinedName("_xlnm.Print_Area", "Data!$A$1:$B$4", SheetName.unsafe("Data"))
    def scopedNames(run: CliRun): Vector[(String, ujson.Value)] =
      data(run)("definedNames").arr.toVector.map(n => (n("name").str, n("scope")))
    for
      wb <- book.fold(e => IO.raiseError(new Exception(e.toString)), IO.pure)
      _ <- ExcelIO.instance[IO].write(wb, path)
      light <- CliHarness.run("-f", path.toString, "describe")
      full <- CliHarness.run("-f", path.toString, "describe", "--full")
      lightJson <- CliHarness.run("-f", path.toString, "--json", "describe")
      fullJson <- CliHarness.run("-f", path.toString, "--json", "describe", "--full")
      listed <- CliHarness.run("-f", path.toString, "names")
    yield
      val line = "_xlnm.Print_Area  Data!$A$1:$B$4  (Data)"
      assertEquals(light.exit, 0, light.stderr)
      assert(light.stdout.contains("Defined names (2):"), light.stdout)
      assert(light.stdout.contains(line), light.stdout)
      assertEquals(full.exit, 0, full.stderr)
      assert(full.stdout.contains("Defined names (2):"), full.stdout)
      assert(full.stdout.contains(line), full.stdout)
      assertEquals(
        scopedNames(fullJson),
        Vector(("TotalRev", ujson.Null), ("_xlnm.Print_Area", ujson.Str("Data")))
      )
      assertEquals(scopedNames(fullJson), scopedNames(lightJson))
      assertEquals(listed.exit, 0, listed.stderr)
      assert(listed.stdout.contains("_xlnm.Print_Area  Data!$A$1:$B$4 (Data)"), listed.stdout)
  }

  test("describe refuses an unknown -s: SHEET_NOT_FOUND with candidates, exit 3") {
    // (`sheets` and `names` are parsed without --sheet at all: `-s X names` is a USAGE error
    // before any read, pinned by names-with-sheet.golden)
    val verbs = Vector(
      Vector("describe"),
      Vector("describe", "--full"),
      Vector("--stream", "describe")
    )
    verbs.traverse_ { verb =>
      val args = Vector("-f", file("linked.xlsx"), "-s", "Sheet3", "--json") ++ verb
      val name = verb.filterNot(_.startsWith("--")).headOption.getOrElse("")
      CliHarness.run(args*).map { run =>
        assertEquals(run.exit, 3, s"$verb: ${run.stdout}")
        val e = envelope(run)
        assertEquals(e("ok"), ujson.False)
        assertEquals(e("error")("code"), ujson.Str("SHEET_NOT_FOUND"), verb.toString)
        assertEquals(e("verb"), ujson.Str(name))
        assertEquals(names(e("error")("candidates")), Vector("Sheet1", "Sheet2"))
        assert(e("error")("message").str.contains("Sheet3"), e("error")("message").str)
      }
    }
  }

  // ---------------------------------------------------------------------------------------------
  // audit
  // ---------------------------------------------------------------------------------------------

  test("audit --json: one entry per bucket, clean:false, exit 0 without --fail-on-findings") {
    CliHarness.run("-f", file("dirty.xlsx"), "--json", "audit").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val d = data(run)
      assertEquals(
        d.obj.keys.toList,
        List(
          "clean",
          "findings",
          "errorCells",
          "uncachedFormulas",
          "unparseable",
          "volatile",
          "dynamic",
          "cycles",
          "externalRefs",
          "unresolvedReaders",
          "calcPr",
          "iterativeCycles"
        )
      )
      assertEquals(d("clean"), ujson.False)
      // the fixture declares iterative calculation, so its 2-cycle is a note, not a finding
      assertEquals(d("findings").num.toInt, 5)
      assertEquals(
        d("errorCells").arr.toVector.map(e => (e("ref").str, e("error").str)),
        Vector(("Calc!A1", "#DIV/0!"))
      )
      assertEquals(names(d("uncachedFormulas")), Vector("Calc!B1", "Notes!A1"))
      assertEquals(d("unparseable").arr.toVector.map(_("ref").str), Vector("Calc!H1"))
      assert(d("unparseable")(0)("message").str.startsWith("UNSUPPORTED(1)"))
      assertEquals(names(d("volatile")), Vector("Calc!C1"))
      assertEquals(names(d("dynamic")), Vector("Calc!D1"))
      assertEquals(d("cycles"), ujson.Arr())
      assertEquals(
        d("iterativeCycles").arr.toVector.map(c => names(c).toSet),
        Vector(Set("Calc!E1", "Calc!F1"))
      )
      assertEquals(names(d("externalRefs")), Vector("Calc!G1"))
      assertEquals(names(d("unresolvedReaders")), Vector("Calc!I1"))
      assertEquals(d("calcPr")("iterativeCalculation"), ujson.True)
      assertEquals(d("calcPr")("maxIterations").num.toInt, 100)
      assertEquals(run.stderr, "")
    }
  }

  test("audit --fail-on-findings: exit 1 with AUDIT_FINDINGS and the report kept as data") {
    CliHarness.run("-f", file("dirty.xlsx"), "--json", "audit", "--fail-on-findings").map { run =>
      assertEquals(run.exit, 1)
      val e = envelope(run)
      assertEquals(e("ok"), ujson.False)
      assertEquals(e("verb"), ujson.Str("audit"))
      assertEquals(e("error")("code"), ujson.Str("AUDIT_FINDINGS"))
      assertEquals(e("data")("clean"), ujson.False)
      assertEquals(e("data")("findings").num.toInt, 5)
      assertEquals(run.stderr, "", "findings keep their report as data: no Error: line, as in text")
    }
  }

  test("audit --fail-on-findings in text mode prints the report, not an Error: line") {
    CliHarness.run("-f", file("dirty.xlsx"), "audit", "--fail-on-findings").map { run =>
      assertEquals(run.exit, 1)
      assert(run.stdout.startsWith("Audit: 5 findings"), run.stdout)
      assertEquals(run.stderr, "")
    }
  }

  test("audit --fail-on-findings on a clean book exits 0 and says so") {
    for
      text <- CliHarness.run("-f", file("simple.xlsx"), "audit", "--fail-on-findings")
      json <- CliHarness.run("-f", file("simple.xlsx"), "--json", "audit", "--fail-on-findings")
    yield
      assertEquals(text.exit, 0, text.stderr)
      assertEquals(text.stdout, "Audit: clean\n")
      assertEquals(json.exit, 0, json.stderr)
      val d = data(json)
      assertEquals(d("clean"), ujson.True)
      assertEquals(d("findings").num.toInt, 0)
      assertEquals(d("calcPr"), ujson.Null)
  }

  test("audit text mode prints one section per non-empty bucket, notes after findings") {
    CliHarness.run("-f", file("dirty.xlsx"), "audit").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val out = run.stdout
      val sections = Vector(
        "Error values (1):",
        "Uncached formulas (2):",
        "Unparseable formulas (1):",
        "Unresolved names (1):",
        "Cycles (iterative calculation on, not a finding) (1):",
        "Volatile (1):",
        "Dynamic (1):",
        "External references (1):",
        "Calculation:"
      )
      val positions = sections.map(s => (s, out.indexOf(s)))
      positions.foreach((s, i) => assert(i >= 0, s"missing section '$s' in:\n$out"))
      assertEquals(positions.map(_._2), positions.map(_._2).sorted, "sections in a fixed order")
      assert(out.contains("Calc!A1  #DIV/0!"), out)
      assert(out.contains("Calc!E1, Calc!F1"), out)
      assert(out.contains("iterative"), out)
    }
  }

  test("audit -s restricts the buckets to one sheet") {
    CliHarness.run("-f", file("dirty.xlsx"), "-s", "Notes", "--json", "audit").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val d = data(run)
      assertEquals(names(d("uncachedFormulas")), Vector("Notes!A1"))
      assertEquals(d("errorCells"), ujson.Arr())
      assertEquals(d("cycles"), ujson.Arr())
      assertEquals(d("findings").num.toInt, 1)
    }
  }

  test("--stream audit is refused: UNSUPPORTED_IN_STREAM, exit 2") {
    CliHarness.run("-f", file("dirty.xlsx"), "--stream", "--json", "audit").map { run =>
      assertEquals(run.exit, 2)
      assertEquals(envelope(run)("error")("code"), ujson.Str("UNSUPPORTED_IN_STREAM"))
    }
  }

  // ---------------------------------------------------------------------------------------------
  // deps
  // ---------------------------------------------------------------------------------------------

  test(
    "deps Sheet2!A1 --json lists Sheet1!A1 as a precedent with its value, Sheet2!B1 as a dependent"
  ) {
    CliHarness.run("-f", file("linked.xlsx"), "--json", "deps", "Sheet2!A1").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val d = data(run)
      assertEquals(
        d.obj.keys.toList,
        List("ref", "formula", "value", "direction", "depth", "expand", "precedents", "dependents")
      )
      assertEquals(d("ref"), ujson.Str("Sheet2!A1"))
      assertEquals(d("formula"), ujson.Str("=Sheet1!A1*2"))
      assertEquals(d("value"), ujson.Num(10))
      assertEquals(d("direction"), ujson.Str("both"))
      assertEquals(d("depth"), ujson.Num(1))
      assertEquals(d("expand"), ujson.False)
      val precedents = d("precedents").arr.toVector
      assertEquals(
        precedents.map(_.obj.keys.toList),
        Vector(List("ref", "kind", "depth", "formula", "value"))
      )
      assertEquals(precedents.map(_("kind").str), Vector("cell"))
      assertEquals(precedents.map(_("ref").str), Vector("Sheet1!A1"))
      assertEquals(precedents.map(_("depth").num.toInt), Vector(1))
      assertEquals(precedents.map(_("formula")), Vector(ujson.Null))
      assertEquals(precedents.map(_("value")), Vector(ujson.Num(5)))
      val dependents = d("dependents").arr.toVector
      assertEquals(dependents.map(_("ref").str), Vector("Sheet2!B1"))
      assertEquals(dependents.map(_("kind").str), Vector("cell"))
      assertEquals(dependents.map(_("formula")), Vector(ujson.Str("=A1+1")))
      assertEquals(dependents.map(_("value")), Vector(ujson.Num(11)))
    }
  }

  test(
    "deps --direction dependents --depth all walks the whole cone in layers; precedents is null"
  ) {
    CliHarness
      .run(
        "-f",
        file("linked.xlsx"),
        "-s",
        "Sheet1",
        "--json",
        "deps",
        "A1",
        "--direction",
        "dependents",
        "--depth",
        "all"
      )
      .map { run =>
        assertEquals(run.exit, 0, run.stderr)
        val d = data(run)
        assertEquals(d("ref"), ujson.Str("Sheet1!A1"))
        assertEquals(d("formula"), ujson.Null)
        assertEquals(d("value"), ujson.Num(5))
        assertEquals(d("direction"), ujson.Str("dependents"))
        assertEquals(d("depth"), ujson.Str("all"))
        assertEquals(d("precedents"), ujson.Null)
        val dependents = d("dependents").arr.toVector
        assertEquals(
          dependents.map(n => (n("ref").str, n("depth").num.toInt)),
          Vector(("Sheet1!A3", 1), ("Sheet2!A1", 1), ("Sheet2!B1", 2))
        )
      }
  }

  test("deps --depth 2 --direction precedents stops after two layers") {
    CliHarness
      .run(
        "-f",
        file("linked.xlsx"),
        "--json",
        "deps",
        "Sheet2!B1",
        "--direction",
        "precedents",
        "--depth",
        "2"
      )
      .map { run =>
        assertEquals(run.exit, 0, run.stderr)
        val d = data(run)
        assertEquals(d("depth"), ujson.Num(2))
        assertEquals(d("dependents"), ujson.Null)
        assertEquals(
          d("precedents").arr.toVector.map(n => (n("ref").str, n("depth").num.toInt)),
          Vector(("Sheet2!A1", 1), ("Sheet1!A1", 2))
        )
      }
  }

  test("deps text mode: the cell, then one line per node with its depth and value") {
    CliHarness.run("-f", file("linked.xlsx"), "deps", "Sheet2!A1").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val lines = run.stdout.linesIterator.toVector
      assertEquals(lines.headOption, Some("Cell: Sheet2!A1"))
      assert(lines.contains("Formula: =Sheet1!A1*2"), run.stdout)
      assert(lines.contains("Value: 10"), run.stdout)
      assert(lines.contains("Precedents (depth 1): 1"), run.stdout)
      assert(lines.contains("Dependents (depth 1): 1"), run.stdout)
      assert(lines.exists(_.matches("\\s+1\\s+Sheet1!A1\\s+5")), run.stdout)
      assert(
        lines.exists(l => l.contains("Sheet2!B1") && l.contains("=A1+1") && l.contains("11")),
        run.stdout
      )
    }
  }

  test("deps on a leaf reports (none) in text mode and empty arrays in JSON") {
    for
      text <- CliHarness.run("-f", file("linked.xlsx"), "deps", "Sheet1!A2")
      json <- CliHarness.run(
        "-f",
        file("linked.xlsx"),
        "--json",
        "deps",
        "Sheet2!B1",
        "--direction",
        "dependents"
      )
    yield
      assertEquals(text.exit, 0, text.stderr)
      assert(text.stdout.contains("Precedents (depth 1): 0\n  (none)"), text.stdout)
      assertEquals(json.exit, 0, json.stderr)
      assertEquals(data(json)("dependents"), ujson.Arr())
  }

  test("deps follows the one sheet rule: unqualified ref on a multi-sheet book is SHEET_REQUIRED") {
    CliHarness.run("-f", file("linked.xlsx"), "--json", "deps", "A1").map { run =>
      assertEquals(run.exit, 3)
      val e = envelope(run)
      assertEquals(e("error")("code"), ujson.Str("SHEET_REQUIRED"))
      assertEquals(e("verb"), ujson.Str("deps"))
    }
  }

  test("deps rejects a range, a bad direction and a negative or non-numeric depth; 0 means all") {
    for
      range <- CliHarness.run("-f", file("linked.xlsx"), "--json", "deps", "Sheet1!A1:A2")
      direction <- CliHarness.run(
        "-f",
        file("linked.xlsx"),
        "--json",
        "deps",
        "Sheet1!A1",
        "--direction",
        "sideways"
      )
      zero <- CliHarness.run(
        "-f",
        file("linked.xlsx"),
        "--json",
        "deps",
        "Sheet1!A1",
        "--depth",
        "0"
      )
      all <- CliHarness.run(
        "-f",
        file("linked.xlsx"),
        "--json",
        "deps",
        "Sheet1!A1",
        "--depth",
        "all"
      )
      negative <- CliHarness.run(
        "-f",
        file("linked.xlsx"),
        "--json",
        "deps",
        "Sheet1!A1",
        "--depth=-1"
      )
      garbage <- CliHarness.run(
        "-f",
        file("linked.xlsx"),
        "--json",
        "deps",
        "Sheet1!A1",
        "--depth",
        "many"
      )
    yield
      assertEquals(range.exit, 3, range.stderr)
      assertEquals(envelope(range)("error")("code"), ujson.Str("INVALID_REFERENCE"))
      assertEquals(direction.exit, 2, direction.stderr)
      assertEquals(envelope(direction)("error")("code"), ujson.Str("USAGE"))
      // `--depth 0` is the plan's spelling of `all`: same payload, byte for byte
      assertEquals(zero.exit, 0, zero.stderr)
      assertEquals(data(zero)("depth"), ujson.Str("all"))
      assertEquals(zero.stdout, all.stdout)
      assertEquals(negative.exit, 2, negative.stderr)
      assertEquals(garbage.exit, 2, garbage.stderr)
  }

  // ---------------------------------------------------------------------------------------------
  // deps: a range is one precedent node
  // ---------------------------------------------------------------------------------------------

  private def refsAndKinds(nodes: ujson.Value): Vector[(String, String)] =
    nodes.arr.toVector.map(n => (n("ref").str, n("kind").str))

  test("the ranges fixture really carries a style-only blank inside Data!B:B") {
    ExcelIO.instance[IO].read(fixtures().resolve("ranges.xlsx")).map { wb =>
      val data = wb.sheets.find(_.name == SheetName.unsafe("Data"))
      val b6 = data.flatMap(_.cells.get(com.tjclp.xl.addressing.ARef.from1(2, 6)))
      assertEquals(b6.map(_.value), Some(com.tjclp.xl.cells.CellValue.Empty))
      assert(b6.exists(_.styleId.isDefined), s"B6 must be a styled record: $b6")
    }
  }

  test("deps --json collapses each range the formula reads into ONE node with its counts") {
    CliHarness
      .run("-f", file("ranges.xlsx"), "--json", "deps", "Summary!B1", "--direction", "precedents")
      .map { run =>
        assertEquals(run.exit, 0, run.stderr)
        val d = data(run)
        assertEquals(d("value"), ujson.Num(170))
        val nodes = d("precedents").arr.toVector
        assertEquals(
          refsAndKinds(d("precedents")),
          Vector(
            ("Data!A:A", "range"),
            ("Data!B:B", "range"),
            ("Data!B2", "cell"),
            ("Data!C2:C4", "range"),
            ("Summary!A1", "cell")
          )
        )
        val ranges = nodes.filter(_("kind").str == "range")
        assertEquals(
          ranges.map(_.obj.keys.toList).distinct,
          Vector(List("ref", "kind", "depth", "formula", "value", "occupied", "formulas"))
        )
        // occupied counts value-holding cells only: B:B's style-only B6 is not one
        assertEquals(
          ranges.map(n => (n("ref").str, n("occupied").num.toInt, n("formulas").num.toInt)),
          Vector(("Data!A:A", 4, 0), ("Data!B:B", 4, 0), ("Data!C2:C4", 3, 3))
        )
        assert(ranges.forall(n => n("formula") == ujson.Null && n("value") == ujson.Null))
        assert(ranges.forall(_("depth") == ujson.Num(1)))
        val cells = nodes.filter(_("kind").str == "cell")
        assertEquals(
          cells.map(_.obj.keys.toList).distinct,
          Vector(List("ref", "kind", "depth", "formula", "value"))
        )
        assertEquals(cells.map(_("value")), Vector(ujson.Num(10), ujson.Str("North")))
      }
  }

  test("deps --depth 2 continues through a range's formulas without relisting its cells") {
    CliHarness
      .run(
        "-f",
        file("ranges.xlsx"),
        "--json",
        "deps",
        "Summary!B1",
        "--direction",
        "precedents",
        "--depth",
        "2"
      )
      .map { run =>
        assertEquals(run.exit, 0, run.stderr)
        val nodes = data(run)("precedents").arr.toVector
        // C2:C4 read B2..B4 (covered by Data!B:B at depth 1) and E1: only E1 is new
        assertEquals(
          nodes.filter(_("depth") == ujson.Num(2)).map(n => (n("ref").str, n("kind").str)),
          Vector(("Data!E1", "cell"))
        )
        assert(!nodes.exists(_("ref").str == "Data!B3"), nodes.toString)
      }
  }

  test("deps --expand lists every occupied cell one by one and echoes expand: true") {
    CliHarness
      .run(
        "-f",
        file("ranges.xlsx"),
        "--json",
        "deps",
        "Summary!B1",
        "--direction",
        "precedents",
        "--depth",
        "2",
        "--expand"
      )
      .map { run =>
        assertEquals(run.exit, 0, run.stderr)
        val d = data(run)
        assertEquals(d("expand"), ujson.True)
        val nodes = d("precedents").arr.toVector
        assert(nodes.forall(_("kind").str == "cell"), nodes.toString)
        assertEquals(
          nodes.map(n => (n("ref").str, n("depth").num.toInt)),
          Vector(
            "Data!A1",
            "Data!B1",
            "Data!A2",
            "Data!B2",
            "Data!C2",
            "Data!A3",
            "Data!B3",
            "Data!C3",
            "Data!A4",
            "Data!B4",
            "Data!C4",
            "Summary!A1"
          ).map(_ -> 1) :+ ("Data!E1" -> 2)
        )
      }
  }

  test("--expand leaves dependents alone and parses before the ref too") {
    def dependents(extra: String*) =
      CliHarness.run(
        List("-f", file("ranges.xlsx"), "--json", "deps") ++ extra ++
          List("--direction", "dependents"),
        ""
      )
    for
      plain <- dependents("Data!B3")
      expanded <- dependents("--expand", "Data!B3")
    yield
      assertEquals(plain.exit, 0, plain.stderr)
      assertEquals(expanded.exit, 0, expanded.stderr)
      assertEquals(data(expanded)("expand"), ujson.True)
      assertEquals(data(expanded)("dependents"), data(plain)("dependents"))
      assertEquals(
        refsAndKinds(data(plain)("dependents")),
        Vector(("Data!C3", "cell"), ("Summary!B1", "cell"))
      )
  }

  test("deps text prints a range as one line with its occupied cells and formulas") {
    for
      text <- CliHarness.run("-f", file("ranges.xlsx"), "deps", "Summary!B1")
      single <- CliHarness.run("-f", file("ranges.xlsx"), "deps", "Summary!C1")
      empty <- CliHarness.run("-f", file("gaps.xlsx"), "deps", "Data!C1")
    yield
      assertEquals(text.exit, 0, text.stderr)
      val lines = text.stdout.linesIterator.toVector
      assert(lines.contains("Precedents (depth 1): 5"), text.stdout)
      assert(lines.contains("  1  Data!A:A  range, 4 occupied cells"), text.stdout)
      assert(lines.contains("  1  Data!C2:C4  range, 3 occupied cells (3 formulas)"), text.stdout)
      assert(lines.contains("  1  Data!B2  10"), text.stdout)
      assertEquals(single.exit, 0, single.stderr)
      assert(
        single.stdout.contains("  1  Data!C4:E4  range, 1 occupied cell (1 formula)\n"),
        single.stdout
      )
      // an empty range explains a zero: it is listed, never dropped
      assertEquals(empty.exit, 0, empty.stderr)
      assert(
        empty.stdout.contains(
          "Precedents (depth 1): 1\n  1  Missing!A1:A3  range, 0 occupied cells"
        ),
        empty.stdout
      )
  }

  test("Weaver's whole-column SUMIFS: three precedents, not every cell of two columns") {
    for
      direct <- CliHarness.run("-f", file("sumifs.xlsx"), "-s", "Summary", "deps", "B1")
      json <- CliHarness.run("-f", file("sumifs.xlsx"), "-s", "Summary", "--json", "deps", "B1")
      deep <- CliHarness.run(
        "-f",
        file("sumifs.xlsx"),
        "-s",
        "Summary",
        "deps",
        "B2",
        "--direction",
        "precedents",
        "--depth",
        "2"
      )
      cell <- CliHarness.run("-f", file("sumifs.xlsx"), "-s", "Summary", "cell", "B1")
    yield
      assertEquals(direct.exit, 0, direct.stderr)
      assert(
        direct.stdout.contains(
          "Precedents (depth 1): 3\n" +
            "  1  Data!A:A  range, 401 occupied cells\n" +
            "  1  Data!B:B  range, 401 occupied cells\n" +
            "  1  Summary!A1  \"North\"\n"
        ),
        direct.stdout
      )
      assertEquals(
        refsAndKinds(data(json)("precedents")),
        Vector(("Data!A:A", "range"), ("Data!B:B", "range"), ("Summary!A1", "cell"))
      )
      assertEquals(deep.exit, 0, deep.stderr)
      val deepLines = deep.stdout.linesIterator.toVector
      assert(
        deepLines.contains("  1  Data!C:C  range, 400 occupied cells (400 formulas)"),
        deep.stdout
      )
      // depth 2: the two columns SUMIFS reads stay one line each; what remains is the 400 cells
      // the filled-down C formulas name one by one (was 1,204 nodes, every cell of A, B and C)
      assert(deepLines.contains("  2  Data!A:A  range, 401 occupied cells"), deep.stdout)
      assert(deepLines.contains("  2  Data!B:B  range, 401 occupied cells"), deep.stdout)
      assert(!deepLines.exists(_.startsWith("  2  Data!A2 ")), deep.stdout)
      assert(deepLines.contains("Precedents (depth 2): 405"), deep.stdout)
      assertEquals(cell.exit, 0, cell.stderr)
      assert(
        cell.stdout.endsWith("Dependencies: Data!A:A, Data!B:B, A1\nDependents: B2\n"),
        cell.stdout
      )
  }

  test("--stream deps is refused: UNSUPPORTED_IN_STREAM, exit 2") {
    CliHarness.run("-f", file("linked.xlsx"), "--stream", "--json", "deps", "Sheet2!A1").map {
      run =>
        assertEquals(run.exit, 2)
        assertEquals(envelope(run)("error")("code"), ujson.Str("UNSUPPORTED_IN_STREAM"))
    }
  }

  // ---------------------------------------------------------------------------------------------
  // cell lists what a formula reads as declared, a range as one entry
  // ---------------------------------------------------------------------------------------------

  test("cell lists a formula's ranges as declared: an empty stretch or a missing sheet included") {
    // Since 0.24.0 a range is one Dependencies entry, spelled as deps spells it (before: its
    // occupied cells one by one, which `deps --expand` still lists); Dependents are unchanged.
    for
      gaps <- CliHarness.run("-f", file("gaps.xlsx"), "cell", "Data!B1")
      absent <- CliHarness.run("-f", file("gaps.xlsx"), "cell", "Data!C1")
      constant <- CliHarness.run("-f", file("gaps.xlsx"), "cell", "Data!A3")
      empty <- CliHarness.run("-f", file("gaps.xlsx"), "cell", "Data!A2")
      declared <- CliHarness.run("-f", file("ranges.xlsx"), "cell", "Summary!B1")
      json <- CliHarness.run("-f", file("ranges.xlsx"), "--json", "cell", "Summary!B1")
    yield
      assertEquals(gaps.exit, 0, gaps.stderr)
      assert(gaps.stdout.contains("Formula: =SUM(A1:A5)\n"), gaps.stdout)
      assert(gaps.stdout.endsWith("Dependencies: A1:A5\nDependents: (none)\n"), gaps.stdout)
      assertEquals(absent.exit, 0, absent.stderr)
      assert(absent.stdout.contains("Formula: =SUM(Missing!A1:A3)\n"), absent.stdout)
      assert(
        absent.stdout.endsWith("Dependencies: Missing!A1:A3\nDependents: (none)\n"),
        absent.stdout
      )
      // an occupied cell inside the range is read by B1; so is the empty one (symbolic index)
      assertEquals(constant.exit, 0, constant.stderr)
      assert(constant.stdout.endsWith("Dependencies: (none)\nDependents: B1\n"), constant.stdout)
      assertEquals(empty.exit, 0, empty.stderr)
      assert(empty.stdout.endsWith("Dependencies: (none)\nDependents: B1\n"), empty.stdout)
      // same-sheet entries unqualified, cross-sheet ones qualified, in deps' node order
      assertEquals(declared.exit, 0, declared.stderr)
      assert(
        declared.stdout.endsWith(
          "Dependencies: Data!A:A, Data!B:B, Data!B2, Data!C2:C4, A1\nDependents: (none)\n"
        ),
        declared.stdout
      )
      assertEquals(
        names(data(json)("dependencies")),
        Vector("Data!A:A", "Data!B:B", "Data!B2", "Data!C2:C4", "A1")
      )
  }

  test("cell: a formula's range input is one entry; a constant's range reader is unchanged") {
    for
      formula <- CliHarness.run("-f", file("simple.xlsx"), "cell", "Data!B4")
      constant <- CliHarness.run("-f", file("simple.xlsx"), "cell", "Data!B1")
      cross <- CliHarness.run("-f", file("linked.xlsx"), "cell", "Sheet1!A1")
    yield
      assertEquals(formula.exit, 0, formula.stderr)
      assertEquals(
        formula.stdout,
        "Cell: B4\nType: formula\nFormula: =SUM(B1:B3)\nCached: 42.5\nDependencies: B1:B3\nDependents: (none)\n"
      )
      assertEquals(constant.exit, 0, constant.stderr)
      assert(constant.stdout.endsWith("Dependencies: (none)\nDependents: B4\n"), constant.stdout)
      assertEquals(cross.exit, 0, cross.stderr)
      // same-sheet dependents are unqualified, cross-sheet ones carry the sheet, sorted by text
      assert(
        cross.stdout.endsWith("Dependencies: (none)\nDependents: A3, Sheet2!A1\n"),
        cross.stdout
      )
  }

  test("cell quotes a cross-sheet qualifier the way deps does: 'On-Premise'!G9 (GH-609)") {
    // A hyphenated sheet name must be quoted in a formula; `cell` printed `On-Premise!G9` (not a
    // reference an agent can paste into `putf`/`eval`) while `deps` printed `'On-Premise'!G9`.
    for
      text <- CliHarness.run("-f", file("qualified.xlsx"), "cell", "Summary!G9")
      json <- CliHarness.run("-f", file("qualified.xlsx"), "--json", "cell", "Summary!G9")
      deps <- CliHarness.run("-f", file("qualified.xlsx"), "deps", "Summary!G9")
      reverse <- CliHarness.run(
        "-f",
        file("qualified.xlsx"),
        "-s",
        "On-Premise",
        "--json",
        "cell",
        "G9"
      )
      plain <- CliHarness.run("-f", file("linked.xlsx"), "--json", "cell", "Sheet1!A1")
      search <- CliHarness.run("-f", file("qualified.xlsx"), "search", "Servers")
    yield
      // `search` spells its Ref column through the same printer
      assertEquals(search.exit, 0, search.stderr)
      assert(search.stdout.contains("| 'On-Premise'!A1 |"), search.stdout)
      assertEquals(text.exit, 0, text.stderr)
      assert(
        text.stdout.endsWith("Dependencies: 'On-Premise'!G9\nDependents: (none)\n"),
        text.stdout
      )
      assertEquals(deps.exit, 0, deps.stderr)
      assert(deps.stdout.contains("'On-Premise'!G9"), deps.stdout)
      assertEquals(names(data(json)("dependencies")), Vector("'On-Premise'!G9"))
      // the reverse edge: the reader lives on a sheet whose name needs no quoting
      assertEquals(names(data(reverse)("dependents")), Vector("Summary!G9"))
      // a plain sheet name stays bare, same-sheet refs stay unqualified
      assertEquals(names(data(plain)("dependents")), Vector("A3", "Sheet2!A1"))
  }

  test("Argv.verbs lists the three inspection verbs in parser order") {
    IO {
      val verbs = Argv.verbs
      assertEquals(
        verbs.slice(verbs.indexOf("filter") + 1, verbs.indexOf("filter") + 4),
        Vector("describe", "audit", "deps")
      )
    }
  }
