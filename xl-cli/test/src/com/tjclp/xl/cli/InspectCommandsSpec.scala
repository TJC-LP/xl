package com.tjclp.xl.cli

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.cli.contract.{Argv, CliHarness, CliRun, EnvelopeSchema, TestFixtures}

/**
 * The inspection verbs of ADR-017 §2.10 through the in-process harness: `describe [--full]`,
 * `audit [--fail-on-findings]` and `deps <ref> [--direction] [--depth]`, each typed under `--json`
 * and readable otherwise; plus the pin that `cell`'s text is unchanged by its move to the bounded
 * graph. Byte-level renderings are pinned by the goldens (`describe*`, `audit*`, `deps*`); this
 * suite pins the shapes and the exit-code contract.
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

  test("describe is a workbook verb: -s is accepted and ignored") {
    for
      plain <- CliHarness.run("-f", file("linked.xlsx"), "--json", "describe")
      withSheet <- CliHarness.run("-f", file("linked.xlsx"), "-s", "Sheet2", "--json", "describe")
    yield assertEquals(withSheet.stdout, plain.stdout)
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
          "calcPr"
        )
      )
      assertEquals(d("clean"), ujson.False)
      assertEquals(d("findings").num.toInt, 6)
      assertEquals(
        d("errorCells").arr.toVector.map(e => (e("ref").str, e("error").str)),
        Vector(("Calc!A1", "#DIV/0!"))
      )
      assertEquals(names(d("uncachedFormulas")), Vector("Calc!B1", "Notes!A1"))
      assertEquals(d("unparseable").arr.toVector.map(_("ref").str), Vector("Calc!H1"))
      assert(d("unparseable")(0)("message").str.startsWith("UNSUPPORTED(1)"))
      assertEquals(names(d("volatile")), Vector("Calc!C1"))
      assertEquals(names(d("dynamic")), Vector("Calc!D1"))
      assertEquals(
        d("cycles").arr.toVector.map(c => names(c).toSet),
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
      assertEquals(e("data")("findings").num.toInt, 6)
      assert(run.stderr.startsWith("Error: "), run.stderr)
    }
  }

  test("audit --fail-on-findings in text mode prints the report, not an Error: line") {
    CliHarness.run("-f", file("dirty.xlsx"), "audit", "--fail-on-findings").map { run =>
      assertEquals(run.exit, 1)
      assert(run.stdout.startsWith("Audit: 6 findings"), run.stdout)
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
        "Cycles (1):",
        "Unresolved names (1):",
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
        List("ref", "formula", "value", "direction", "depth", "precedents", "dependents")
      )
      assertEquals(d("ref"), ujson.Str("Sheet2!A1"))
      assertEquals(d("formula"), ujson.Str("=Sheet1!A1*2"))
      assertEquals(d("value"), ujson.Num(10))
      assertEquals(d("direction"), ujson.Str("both"))
      assertEquals(d("depth"), ujson.Num(1))
      val precedents = d("precedents").arr.toVector
      assertEquals(precedents.map(_("ref").str), Vector("Sheet1!A1"))
      assertEquals(precedents.map(_("depth").num.toInt), Vector(1))
      assertEquals(precedents.map(_("formula")), Vector(ujson.Null))
      assertEquals(precedents.map(_("value")), Vector(ujson.Num(5)))
      val dependents = d("dependents").arr.toVector
      assertEquals(dependents.map(_("ref").str), Vector("Sheet2!B1"))
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

  test("--stream deps is refused: UNSUPPORTED_IN_STREAM, exit 2") {
    CliHarness.run("-f", file("linked.xlsx"), "--stream", "--json", "deps", "Sheet2!A1").map {
      run =>
        assertEquals(run.exit, 2)
        assertEquals(envelope(run)("error")("code"), ujson.Str("UNSUPPORTED_IN_STREAM"))
    }
  }

  // ---------------------------------------------------------------------------------------------
  // cell keeps its text while moving to the bounded graph
  // ---------------------------------------------------------------------------------------------

  test("cell lists a range's OCCUPIED cells: gaps and absent-sheet ranges are not Dependencies") {
    // The declared behaviour change from the unbounded fromWorkbook expansion (ADR-017 §2.10):
    // the old listing named every cell of the range (SUM(A:A) printed 1,048,576 entries) and
    // three refs on a sheet the workbook does not have; Dependents lines are unchanged.
    for
      gaps <- CliHarness.run("-f", file("gaps.xlsx"), "cell", "Data!B1")
      absent <- CliHarness.run("-f", file("gaps.xlsx"), "cell", "Data!C1")
      constant <- CliHarness.run("-f", file("gaps.xlsx"), "cell", "Data!A3")
      empty <- CliHarness.run("-f", file("gaps.xlsx"), "cell", "Data!A2")
    yield
      assertEquals(gaps.exit, 0, gaps.stderr)
      assert(gaps.stdout.contains("Formula: =SUM(A1:A5)\n"), gaps.stdout)
      assert(gaps.stdout.endsWith("Dependencies: A1, A3\nDependents: (none)\n"), gaps.stdout)
      assertEquals(absent.exit, 0, absent.stderr)
      assert(absent.stdout.contains("Formula: =SUM(Missing!A1:A3)\n"), absent.stdout)
      assert(absent.stdout.endsWith("Dependencies: (none)\nDependents: (none)\n"), absent.stdout)
      // an occupied cell inside the range is read by B1; so is the empty one (symbolic index)
      assertEquals(constant.exit, 0, constant.stderr)
      assert(constant.stdout.endsWith("Dependencies: (none)\nDependents: B1\n"), constant.stdout)
      assertEquals(empty.exit, 0, empty.stderr)
      assert(empty.stdout.endsWith("Dependencies: (none)\nDependents: B1\n"), empty.stdout)
  }

  test("cell text output is unchanged: a formula's range inputs and a constant's range reader") {
    for
      formula <- CliHarness.run("-f", file("simple.xlsx"), "cell", "Data!B4")
      constant <- CliHarness.run("-f", file("simple.xlsx"), "cell", "Data!B1")
      cross <- CliHarness.run("-f", file("linked.xlsx"), "cell", "Sheet1!A1")
    yield
      assertEquals(formula.exit, 0, formula.stderr)
      assertEquals(
        formula.stdout,
        "Cell: B4\nType: formula\nFormula: =SUM(B1:B3)\nCached: 42.5\nDependencies: B1, B2, B3\nDependents: (none)\n"
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

  test("Argv.verbs lists the three inspection verbs in parser order") {
    IO {
      val verbs = Argv.verbs
      assertEquals(
        verbs.slice(verbs.indexOf("filter") + 1, verbs.indexOf("filter") + 4),
        Vector("describe", "audit", "deps")
      )
    }
  }
