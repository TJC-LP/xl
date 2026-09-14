package com.tjclp.xl.cli.batch

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.cli.contract.{CliHarness, CliRun, TestFixtures}

/**
 * GH-462: `define-name` / `remove-name`, the batch twins of `name add` / `name rm` (`scope` is the
 * verb's `-s`). The field shape is EditSchema's (`name`, `refersTo`, `scope`); refusals are the
 * verb's, at the op index; matching is case-insensitive (GH-538); a `sheet` key is not a scope; the
 * ops neither stream nor recalculate.
 */
class BatchNameOpsSpec extends CatsEffectSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "batch-name-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  /** `simple.xlsx` (sheets `Data`, `Summary`) plus the workbook-scoped `Total` → `Data!$B$4`. */
  private def named: String = fixtures().resolve("named.xlsx").toString

  private def fresh(name: String): Path = fixtures().resolve(name)

  private def batch(input: String, out: Path, json: String, flags: String*): IO[CliRun] =
    CliHarness.run(
      List("-f", input, "-o", out.toString) ++ flags.toList ++ List("batch", "-"),
      json
    )

  /** `names --json` of a written file as `(name, refersTo, scope)` — `data.names` (GH-618). */
  private def names(path: Path): IO[Vector[(String, String, ujson.Value)]] =
    CliHarness.run("-f", path.toString, "--json", "names").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      ujson
        .read(run.stdout)("data")("names")
        .arr
        .map(n => (n("name").str, n("refersTo").str, n("scope")))
        .toVector
    }

  test("define-name adds a workbook-scoped name; remove-name removes one case-insensitively") {
    val out = fresh("name-ops-global.xlsx")
    val json =
      """[{"op":"define-name","name":"Tax","refersTo":"Data!$A$1"},{"op":"remove-name","name":"TOTAL"}]"""
    for
      run <- batch(named, out, json)
      listed <- names(out)
    yield
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(listed, Vector(("Tax", "Data!$A$1", ujson.Null)))
  }

  test("scope authors a sheet-scoped name (kebab refers-to accepted); a scoped remove drops it") {
    val added = fresh("name-ops-scoped.xlsx")
    val removed = fresh("name-ops-scoped-rm.xlsx")
    for
      add <- batch(
        named,
        added,
        """[{"op":"define-name","name":"Local","refers-to":"Data!$B$2","scope":"Data"}]"""
      )
      afterAdd <- names(added)
      // the scope is matched case-insensitively, like every sheet lookup; so is the name
      rm <- batch(
        added.toString,
        removed,
        """[{"op":"remove-name","name":"local","scope":"data"}]"""
      )
      afterRm <- names(removed)
    yield
      assertEquals(add.exit, 0, add.stderr)
      assertEquals(
        afterAdd,
        Vector(("Total", "Data!$B$4", ujson.Null), ("Local", "Data!$B$2", ujson.Str("Data")))
      )
      assertEquals(rm.exit, 0, rm.stderr)
      assertEquals(afterRm, Vector(("Total", "Data!$B$4", ujson.Null)))
  }

  test("remove-name with scope removes a Print_Area the read lifted into page setup") {
    // GH-462: after a read the print area lives in pageSetup.printArea, not in the table; the
    // twin of `name rm -s` must still remove it from the file `define-name` produced.
    val added = fresh("name-ops-print-add.xlsx")
    val removed = fresh("name-ops-print-rm.xlsx")
    for
      add <- batch(
        named,
        added,
        """[{"op":"define-name","name":"_xlnm.Print_Area","refersTo":"Data!$A$1:$B$2","scope":"Data"}]"""
      )
      afterAdd <- names(added)
      rm <- batch(
        added.toString,
        removed,
        """[{"op":"remove-name","name":"_xlnm.print_area","scope":"data"}]"""
      )
      afterRm <- names(removed)
    yield
      assertEquals(add.exit, 0, add.stderr)
      assertEquals(
        afterAdd,
        Vector(
          ("Total", "Data!$B$4", ujson.Null),
          ("_xlnm.Print_Area", "Data!$A$1:$B$2", ujson.Str("Data"))
        )
      )
      assertEquals(rm.exit, 0, rm.stderr)
      assertEquals(afterRm, Vector(("Total", "Data!$B$4", ujson.Null)))
  }

  test("remove-name of an unknown name fails at the op index with the verb's did-you-mean") {
    val out = fresh("name-ops-missing.xlsx")
    val json =
      """[{"op":"define-name","name":"Tax","refersTo":"Data!$A$1"},{"op":"remove-name","name":"Totl"}]"""
    batch(named, out, json, "--json").map { run =>
      assertEquals(run.exit, 3, run.stderr)
      val error = ujson.read(run.stdout)("error")
      assertEquals(error("code"), ujson.Str("BATCH_OP_FAILED"))
      assert(
        error("message").str.contains("Named range 'Totl' not found. Available: Total"),
        error("message").str
      )
      assertEquals(error("candidates"), ujson.Arr(ujson.Str("Total")))
      assertEquals(error("location")("opIndex"), ujson.Num(2))
      assert(!Files.exists(out), "all-or-nothing: a failing op writes nothing")
    }
  }

  test("define-name with an unknown scope is the verb's SHEET_NOT_FOUND at the op index") {
    val out = fresh("name-ops-bad-scope.xlsx")
    val json = """[{"op":"define-name","name":"Local","refersTo":"Data!$A$1","scope":"Dat"}]"""
    batch(named, out, json, "--json").map { run =>
      assertEquals(run.exit, 3, run.stderr)
      val error = ujson.read(run.stdout)("error")
      assertEquals(error("code"), ujson.Str("BATCH_OP_FAILED"))
      assert(
        error("message").str.contains("Sheet not found: Dat. Available: Data, Summary"),
        error("message").str
      )
      assertEquals(error("candidates"), ujson.Arr(ujson.Str("Data")))
      assertEquals(error("location")("opIndex"), ujson.Num(1))
      assert(!Files.exists(out))
    }
  }

  test("a `sheet` key on define-name is not its scope: UNKNOWN_PROPERTY, the name stays global") {
    val out = fresh("name-ops-sheet-key.xlsx")
    val json = """[{"op":"define-name","name":"X","refersTo":"Data!$A$1","sheet":"Data"}]"""
    for
      run <- batch(named, out, json, "--json")
      listed <- names(out)
    yield
      assertEquals(run.exit, 0, run.stderr)
      val warnings = ujson.read(run.stdout)("warnings").arr
      assertEquals(warnings.map(_("code").str).toVector, Vector("UNKNOWN_PROPERTY"))
      assert(warnings(0)("message").str.contains("sheet"), warnings(0)("message").str)
      assertEquals(listed.map(n => (n._1, n._3)), Vector(("Total", ujson.Null), ("X", ujson.Null)))
  }

  test("a define-name without refersTo is refused at parse time, naming the field") {
    val out = fresh("name-ops-no-refers.xlsx")
    batch(named, out, """[{"op":"define-name","name":"X"}]""", "--json").map { run =>
      assertEquals(run.exit, 2, run.stderr)
      val error = ujson.read(run.stdout)("error")
      assertEquals(error("code"), ujson.Str("BATCH_OP_INVALID"))
      assert(error("message").str.contains("refersTo"), error("message").str)
      assert(!Files.exists(out))
    }
  }

  test("--dry-run validates define-name / remove-name and writes nothing") {
    val out = fresh("name-ops-dry.xlsx")
    val json =
      """[{"op":"define-name","name":"Tax","refersTo":"Data!$A$1"},{"op":"remove-name","name":"Total","scope":"Data"}]"""
    // `--dry-run` is `batch`'s own option, so it follows the verb (the globals precede it)
    CliHarness.run(List("-f", named, "-o", out.toString, "batch", "--dry-run", "-"), json).map {
      run =>
        assertEquals(run.exit, 0, run.stderr)
        assert(run.stdout.contains("DEFINE-NAME Tax -> Data!$A$1"), run.stdout)
        assert(run.stdout.contains("REMOVE-NAME Total (scope: Data)"), run.stdout)
        assert(!Files.exists(out), "a dry run writes nothing")
    }
  }

  test("--stream batch refuses define-name / remove-name by index before writing anything") {
    val out = fresh("name-ops-stream.xlsx")
    val json =
      """[{"op":"put","ref":"A9","value":1},{"op":"define-name","name":"Tax","refersTo":"Data!$A$1"},{"op":"remove-name","name":"Total"}]"""
    CliHarness
      .run(List("-f", named, "-s", "Data", "-o", out.toString, "--stream", "batch", "-"), json)
      .map { run =>
        assertEquals(run.exit, 2, run.stderr)
        assert(run.stderr.contains("  code: UNSUPPORTED_IN_STREAM"), run.stderr)
        assert(run.stderr.contains("[2 define-name, 3 remove-name]"), run.stderr)
        assert(!Files.exists(out), "a refusal writes nothing")
      }
  }
