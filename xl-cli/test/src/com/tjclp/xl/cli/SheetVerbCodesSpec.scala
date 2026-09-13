package com.tjclp.xl.cli

import java.nio.file.Files

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.cli.contract.{CliHarness, CliRun, EnvelopeSchema, TestFixtures}
import com.tjclp.xl.ooxml.XlsxReader

/**
 * GH-608: the sheet-management verbs' refusals are typed — none falls through to `INTERNAL`, the
 * code the contract reserves for defects (ADR-017 §2.3). One harness run per code: `USAGE` (exit 2)
 * for a `move-sheet` with no position, `SHEET_NOT_FOUND` (with did-you-mean candidates) for an
 * unknown sheet on `add-sheet --after`, `remove-sheet`, `copy-sheet`, `sheets hide` and `sheets
 * show`, `INVALID_SHEET_NAME` for a name Excel rejects, `NAME_NOT_FOUND` for an unknown named
 * range. Every refusal leaves no output file behind.
 */
class SheetVerbCodesSpec extends CatsEffectSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "sheet-verb-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString

  /**
   * `--json` run against `simple.xlsx` (sheets `Data`, `Summary`) writing to `<DIR>/<tag>.xlsx`.
   */
  private def refused(tag: String, verb: String*): IO[(CliRun, ujson.Value)] =
    val out = fixtures().resolve(s"$tag.xlsx")
    CliHarness
      .run(List("-f", file("simple.xlsx"), "-o", out.toString, "--json") ++ verb.toList, "")
      .map { run =>
        val envelope = ujson.read(run.stdout)
        EnvelopeSchema.assertValid(envelope)
        assertEquals(envelope("ok"), ujson.False, run.stdout)
        assert(!Files.exists(out), s"$tag: nothing may be written when the verb is refused")
        (run, envelope("error"))
      }

  test("move-sheet without --to/--after/--before is USAGE, exit 2, before the file is read") {
    // The position is validated by the command-line parser: a file that does not exist is never
    // opened, so the code is USAGE, not IO_READ.
    val missing = fixtures().resolve("no-such-file.xlsx").toString
    val out = fixtures().resolve("move-usage.xlsx").toString
    CliHarness.run("-f", missing, "-o", out, "--json", "move-sheet", "Data").map { run =>
      assertEquals(run.exit, 2, run.stderr)
      val error = ujson.read(run.stdout)("error")
      assertEquals(error("code"), ujson.Str("USAGE"))
      assert(
        error("message").str.contains("move-sheet requires --to, --after, or --before"),
        error("message").str
      )
    }
  }

  test("add-sheet --after an unknown sheet is SHEET_NOT_FOUND with did-you-mean candidates") {
    for
      (json, error) <- refused("add-after", "add-sheet", "New", "--after", "Dat")
      text <- CliHarness.run(
        "-f",
        file("simple.xlsx"),
        "-o",
        fixtures().resolve("add-after-text.xlsx").toString,
        "add-sheet",
        "New",
        "--after",
        "Dat"
      )
    yield
      assertEquals(json.exit, 3, json.stderr)
      assertEquals(error("code"), ujson.Str("SHEET_NOT_FOUND"))
      // the CLI's one SHEET_NOT_FOUND text (Resolve's), not a verb-specific spelling
      assertEquals(error("message"), ujson.Str("Sheet not found: Dat. Available: Data, Summary"))
      assertEquals(error("candidates").arr.map(_.str).toVector, Vector("Data"))
      assertEquals(text.exit, 3, text.stderr)
      assert(text.stderr.contains("  code: SHEET_NOT_FOUND\n"), text.stderr)
      assert(text.stderr.contains("  did you mean: Data\n"), text.stderr)
  }

  test("add-sheet with a name Excel rejects is INVALID_SHEET_NAME") {
    refused("add-invalid", "add-sheet", "a[b").map { (run, error) =>
      assertEquals(run.exit, 3, run.stderr)
      assertEquals(error("code"), ujson.Str("INVALID_SHEET_NAME"))
      assert(error("message").str.contains("invalid characters"), error("message").str)
    }
  }

  test("GH-626: name rm of an unknown named range is NAME_NOT_FOUND with did-you-mean candidates") {
    val out = fixtures().resolve("name-rm.xlsx")
    CliHarness
      .run("-f", file("named.xlsx"), "-o", out.toString, "--json", "name", "rm", "Totl")
      .map { run =>
        assertEquals(run.exit, 3, run.stderr)
        val error = ujson.read(run.stdout)("error")
        assertEquals(error("code"), ujson.Str("NAME_NOT_FOUND"))
        assertEquals(error("message"), ujson.Str("Named range 'Totl' not found. Available: Total"))
        assertEquals(error("candidates"), ujson.Arr(ujson.Str("Total")))
        assertEquals(error("hint"), ujson.Str("list defined names with `xl -f <file> names`"))
        assert(!Files.exists(out), "nothing may be written when the name is unknown")
      }
  }

  test("name rm on a book with no defined names is NAME_NOT_FOUND without candidates") {
    refused("name-rm-none", "name", "rm", "Nope").map { (run, error) =>
      assertEquals(run.exit, 3, run.stderr)
      assertEquals(error("code"), ujson.Str("NAME_NOT_FOUND"))
      assertEquals(error("message"), ujson.Str("Named range 'Nope' not found"))
      assertEquals(error("candidates"), ujson.Arr())
    }
  }

  test(
    "remove-sheet, copy-sheet, sheets hide and sheets show of an unknown sheet: SHEET_NOT_FOUND"
  ) {
    val cases = Vector(
      "remove" -> Vector("remove-sheet", "Nope"),
      "copy" -> Vector("copy-sheet", "Nope", "Copy"),
      "hide" -> Vector("sheets", "hide", "Nope"),
      "show" -> Vector("sheets", "show", "Nope")
    )
    cases.foldLeft(IO.unit) { case (acc, (tag, verb)) =>
      acc *> refused(s"nf-$tag", verb*).map { (run, error) =>
        assertEquals(run.exit, 3, s"$tag: ${run.stderr}")
        assertEquals(error("code"), ujson.Str("SHEET_NOT_FOUND"), tag)
        assertEquals(
          error("message"),
          ujson.Str("Sheet not found: Nope. Available: Data, Summary"),
          tag
        )
        assertEquals(error("hint"), ujson.Str("list sheets with `xl -f <file> sheets`"), tag)
      }
    }
  }

  // ===== GH-462 / GH-538: `-s` is the scope of `name add|rm`; names match case-insensitively =====

  /** `names --json` of a written file: `[{name, refersTo, scope, hidden}]`. */
  private def namesJson(path: java.nio.file.Path): IO[Vector[(String, ujson.Value)]] =
    CliHarness.run("-f", path.toString, "--json", "names").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      ujson.read(run.stdout)("data").arr.map(n => (n("name").str, n("scope"))).toVector
    }

  test("GH-538: name rm matches case-insensitively — `name rm total` removes Total") {
    val out = fixtures().resolve("name-rm-case.xlsx")
    for
      run <- CliHarness.run("-f", file("named.xlsx"), "-o", out.toString, "name", "rm", "total")
      names <- namesJson(out)
    yield
      assertEquals(run.exit, 0, run.stderr)
      assert(run.stdout.contains("Removed named range 'total'\n"), run.stdout)
      assertEquals(names, Vector.empty)
  }

  test("GH-462: name add -s scopes the name to that sheet; name rm -s removes only that one") {
    val added = fixtures().resolve("name-add-scoped.xlsx")
    val removed = fixtures().resolve("name-rm-scoped.xlsx")
    for
      add <- CliHarness.run(
        List("-f", file("named.xlsx"), "-s", "Data", "-o", added.toString) ++
          List("name", "add", "Local", "Data!$A$1"),
        ""
      )
      afterAdd <- namesJson(added)
      rm <- CliHarness.run(
        List("-f", added.toString, "-s", "Data", "-o", removed.toString, "name", "rm", "local"),
        ""
      )
      afterRm <- namesJson(removed)
    yield
      assertEquals(add.exit, 0, add.stderr)
      assert(
        add.stdout.contains("Added named range 'Local' -> Data!$A$1 (scope: Data)\n"),
        add.stdout
      )
      assertEquals(afterAdd, Vector("Total" -> ujson.Null, "Local" -> ujson.Str("Data")))
      assertEquals(rm.exit, 0, rm.stderr)
      assert(rm.stdout.contains("Removed named range 'local' (scope: Data)\n"), rm.stdout)
      assertEquals(afterRm, Vector("Total" -> ujson.Null))
  }

  test("GH-462: name rm without -s of a name that exists only sheet-scoped is NAME_NOT_FOUND") {
    val added = fixtures().resolve("name-scoped-only.xlsx")
    val out = fixtures().resolve("name-rm-scoped-only.xlsx")
    for
      _ <- CliHarness.run(
        List("-f", file("named.xlsx"), "-s", "Data", "-o", added.toString) ++
          List("name", "add", "Local", "Data!$A$1"),
        ""
      )
      run <- CliHarness.run(
        "-f",
        added.toString,
        "-o",
        out.toString,
        "--json",
        "name",
        "rm",
        "Local"
      )
    yield
      assertEquals(run.exit, 3, run.stderr)
      val error = ujson.read(run.stdout)("error")
      assertEquals(error("code"), ujson.Str("NAME_NOT_FOUND"))
      // the names THIS form can remove — the workbook-scoped ones — never the scoped entry
      assertEquals(error("message"), ujson.Str("Named range 'Local' not found. Available: Total"))
      assert(!Files.exists(out), "nothing may be written when the name is unknown in that scope")
  }

  test("GH-462: name rm -s of an unknown sheet is SHEET_NOT_FOUND before any write") {
    val out = fixtures().resolve("name-rm-nope.xlsx")
    CliHarness
      .run(
        List("-f", file("named.xlsx"), "-s", "Nope", "-o", out.toString, "--json") ++
          List("name", "rm", "Total"),
        ""
      )
      .map { run =>
        assertEquals(run.exit, 3, run.stderr)
        val error = ujson.read(run.stdout)("error")
        assertEquals(error("code"), ujson.Str("SHEET_NOT_FOUND"))
        assertEquals(error("message"), ujson.Str("Sheet not found: Nope. Available: Data, Summary"))
        assert(!Files.exists(out), "nothing may be written when the scope is unknown")
      }
  }

  /** Sheet `Data`'s print fields as the reader lifts them from a written file. */
  private def dataPrintFields(
    path: java.nio.file.Path
  ): (Option[com.tjclp.xl.addressing.CellRange], Option[(Int, Int)]) =
    val setup = XlsxReader
      .read(path)
      .fold(e => fail(s"reread failed: $e"), identity)
      .sheets
      .find(_.name.value == "Data")
      .flatMap(_.pageSetup)
    (setup.flatMap(_.printArea), setup.flatMap(_.repeatRows))

  test("GH-462: name rm -s removes a Print_Area the read lifted into page setup (its own output)") {
    // After any read a sheet's _xlnm.Print_Area lives in pageSetup.printArea, not in the table
    // `names` lists raw from workbook.xml — the documented inverse of `name add -s` must still work.
    val added = fixtures().resolve("print-area-add.xlsx")
    val removed = fixtures().resolve("print-area-rm.xlsx")
    val missing = fixtures().resolve("print-area-missing.xlsx")
    for
      add <- CliHarness.run(
        List("-f", file("named.xlsx"), "-s", "Data", "-o", added.toString) ++
          List("name", "add", "_xlnm.Print_Area", "Data!$A$1:$B$2"),
        ""
      )
      afterAdd <- namesJson(added)
      rm <- CliHarness.run(
        List("-f", added.toString, "-s", "Data", "-o", removed.toString) ++
          List("name", "rm", "_xlnm.print_area"),
        ""
      )
      afterRm <- namesJson(removed)
      // the lifted print name is an entry of the sheet's scope: it is offered as a candidate
      unknown <- CliHarness.run(
        List("-f", added.toString, "-s", "Data", "-o", missing.toString, "--json") ++
          List("name", "rm", "Nope"),
        ""
      )
    yield
      assertEquals(add.exit, 0, add.stderr)
      assertEquals(afterAdd, Vector("Total" -> ujson.Null, "_xlnm.Print_Area" -> ujson.Str("Data")))
      assertEquals(rm.exit, 0, rm.stderr)
      assert(
        rm.stdout.contains("Removed named range '_xlnm.print_area' (scope: Data)\n"),
        rm.stdout
      )
      assertEquals(afterRm, Vector("Total" -> ujson.Null))
      assertEquals(dataPrintFields(removed), (None, None))
      assertEquals(unknown.exit, 3, unknown.stderr)
      val error = ujson.read(unknown.stdout)("error")
      assertEquals(error("code"), ujson.Str("NAME_NOT_FOUND"))
      assertEquals(
        error("message"),
        ujson.Str("Named range 'Nope' not found. Available: _xlnm.Print_Area")
      )
      assert(!Files.exists(missing), "nothing may be written when the name is unknown")
  }

  test("GH-462: name rm -s removes a Print_Titles the read lifted into page setup") {
    val added = fixtures().resolve("print-titles-add.xlsx")
    val removed = fixtures().resolve("print-titles-rm.xlsx")
    for
      add <- CliHarness.run(
        List("-f", file("named.xlsx"), "-s", "Data", "-o", added.toString) ++
          List("name", "add", "_xlnm.Print_Titles", "Data!$1:$2"),
        ""
      )
      rm <- CliHarness.run(
        List("-f", added.toString, "-s", "Data", "-o", removed.toString) ++
          List("name", "rm", "_XLNM.PRINT_TITLES"),
        ""
      )
      afterRm <- namesJson(removed)
    yield
      assertEquals(add.exit, 0, add.stderr)
      assertEquals(dataPrintFields(added), (None, Some((1, 2))))
      assertEquals(rm.exit, 0, rm.stderr)
      assertEquals(afterRm, Vector("Total" -> ujson.Null))
      assertEquals(dataPrintFields(removed), (None, None))
  }

  test(
    "GH-462: without -s, name add on a single-sheet book stays workbook-scoped (no auto-select)"
  ) {
    val out = fixtures().resolve("name-add-single.xlsx")
    for
      run <- CliHarness.run(
        List("-f", file("single.xlsx"), "-o", out.toString, "name", "add", "Solo", "Sheet1!$A$1"),
        ""
      )
      names <- namesJson(out)
    yield
      assertEquals(run.exit, 0, run.stderr)
      assert(run.stdout.contains("Added named range 'Solo' -> Sheet1!$A$1\n"), run.stdout)
      assert(!run.stdout.contains("scope"), run.stdout)
      assertEquals(names, Vector("Solo" -> ujson.Null))
  }
