package com.tjclp.xl.cli

import java.nio.file.Files

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.cli.contract.{CliHarness, CliRun, EnvelopeSchema, TestFixtures}

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
