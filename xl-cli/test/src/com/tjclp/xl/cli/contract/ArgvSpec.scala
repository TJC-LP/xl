package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}
import org.scalacheck.Prop.forAll

import com.tjclp.xl.cli.{Cli, CliIO}

/**
 * The argv contract (ADR-017 §2.2): global flags are accepted anywhere on the command line, an
 * unknown verb gets a suggestion and a one-line usage, and a wrong command line is one usage line
 * plus the first parser error — never the full subcommand dump.
 */
class ArgvSpec extends CatsEffectSuite with ScalaCheckSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "argv-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString

  // ---------------------------------------------------------------------------------------------
  // wantsJson: the flag, not a value spelled like it

  test("wantsJson: --json anywhere before `--` is the flag; as a global's value it is data") {
    assertEquals(Argv.wantsJson(List("--json", "sheets")), true)
    assertEquals(Argv.wantsJson(List("-f", "a.xlsx", "view", "A1", "--json")), true)
    assertEquals(Argv.wantsJson(List("-o", "--json", "put", "A1", "1")), false)
    assertEquals(Argv.wantsJson(List("--output", "--json", "put", "A1", "1")), false)
    assertEquals(Argv.wantsJson(List("-s", "--json", "--json", "view", "A1")), true)
    assertEquals(Argv.wantsJson(List("--file=--json", "--json", "sheets")), true)
    assertEquals(Argv.wantsJson(List("search", "--", "--json")), false)
    assertEquals(Argv.wantsJson(Nil), false)
  }

  test(
    "wantsJson agrees with hoist: the mode a usage failure renders is the mode a parse would run"
  ) {
    val globals = Argv.globals.keys.toList
    val gen = Gen.listOfN(6, Gen.oneOf(globals ++ List("--json", "view", "A1", "x.xlsx", "--")))
    Prop.forAll(gen) { args =>
      // hoisting never changes whether --json is the flag
      Argv.wantsJson(Argv.hoist(args)) == Argv.wantsJson(args)
    }
  }

  // ---------------------------------------------------------------------------------------------
  // hoist: laws
  // ---------------------------------------------------------------------------------------------

  /** Tokens that are neither a global nor `--`: verbs, refs, values, a verb's own flags. */
  private val genPlain: Gen[String] =
    Gen.oneOf("view", "A1:B2", "5", "--format", "json", "--eval", "put", "x.xlsx", "Data", "-5")

  /**
   * A well-formed global: a switch, or a value-taking flag WITH its value (`-f x` or `--file=x`). A
   * value-taking flag left dangling at the end is a malformed line decline rejects; its hoisting is
   * pinned by a unit test below, not by the laws.
   */
  private val genGlobal: Gen[List[String]] =
    Gen.oneOf(Argv.globals.toList).flatMap { (name, takesValue) =>
      if takesValue then
        Gen.oneOf(
          Gen.oneOf("f.xlsx", "Data", "out.xlsx", "500").map(v => List(name, v)),
          if name.startsWith("--") then Gen.const(List(s"$name=value"))
          else Gen.const(List(name, "value"))
        )
      else Gen.const(List(name))
    }

  private val genArgs: Gen[List[String]] =
    Gen.listOf(Gen.frequency(3 -> genPlain.map(List(_)), 2 -> genGlobal)).map(_.flatten)

  private val genPlainArgs: Gen[List[String]] = Gen.listOf(genPlain)

  property("hoist is idempotent (on well-formed command lines)") {
    forAll(genArgs) { args =>
      Argv.hoist(Argv.hoist(args)) == Argv.hoist(args)
    }
  }

  property("hoist is a permutation: no token appears, disappears or changes") {
    forAll(genArgs) { args =>
      Argv.hoist(args).sorted == args.sorted
    }
  }

  property("a command line without globals is untouched") {
    forAll(genPlainArgs) { args =>
      Argv.hoist(args) == args
    }
  }

  property("a global and its value move in front of the verb; every other token keeps its order") {
    val gen = for
      before <- genPlainArgs.map(_.filterNot(_ == "view"))
      after <- genPlainArgs
      global <- genGlobal
    yield (before, after, global)
    forAll(gen) { (before, after, global) =>
      // `view` first so --strict (the one verb-owned global) is never what we hoist here
      val plain = "view" :: before ++ after
      val hoisted = Argv.hoist(("view" :: before) ++ global ++ after)
      global.headOption.contains("--strict") || hoisted == global ++ plain
    }
  }

  property("`--` ends hoisting: everything after it stays where it is, verbatim") {
    forAll(genArgs, genArgs) { (head, tail) =>
      val cut = head.takeWhile(_ != "--")
      Argv.hoist(cut ++ ("--" :: tail)) == Argv.hoist(cut) ++ ("--" :: tail)
    }
  }

  test("`--` shields a positional that spells a global: `search -- --json` keeps its pattern") {
    // Without the escape the flag is hoisted and `search` is left with no pattern (a usage error);
    // after `--` the same token is data, so `search` looks for the text `--json`.
    assertEquals(
      Argv.hoist(List("-f", "a.xlsx", "search", "--json")),
      List("-f", "a.xlsx", "--json", "search")
    )
    assertEquals(
      Argv.hoist(List("-f", "a.xlsx", "search", "--", "--json")),
      List("-f", "a.xlsx", "search", "--", "--json")
    )
  }

  // ---------------------------------------------------------------------------------------------
  // hoist: the verb-owned exceptions and the --flag=value form
  // ---------------------------------------------------------------------------------------------

  test("view owns --strict: `view A1:B2 --eval --strict` is not hoisted") {
    assertEquals(
      Argv.hoist(List("-f", "f.xlsx", "view", "A1:B2", "--eval", "--strict")),
      List("-f", "f.xlsx", "view", "A1:B2", "--eval", "--strict")
    )
  }

  test("`recalc --strict` is hoisted (the write gate is the global flag)") {
    assertEquals(
      Argv.hoist(List("-f", "f.xlsx", "recalc", "--strict", "-o", "out.xlsx")),
      List("-f", "f.xlsx", "--strict", "-o", "out.xlsx", "recalc")
    )
  }

  test(
    "new owns --sheet and --backend: `new out.xlsx --sheet Data --backend saxstax` is untouched"
  ) {
    val args =
      List("new", "out.xlsx", "--sheet", "Data", "--sheet", "Summary", "--backend", "saxstax")
    assertEquals(Argv.hoist(args), args)
  }

  test("the --flag=value form hoists as one token") {
    assertEquals(
      Argv.hoist(List("view", "A1:B2", "--file=f.xlsx", "--sheet=Data")),
      List("--file=f.xlsx", "--sheet=Data", "view", "A1:B2")
    )
  }

  test("globals after the verb land in front of it, in their original order") {
    assertEquals(
      Argv.hoist(List("view", "A1:B2", "-f", "f.xlsx", "--json", "-s", "Data")),
      List("-f", "f.xlsx", "--json", "-s", "Data", "view", "A1:B2")
    )
  }

  test("a value-taking global as the last token stays put (decline reports it, not a bogus verb)") {
    assertEquals(Argv.hoist(List("view", "A1", "-f")), List("view", "A1", "-f"))
    assertEquals(Argv.verbOf(Argv.hoist(List("view", "A1", "-f"))), Some("view"))
  }

  test("verbOf skips every flag, known or not: decline reports the unknown one") {
    assertEquals(Argv.verbOf(List("--frobnicate", "view")), Some("view"))
    assertEquals(Argv.verbOf(List("-f", "x.xlsx", "--frobnicate", "view", "A1:B2")), Some("view"))
    assertEquals(Argv.verbOf(List("--frobnicate")), None)
  }

  test("an unknown flag is decline's `Unexpected option`, never an unknown verb") {
    for
      unknownFlag <- CliHarness.run("-f", file("simple.xlsx"), "--frobnicate", "view", "A1:B2")
      dangling <- CliHarness.run("view", "A1:B2", "-f")
      bareFlag <- CliHarness.run("--frobnicate")
    yield
      assertEquals(unknownFlag.exit, 2)
      assert(
        unknownFlag.stderr.startsWith("Error: Unexpected option: --frobnicate"),
        unknownFlag.stderr
      )
      assertEquals(dangling.exit, 2)
      assert(!dangling.stderr.contains("unknown verb"), dangling.stderr)
      assert(dangling.stderr.contains("  code: USAGE"), dangling.stderr)
      assertEquals(bareFlag.exit, 2)
      assert(!bareFlag.stderr.contains("unknown verb"), bareFlag.stderr)
  }

  test("a verb missing its sub-verb keeps decline's own short list") {
    for
      name <- CliHarness.run("-f", file("simple.xlsx"), "name")
      cf <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "cf")
    yield
      assertEquals(name.exit, 2)
      assert(name.stderr.contains("Missing expected command (add or rm)!"), name.stderr)
      assert(name.stderr.contains("run `xl name --help`"), name.stderr)
      assertEquals(cf.exit, 2)
      assert(cf.stderr.contains("Missing expected command (add or list)!"), cf.stderr)
  }

  // ---------------------------------------------------------------------------------------------
  // verbOf and verbs
  // ---------------------------------------------------------------------------------------------

  test("verbOf skips globals and their values, --help/--version, and stops at --") {
    assertEquals(Argv.verbOf(List("-s", "view", "--json", "cell", "A1")), Some("cell"))
    assertEquals(Argv.verbOf(List("--file=view", "--json", "cell")), Some("cell"))
    assertEquals(Argv.verbOf(List("-f", "x.xlsx", "search", "--", "view")), Some("search"))
    assertEquals(Argv.verbOf(List("-s")), None)
    assertEquals(Argv.verbOf(List("-f", "view")), None)
    assertEquals(Argv.verbOf(List("--help")), None)
    assertEquals(Argv.verbOf(List("-f", "x.xlsx", "--version")), None)
    assertEquals(Argv.verbOf(List("view", "--help")), Some("view"))
    assertEquals(Argv.verbOf(List("viwe", "A1")), Some("viwe"))
    assertEquals(Argv.verbOf(Nil), None)
  }

  test("Argv.verbs is exactly the subcommand list decline renders, in usage order") {
    val usage = Cli.command(CliIO.system).showHelp.linesIterator.toVector
    val rendered = usage
      .takeWhile(_.trim.nonEmpty)
      .drop(1) // "Usage:"
      .map(_.trim.split(" ").last)
      .distinct
    assertEquals(rendered, Argv.verbs)
    assert(Argv.verbs.containsSlice(Vector("describe", "audit", "deps")), Argv.verbs.toString)
    assertEquals(Argv.verbs.distinct, Argv.verbs, "no verb listed twice")
  }

  test("globals is the documented table: value-taking flags true, switches false") {
    val valueTaking =
      Set("-f", "--file", "-s", "--sheet", "-o", "--output", "--backend", "--max-size")
    val switches =
      Set("-i", "--in-place", "--stream", "--no-recalc", "--preserve-caches", "--json", "--strict")
    assertEquals(Argv.globals.keySet, valueTaking ++ switches)
    valueTaking.foreach(k => assertEquals(Argv.globals(k), true, k))
    switches.foreach(k => assertEquals(Argv.globals(k), false, k))
  }

  // ---------------------------------------------------------------------------------------------
  // Through the harness
  // ---------------------------------------------------------------------------------------------

  test("`xl view A1:B2 -f f -s Data` ≡ `xl -f f -s Data view A1:B2` (exit, stdout, stderr)") {
    val simple = file("simple.xlsx")
    for
      canonical <- CliHarness.run("-f", simple, "-s", "Data", "view", "A1:B2")
      trailing <- CliHarness.run("view", "A1:B2", "-f", simple, "-s", "Data")
      mixed <- CliHarness.run("-s", "Data", "view", "-f", simple, "A1:B2")
    yield
      assertEquals(canonical.exit, 0, canonical.stderr)
      assertEquals(trailing, canonical)
      assertEquals(mixed, canonical)
  }

  test("`--json` after the verb still yields the envelope") {
    CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "view", "A1:B2", "--json").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val envelope = ujson.read(run.stdout)
      assertEquals(envelope("ok"), ujson.True)
      assertEquals(envelope("verb"), ujson.Str("view"))
    }
  }

  test("`xl -f f recalc --strict -o out` gates: --strict after the verb is the write gate") {
    val out = file("recalc-strict-after-verb.xlsx")
    for
      trailing <- CliHarness.run("-f", file("circular.xlsx"), "recalc", "--strict", "-o", out)
      leading <- CliHarness.run("-f", file("circular.xlsx"), "--strict", "-o", out, "recalc")
    yield
      assertEquals(trailing.exit, 1, trailing.stderr)
      assertEquals(trailing, leading)
  }

  test("`view A1:B2 --eval --strict` keeps --strict as view's own flag") {
    CliHarness
      .run("-f", file("circular.xlsx"), "-s", "Data", "view", "A1:B2", "--eval", "--strict")
      .map { run =>
        assertEquals(run.exit, 1, run.stderr)
        assertEquals(run.stdout, "")
        assert(run.stderr.contains("Error:"), run.stderr)
      }
  }

  test("an unknown verb exits 2 with UNKNOWN_VERB, a did-you-mean line and fewer than 10 lines") {
    CliHarness.run("-f", file("simple.xlsx"), "viwe", "A1").map { run =>
      assertEquals(run.exit, 2)
      assertEquals(run.stdout, "")
      assert(run.stderr.contains("  did you mean: view"), run.stderr)
      assert(run.stderr.contains("  code: UNKNOWN_VERB"), run.stderr)
      assert(run.stderr.startsWith("Error: unknown verb 'viwe'\n"), run.stderr)
      assert(
        run.stderr.contains("usage: xl [-f FILE] [-s SHEET] [-o OUT | -i] [--json] <verb>"),
        run.stderr
      )
      assert(run.stderr.linesIterator.size < 10, run.stderr)
    }
  }

  test("an unknown verb under --json is the UNKNOWN_VERB envelope with candidates") {
    CliHarness.run("--json", "viwe").map { run =>
      assertEquals(run.exit, 2)
      val envelope = ujson.read(run.stdout)
      assertEquals(envelope("ok"), ujson.False)
      assertEquals(envelope("verb"), ujson.Str(""))
      assertEquals(envelope("error")("code"), ujson.Str("UNKNOWN_VERB"))
      assertEquals(envelope("error")("candidates"), ujson.Arr(ujson.Str("view")))
      assertEquals(run.stderr, "Error: unknown verb 'viwe'\n")
    }
  }

  test("a wrong command line is one usage line plus the first parser error, never the verb dump") {
    for
      // W2.4: `view` no longer needs a range (it defaults to the used range); `cell` still does
      missing <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "cell")
      noArgs <- CliHarness.run()
    yield
      assertEquals(missing.exit, 2)
      assertEquals(missing.stdout, "")
      val lines = missing.stderr.linesIterator.toVector
      // Error: first, as on every failure (the stderr contract); the usage line closes the block
      assertEquals(lines.headOption, Some("Error: Missing expected positional argument!"))
      assertEquals(
        lines.lastOption,
        Some("usage: xl [-f FILE] [-s SHEET] [-o OUT | -i] [--json] <verb> …")
      )
      assert(lines.contains("  code: USAGE"), missing.stderr)
      assert(lines.exists(_.contains("run `xl cell --help`")), missing.stderr)
      assert(lines.size < 10, missing.stderr)
      assert(!missing.stderr.contains("ungroup-cols"), "no subcommand dump: " + missing.stderr)

      assertEquals(noArgs.exit, 2)
      assert(noArgs.stderr.startsWith("Error: "), noArgs.stderr)
      assert(noArgs.stderr.contains("usage: xl "), noArgs.stderr)
      assert(noArgs.stderr.linesIterator.size < 10, noArgs.stderr)
      assert(!noArgs.stderr.contains("ungroup-cols"), "no subcommand dump: " + noArgs.stderr)
  }

  test("--help, <verb> --help and --version keep working") {
    for
      help <- CliHarness.run("--help")
      verbHelp <- CliHarness.run("put", "--help")
      version <- CliHarness.run("--version")
      verbHelpAfterGlobals <- CliHarness.run("-f", file("simple.xlsx"), "view", "--help")
    yield
      assertEquals(help.exit, 0)
      assert(help.stdout.startsWith("Usage:"), help.stdout)
      assertEquals(verbHelp.exit, 0)
      assert(verbHelp.stdout.contains("Write value(s) to cell or range."), verbHelp.stdout)
      assertEquals(verbHelp.stderr, "")
      assertEquals(version.exit, 0)
      assert(version.stdout.trim.nonEmpty)
      assertEquals(verbHelpAfterGlobals.exit, 0)
      assert(verbHelpAfterGlobals.stdout.contains("View a range"), verbHelpAfterGlobals.stdout)
  }

  test("GH-619: positionalFile names the workbook an agent passed without -f") {
    assertEquals(Argv.positionalFile(List("view", "input.xlsx", "A1:B4")), Some("input.xlsx"))
    assertEquals(Argv.positionalFile(List("view", "A1:B4", "in.XLSM")), Some("in.XLSM"))
    assertEquals(Argv.positionalFile(List("-s", "Data", "view", "in.xlsx")), Some("in.xlsx"))
    // the value of -f/--file, -o or -g is not a misplaced -f
    assertEquals(Argv.positionalFile(List("-f", "in.xlsx", "view", "A1")), None)
    assertEquals(Argv.positionalFile(List("--file=in.xlsx", "view", "A1")), None)
    assertEquals(Argv.positionalFile(List("-o", "out.xlsx", "put", "A1", "1")), None)
    assertEquals(Argv.positionalFile(List("diff", "-g", "b.xlsx")), None)
    // verbs whose own positional is a file, tokens behind --, flags, and non-workbook names
    assertEquals(Argv.positionalFile(List("lint", "book.xlsx")), None)
    assertEquals(Argv.positionalFile(List("new", "book.xlsx")), None)
    assertEquals(Argv.positionalFile(List("search", "--", "report.xlsx")), None)
    assertEquals(Argv.positionalFile(List("view", "--eval", "A1", "data.csv")), None)
    assertEquals(Argv.positionalFile(List("view", ".xlsx")), None)
    assertEquals(Argv.positionalFile(Nil), None)
  }

  test("GH-619: a positional workbook gets the -f hint in text and --json usage errors") {
    for
      text <- CliHarness.run("view", "input.xlsx", "A1:B4")
      json <- CliHarness.run("--json", "view", "input.xlsx", "A1:B4")
      plain <- CliHarness.run("-f", file("simple.xlsx"), "view", "A1:B4", "extra")
    yield
      assertEquals(text.exit, 2)
      assertEquals(text.stdout, "")
      assert(text.stderr.startsWith("Error: Unexpected argument: A1:B4"), text.stderr)
      assert(
        text.stderr.contains(
          "  hint: did you mean -f input.xlsx? xl takes the file as -f/--file; " +
            "the positional argument is the range or verb argument"
        ),
        text.stderr
      )
      assertEquals(json.exit, 2)
      val error = ujson.read(json.stdout)("error")
      assertEquals(error("code"), ujson.Str("USAGE"))
      assert(error("hint").str.startsWith("did you mean -f input.xlsx?"), error("hint").str)
      // with -f given, the usual --help hint stands
      assertEquals(plain.exit, 2)
      assert(plain.stderr.contains("  hint: run `xl view --help` for the usage"), plain.stderr)
  }

  test("`search -- --json` searches for the text `--json`; `search --json` is a usage error") {
    for
      escaped <- CliHarness.run("-f", file("simple.xlsx"), "search", "--", "--json")
      hoisted <- CliHarness.run("-f", file("simple.xlsx"), "search", "--json")
    yield
      assertEquals(escaped.exit, 0, escaped.stderr)
      assert(escaped.stdout.startsWith("Found 0 matches"), escaped.stdout)
      // the hoisted --json puts the run in JSON mode: the usage error is an envelope on stdout
      assertEquals(hoisted.exit, 2, hoisted.stdout)
      val envelope = ujson.read(hoisted.stdout)
      assertEquals(envelope("ok"), ujson.False)
      assertEquals(envelope("error")("code"), ujson.Str("USAGE"))
  }

  test("decline's own `--` handling: `put A1 -- -5` stores -5") {
    val out = file("put-negative.xlsx")
    for
      put <- CliHarness.run("-f", file("single.xlsx"), "-o", out, "put", "A1", "--", "-5")
      cell <- CliHarness.run("-f", out, "cell", "A1")
    yield
      assertEquals(put.exit, 0, put.stderr)
      assert(Files.exists(Path.of(out)))
      assertEquals(cell.exit, 0, cell.stderr)
      assert(cell.stdout.contains("-5"), cell.stdout)
  }
