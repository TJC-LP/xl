package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.cli.Cli

/**
 * The global `--json` flag end to end (ADR-017 §2.4), through the in-process harness: every verb,
 * success or failure, prints exactly one envelope on stdout; `--format json` payloads ride inside
 * it unchanged; failures put one `Error: <message>` line on stderr; and every envelope validates
 * against `envelope.schema.json`.
 */
class EnvelopeSpec extends CatsEffectSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "envelope-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  private def file(name: String): String = fixtures().resolve(name).toString

  /**
   * The envelope a run printed: stdout must be exactly one indented JSON document (plus the final
   * newline), valid against the schema, with the seven keys.
   */
  private def envelope(run: CliRun): ujson.Obj =
    val parsed = ujson.read(run.stdout)
    EnvelopeSchema.assertValid(parsed)
    assertEquals(run.stdout, ujson.write(parsed, indent = 2) + "\n", "stdout is the envelope alone")
    assertEquals(
      parsed.obj.keys.toList,
      List("ok", "exitCode", "verb", "version", "data", "warnings", "error")
    )
    assertEquals(parsed("exitCode"), ujson.Num(run.exit), "exitCode mirrors the process exit code")
    assertEquals(parsed("ok"), ujson.Bool(parsed("error") == ujson.Null), "ok ⇔ error == null")
    parsed.obj

  private def oneErrorLine(run: CliRun, message: String): Unit =
    assertEquals(run.stderr, s"Error: $message\n", "one Error: line on stderr, nothing else")

  // ---------------------------------------------------------------------------------------------
  // Payload pass-through and the legacy Text bridge
  // ---------------------------------------------------------------------------------------------

  test("view --json: data equals what view --format json prints bare") {
    val simple = file("simple.xlsx")
    for
      bare <- CliHarness.run("-f", simple, "-s", "Data", "view", "A1:C4", "--format", "json")
      wrapped <- CliHarness
        .run("-f", simple, "-s", "Data", "--json", "view", "A1:C4", "--format", "json")
    yield
      assertEquals(bare.exit, 0, bare.stderr)
      assertEquals(wrapped.exit, 0, wrapped.stderr)
      val e = envelope(wrapped)
      assertEquals(e("ok"), ujson.True)
      assertEquals(e("verb"), ujson.Str("view"))
      assertEquals(e("data"), ujson.read(bare.stdout))
      assertEquals(wrapped.stderr, "")
  }

  test("view --json (markdown): the table rides as data.text with saved:null, written:false") {
    for
      text <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "view", "A1:B2")
      wrapped <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "--json", "view", "A1:B2")
    yield
      val e = envelope(wrapped)
      assertEquals(e("data")("text"), ujson.Str(text.stdout.stripSuffix("\n")))
      assertEquals(e("data")("saved"), ujson.Null)
      assertEquals(e("data")("written"), ujson.False)
  }

  test("--json never alters text-mode stdout: data.text is the text run's stdout") {
    val simple = file("simple.xlsx")
    // (global flags, verb and its arguments): --json goes between the two
    val cases: List[(List[String], List[String])] = List(
      (List("-f", simple), List("search", "Hello")),
      (List("-f", simple, "-s", "Data"), List("cell", "B4")),
      (List("-f", simple, "-s", "Data"), List("stats", "B1:B3")),
      (List("-f", simple), List("sheets")),
      (List("-f", simple, "-s", "Data"), List("bounds")),
      (Nil, List("eval", "=1+1"))
    )
    cases.foldLeft(IO.unit) { case (acc, (globals, rest)) =>
      acc *> {
        for
          text <- CliHarness.run(globals ++ rest, "")
          wrapped <- CliHarness.run(globals ++ ("--json" :: rest), "")
        yield
          assertEquals(text.exit, 0, s"$rest: ${text.stderr}")
          assertEquals(wrapped.exit, 0, s"$rest --json: ${wrapped.stderr}")
          val e = envelope(wrapped)
          e("data") match
            case obj: ujson.Obj if obj.value.contains("text") =>
              assertEquals(obj("text"), ujson.Str(text.stdout.stripSuffix("\n")), rest.toString)
            case _ => () // a typed verb: its shape is pinned by its own test below
      }
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Writes: saved / written
  // ---------------------------------------------------------------------------------------------

  test("put --json success: ok:true, data.saved is the output path, data.written:true") {
    val out = file("put-json-out.xlsx")
    CliHarness
      .run("-f", file("simple.xlsx"), "-s", "Data", "-o", out, "--json", "put", "A5", "42")
      .map { run =>
        assertEquals(run.exit, 0, run.stderr)
        val e = envelope(run)
        assertEquals(e("ok"), ujson.True)
        assertEquals(e("verb"), ujson.Str("put"))
        assertEquals(e("data")("saved"), ujson.Str(out))
        assertEquals(e("data")("written"), ujson.True)
        assert(e("data")("text").str.startsWith("Put: A5 = 42 (number)"), e("data")("text").str)
        assert(e("data")("text").str.contains(s"Saved: $out"), e("data")("text").str)
        assertEquals(run.stderr, "")
        assert(Files.size(Path.of(out)) > 0L, "the file was written")
      }
  }

  test(
    "put --json sheet-not-found: ok:false, SHEET_NOT_FOUND, exit 3, data:null, one stderr line"
  ) {
    val out = file("put-json-nope.xlsx")
    CliHarness
      .run("-f", file("simple.xlsx"), "-s", "Nope", "-o", out, "--json", "put", "A1", "1")
      .map { run =>
        assertEquals(run.exit, 3)
        val e = envelope(run)
        assertEquals(e("ok"), ujson.False)
        assertEquals(e("exitCode"), ujson.Num(3))
        assertEquals(e("verb"), ujson.Str("put"))
        assertEquals(e("data"), ujson.Null)
        assertEquals(e("error")("code"), ujson.Str("SHEET_NOT_FOUND"))
        assertEquals(
          e("error")("message"),
          ujson.Str("Sheet not found: Nope. Available: Data, Summary")
        )
        oneErrorLine(run, "Sheet not found: Nope. Available: Data, Summary")
        assert(!Files.exists(Path.of(out)), "nothing is written on a failure")
      }
  }

  test("sheets hide --json: the verb is the subcommand path and the write is recorded") {
    val out = file("hide-json-out.xlsx")
    CliHarness
      .run("-f", file("simple.xlsx"), "-o", out, "--json", "sheets", "hide", "Summary")
      .map { run =>
        assertEquals(run.exit, 0, run.stderr)
        val e = envelope(run)
        assertEquals(e("verb"), ujson.Str("sheets hide"))
        assertEquals(e("data")("saved"), ujson.Str(out))
        assertEquals(e("data")("written"), ujson.True)
      }
  }

  test("new --json: data.saved is the created file, written:true") {
    val out = fixtures().resolve("new-json.xlsx")
    CliHarness.run("--json", "new", out.toString).map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val e = envelope(run)
      assertEquals(e("verb"), ujson.Str("new"))
      assertEquals(e("data")("saved"), ujson.Str(out.toAbsolutePath.toString))
      assertEquals(e("data")("written"), ujson.True)
      assert(e("data")("text").str.startsWith("Created "), e("data")("text").str)
    }
  }

  test(
    "-i --strict recalc --json on a cyclic book: exit 1, RECALC_GATE, written:false, input intact"
  ) {
    val input = fixtures().resolve("circular.xlsx")
    for
      before <- IO.blocking(Files.readAllBytes(input))
      run <- CliHarness.run("-f", input.toString, "-i", "--strict", "--json", "recalc")
      after <- IO.blocking(Files.readAllBytes(input))
    yield
      assertEquals(run.exit, 1, run.stderr)
      val e = envelope(run)
      assertEquals(e("ok"), ujson.False)
      assertEquals(e("exitCode"), ujson.Num(1))
      assertEquals(e("verb"), ujson.Str("recalc"))
      assertEquals(e("error")("code"), ujson.Str("RECALC_GATE"))
      assertEquals(e("data")("written"), ujson.False)
      assertEquals(e("data")("saved"), ujson.Null)
      assert(e("data")("text").str.contains("NOT saved (--strict failure)"), e("data")("text").str)
      assert(e("data")("text").str.contains("STRICT FAILURE (--strict)"), e("data")("text").str)
      assertEquals(run.stderr.linesIterator.size, 1, run.stderr)
      assert(run.stderr.startsWith("Error: "), run.stderr)
      assert(java.util.Arrays.equals(before, after), "-i leaves the input untouched")
  }

  test(
    "-o --strict recalc --json on a cyclic book: exit 1 but the output is written and recorded"
  ) {
    val out = file("strict-json-out.xlsx")
    CliHarness
      .run("-f", file("circular.xlsx"), "-o", out, "--strict", "--json", "recalc")
      .map { run =>
        assertEquals(run.exit, 1, run.stderr)
        val e = envelope(run)
        assertEquals(e("error")("code"), ujson.Str("RECALC_GATE"))
        assertEquals(e("data")("saved"), ujson.Str(out))
        assertEquals(e("data")("written"), ujson.True)
        assert(e("data")("text").str.contains(s"Saved: $out"), e("data")("text").str)
        assert(Files.size(Path.of(out)) > 0L)
      }
  }

  test("view --eval --strict --json on a cyclic book: the gate has no payload, data:null") {
    CliHarness
      .run(
        "-f",
        file("circular.xlsx"),
        "-s",
        "Data",
        "--json",
        "view",
        "A1:A1",
        "--eval",
        "--strict"
      )
      .map { run =>
        assertEquals(run.exit, 1)
        val e = envelope(run)
        assertEquals(e("ok"), ujson.False)
        assertEquals(e("error")("code"), ujson.Str("RECALC_GATE"))
        assertEquals(e("data"), ujson.Null)
        assertEquals(run.stderr.linesIterator.size, 1, run.stderr)
      }
  }

  // ---------------------------------------------------------------------------------------------
  // Typed payloads
  // ---------------------------------------------------------------------------------------------

  test("sheets --json: typed [{name, index, state, dimension}]") {
    CliHarness.run("-f", file("simple.xlsx"), "--json", "sheets").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val e = envelope(run)
      assertEquals(e("verb"), ujson.Str("sheets"))
      assertEquals(
        e("data"),
        ujson.Arr(
          ujson.Obj(
            "name" -> ujson.Str("Data"),
            "index" -> ujson.Num(1),
            "state" -> ujson.Str("visible"),
            "dimension" -> ujson.Str("A1:C4")
          ),
          ujson.Obj(
            "name" -> ujson.Str("Summary"),
            "index" -> ujson.Num(2),
            "state" -> ujson.Str("visible"),
            "dimension" -> ujson.Null
          )
        )
      )
    }
  }

  test("sheets --stats --json: the full listing adds cells and formulas") {
    CliHarness.run("-f", file("simple.xlsx"), "--json", "sheets", "--stats").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val e = envelope(run)
      val data = e("data").arr
      assertEquals(data.size, 2)
      assertEquals(data(0)("name"), ujson.Str("Data"))
      assertEquals(data(0)("index"), ujson.Num(1))
      assertEquals(data(0)("state"), ujson.Str("visible"))
      assertEquals(data(0)("dimension"), ujson.Str("A1:C4"))
      assertEquals(data(0)("cells"), ujson.Num(9))
      assertEquals(data(0)("formulas"), ujson.Num(1))
      assertEquals(data(1)("dimension"), ujson.Null)
      assertEquals(data(1)("cells"), ujson.Num(0))
    }
  }

  test("names --json: typed [{name, refersTo, scope, hidden}]; [] when there are none") {
    for
      named <- CliHarness.run("-f", file("named.xlsx"), "--json", "names")
      none <- CliHarness.run("-f", file("simple.xlsx"), "--json", "names")
    yield
      assertEquals(named.exit, 0, named.stderr)
      assertEquals(
        envelope(named)("data"),
        ujson.Arr(
          ujson.Obj(
            "name" -> ujson.Str("Total"),
            "refersTo" -> ujson.Str("Data!$B$4"),
            "scope" -> ujson.Null,
            "hidden" -> ujson.False
          )
        )
      )
      assertEquals(none.exit, 0, none.stderr)
      assertEquals(envelope(none)("data"), ujson.Arr())
  }

  test("bounds --json: typed {sheet, range, dimension}") {
    CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "--json", "bounds").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(
        envelope(run)("data"),
        ujson.Obj(
          "sheet" -> ujson.Str("Data"),
          "range" -> ujson.Str("A1:C4"),
          "dimension" -> ujson.True
        )
      )
    }
  }

  test("eval --json: typed {formula, result: {type, value, formatted}, overrides}") {
    for
      constant <- CliHarness.run("--json", "eval", "=1+1")
      withRefs <- CliHarness
        .run("-f", file("simple.xlsx"), "-s", "Data", "--json", "eval", "=B1*2", "--with", "B1=5")
    yield
      assertEquals(constant.exit, 0, constant.stderr)
      val e = envelope(constant)
      assertEquals(e("verb"), ujson.Str("eval"))
      assertEquals(
        e("data"),
        ujson.Obj(
          "formula" -> ujson.Str("=1+1"),
          "result" -> ujson.Obj(
            "type" -> ujson.Str("number"),
            "value" -> ujson.Num(2),
            "formatted" -> ujson.Str("2")
          ),
          "overrides" -> ujson.Arr()
        )
      )
      assertEquals(withRefs.exit, 0, withRefs.stderr)
      val d = envelope(withRefs)("data")
      assertEquals(d("result")("value"), ujson.Num(10))
      assertEquals(d("overrides"), ujson.Arr(ujson.Str("B1=5")))
  }

  test("evala --json: typed {formula, spillRange, result, overrides} with the view JSON grid") {
    CliHarness
      .run("-f", file("simple.xlsx"), "-s", "Data", "--json", "evala", "=TRANSPOSE(A1:B2)")
      .map { run =>
        assertEquals(run.exit, 0, run.stderr)
        val e = envelope(run)
        assertEquals(e("verb"), ujson.Str("evala"))
        val d = e("data")
        assertEquals(d("formula"), ujson.Str("=TRANSPOSE(A1:B2)"))
        assertEquals(d("spillRange"), ujson.Str("Z1000:AA1001"))
        assertEquals(d("overrides"), ujson.Arr())
        assertEquals(d("result")("range"), ujson.Str("Z1000:AA1001"))
        val rows = d("result")("rows").arr
        assertEquals(rows.size, 2)
        assertEquals(rows(0)("cells")(0)("value"), ujson.Str("Hello"))
        assertEquals(rows(1)("cells")(0)("value"), ujson.Num(10))
      }
  }

  test("functions --json: typed [{name}]") {
    CliHarness.run("--json", "functions").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val e = envelope(run)
      assertEquals(e("verb"), ujson.Str("functions"))
      val names = e("data").arr.map(_("name").str)
      assert(names.contains("SUM"), names.mkString(", "))
      assert(names.size > 100, names.size.toString)
    }
  }

  test("rasterizers --json: typed {backends: [{name, status, note}], anyAvailable}") {
    CliHarness.run("--json", "rasterizers").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val d = envelope(run)("data")
      val backends = d("backends").arr
      assertEquals(
        backends.map(_("name").str).toList,
        List("batik", "cairosvg", "rsvg-convert", "resvg", "imagemagick")
      )
      backends.foreach { b =>
        assert(Set("available", "missing", "broken", "unavailable").contains(b("status").str))
        assert(b("note").strOpt.isDefined, b("note").toString)
      }
      assert(d("anyAvailable").boolOpt.isDefined, d("anyAvailable").toString)
    }
  }

  test("batch --dry-run --json: data is {ops: [{index, op, summary}], warnings}") {
    val stdin =
      """[{"op":"put","ref":"A1","value":"$1,234.56"},{"op":"merge","range":"A1:B1"},{"op":"put","ref":"A2","value":1,"bogus":true}]"""
    CliHarness.run(List("--json", "batch", "--dry-run", "-"), stdin).map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val e = envelope(run)
      assertEquals(e("verb"), ujson.Str("batch"))
      val d = e("data")
      assertEquals(
        d("ops"),
        ujson.Arr(
          ujson.Obj(
            "index" -> ujson.Num(0),
            "op" -> ujson.Str("PUT"),
            "summary" -> ujson.Str("PUT A1 = Number(1234.56) (Currency)")
          ),
          ujson.Obj(
            "index" -> ujson.Num(1),
            "op" -> ujson.Str("MERGE"),
            "summary" -> ujson.Str("MERGE A1:B1")
          ),
          ujson.Obj(
            "index" -> ujson.Num(2),
            "op" -> ujson.Str("PUT"),
            "summary" -> ujson.Str("PUT A2 = Number(1.0)")
          )
        )
      )
      val warnings = d("warnings").arr.map(_.str)
      assertEquals(warnings.size, 1, warnings.mkString("; "))
      assert(warnings(0).contains("bogus"), warnings(0))
      assertEquals(run.stderr, "", "dry-run parse warnings are data in JSON mode, not stderr")
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Findings, gates, failures
  // ---------------------------------------------------------------------------------------------

  test(
    "diff --json: differences are ok:false / DIFFERENCES_FOUND / exit 1 with the report as data"
  ) {
    val simple = file("simple.xlsx")
    for
      differs <- CliHarness.run("-f", simple, "--json", "diff", "-g", file("changed.xlsx"))
      typed <- CliHarness
        .run("-f", simple, "--json", "diff", "-g", file("changed.xlsx"), "--format", "json")
      same <- CliHarness.run("-f", simple, "--json", "diff", "-g", file("simple-copy.xlsx"))
    yield
      assertEquals(differs.exit, 1)
      val e = envelope(differs)
      assertEquals(e("verb"), ujson.Str("diff"))
      assertEquals(e("error")("code"), ujson.Str("DIFFERENCES_FOUND"))
      assert(e("data")("text").str.startsWith("Comparing "), e("data")("text").str)
      assertEquals(differs.stderr.linesIterator.size, 1, differs.stderr)

      assertEquals(typed.exit, 1)
      assertEquals(envelope(typed)("data")("identical"), ujson.False)

      assertEquals(same.exit, 0, same.stderr)
      val s = envelope(same)
      assertEquals(s("ok"), ujson.True)
      assert(s("data")("text").str.contains("identical"), s("data")("text").str)
  }

  test("lint --json: clean is ok:true; --format json rides as data") {
    val simple = file("simple.xlsx")
    for
      text <- CliHarness.run("--json", "lint", simple)
      typed <- CliHarness.run("--json", "lint", simple, "--format", "json")
    yield
      assertEquals(text.exit, 0, text.stderr)
      val e = envelope(text)
      assertEquals(e("verb"), ujson.Str("lint"))
      assertEquals(e("data")("text"), ujson.Str(s"$simple: clean (no findings)"))
      assertEquals(typed.exit, 0, typed.stderr)
      assertEquals(envelope(typed)("data")("clean"), ujson.True)
  }

  test("a missing input file is IO_READ, exit 3, with the file in error.location") {
    CliHarness.run("-f", file("nope.xlsx"), "--json", "sheets").map { run =>
      assertEquals(run.exit, 3)
      val e = envelope(run)
      assertEquals(e("error")("code"), ujson.Str("IO_READ"))
      assertEquals(e("error")("location")("file"), ujson.Str(file("nope.xlsx")))
      assertEquals(e("data"), ujson.Null)
      assertEquals(run.stderr.linesIterator.size, 1, run.stderr)
    }
  }

  test("--json with a decline parse error still yields an envelope (USAGE, exit 2)") {
    for
      unknown <- CliHarness.run("frob", "--json")
      flagAfterVerb <- CliHarness
        .run("-f", file("simple.xlsx"), "--json", "view", "A1:B2", "-s", "Data")
      noArgs <- CliHarness.run("--json")
    yield
      assertEquals(unknown.exit, 2)
      val u = envelope(unknown)
      assertEquals(u("ok"), ujson.False)
      assertEquals(u("exitCode"), ujson.Num(2))
      assertEquals(u("error")("code"), ujson.Str("USAGE"))
      assertEquals(u("error")("message"), ujson.Str("Unexpected argument: frob"))
      assertEquals(u("data"), ujson.Null)
      oneErrorLine(unknown, "Unexpected argument: frob")

      assertEquals(flagAfterVerb.exit, 2)
      val f = envelope(flagAfterVerb)
      assertEquals(f("error")("code"), ujson.Str("USAGE"))
      assertEquals(f("verb"), ujson.Str("view"))
      oneErrorLine(flagAfterVerb, "Unexpected option: -s")

      assertEquals(noArgs.exit, 2)
      assertEquals(envelope(noArgs)("error")("code"), ujson.Str("USAGE"))
  }

  test("`--` ends the --json scan: a literal --json after it keeps the text-mode usage error") {
    CliHarness.run("-f", file("simple.xlsx"), "search", "--", "--json", "extra").map { run =>
      assertEquals(run.exit, 2)
      assertEquals(run.stdout, "", "no envelope: that --json was data, not a flag")
      assert(run.stderr.startsWith("Unexpected argument: extra"), run.stderr)
      assert(run.stderr.contains("Usage:"), run.stderr)
    }
  }

  test("verbOf skips the values of global options and stops at --") {
    assertEquals(Cli.verbOf(List("-s", "view", "--json", "cell", "A1")), "cell")
    assertEquals(Cli.verbOf(List("--file=view", "--json", "cell")), "cell")
    assertEquals(Cli.verbOf(List("-f", "x.xlsx", "search", "--", "view")), "search")
    assertEquals(Cli.verbOf(List("-s")), "")
    assertEquals(Cli.verbOf(List("-f", "view")), "")
    CliHarness.run("-s", "view", "--json", "cell", "A1").map { run =>
      assertEquals(run.exit, 2)
      val e = envelope(run)
      assertEquals(e("error")("code"), ujson.Str("USAGE"))
      assertEquals(e("verb"), ujson.Str("cell"))
    }
  }

  test("bounds on a missing input file is IO_READ, exit 3, like every other verb") {
    for
      dimension <- CliHarness.run("-f", file("nope.xlsx"), "-s", "Data", "--json", "bounds")
      scan <- CliHarness.run("-f", file("nope.xlsx"), "-s", "Data", "--json", "bounds", "--scan")
      text <- CliHarness.run("-f", file("nope.xlsx"), "-s", "Data", "bounds")
    yield
      assertEquals(dimension.exit, 3)
      assertEquals(envelope(dimension)("error")("code"), ujson.Str("IO_READ"))
      assertEquals(scan.exit, 3)
      assertEquals(envelope(scan)("error")("code"), ujson.Str("IO_READ"))
      assertEquals(text.exit, 3)
      assert(text.stderr.contains("  code: IO_READ"), text.stderr)
  }

  test("Cli.verbs is the parser's own verb list, so a usage envelope's verb cannot drift") {
    CliHarness.run().map { run =>
      assertEquals(run.exit, 2)
      val listed = run.stderr.linesIterator.toVector.headOption.getOrElse("")
      val names = listed
        .stripPrefix("Missing expected command (")
        .takeWhile(_ != ')')
        .split(" or ")
        .toVector
      assertEquals(names, Cli.verbs)
    }
  }

  test("-i with -o under --json is a USAGE envelope, exit 2, nothing written") {
    val out = file("both-json.xlsx")
    CliHarness
      .run("-f", file("simple.xlsx"), "-s", "Data", "-o", out, "-i", "--json", "put", "A1", "1")
      .map { run =>
        assertEquals(run.exit, 2)
        val e = envelope(run)
        assertEquals(e("error")("code"), ujson.Str("USAGE"))
        assertEquals(e("verb"), ujson.Str("put"))
        assert(!Files.exists(Path.of(out)))
      }
  }

  // ---------------------------------------------------------------------------------------------
  // Warnings
  // ---------------------------------------------------------------------------------------------

  test(
    "--stream view --format csv --limit: TRUNCATED in warnings[], stderr empty; text mode warns"
  ) {
    val notice = "… showing 2 of 4 rows (use --limit to raise; --limit 0 = no limit)"
    val common = List("-f", file("simple.xlsx"), "-s", "Data", "--stream")
    val verb = List("view", "A1:C4", "--format", "csv", "--limit", "2")
    for
      wrapped <- CliHarness.run(common ++ ("--json" :: verb), "")
      text <- CliHarness.run(common ++ verb, "")
    yield
      assertEquals(wrapped.exit, 0, wrapped.stderr)
      val e = envelope(wrapped)
      assertEquals(e("ok"), ujson.True)
      assertEquals(
        e("warnings"),
        ujson.Arr(ujson.Obj("code" -> ujson.Str("TRUNCATED"), "message" -> ujson.Str(notice)))
      )
      assert(e("data")("text").str.startsWith("Hello,10,"), e("data")("text").str)
      assertEquals(wrapped.stderr, "", "the notice is in the envelope, not on stderr")
      assertEquals(text.exit, 0, text.stderr)
      assertEquals(
        text.stderr,
        s"Warning[TRUNCATED]: $notice\n",
        "same channel as the in-memory view"
      )
  }

  test(
    "view --eval without --strict on a cyclic book: ok:true, one EVAL_FAILED warning, stderr empty"
  ) {
    for
      wrapped <- CliHarness
        .run("-f", file("circular.xlsx"), "-s", "Data", "--json", "view", "A1:A1", "--eval")
      text <- CliHarness.run("-f", file("circular.xlsx"), "-s", "Data", "view", "A1:A1", "--eval")
    yield
      assertEquals(wrapped.exit, 0, wrapped.stderr)
      val e = envelope(wrapped)
      assertEquals(e("ok"), ujson.True)
      val warnings = e("warnings").arr
      assertEquals(warnings.size, 1, warnings.toString)
      assertEquals(warnings.headOption.map(_("code")), Some(ujson.Str("EVAL_FAILED")))
      assert(
        warnings.headOption.exists(_("message").str.startsWith("Formula evaluation failed: ")),
        warnings.toString
      )
      assert(e("data")("text").str.contains("| A"), e("data")("text").str)
      assertEquals(wrapped.stderr, "")
      assertEquals(text.exit, 0)
      assert(
        text.stderr.startsWith("Warning[EVAL_FAILED]: Formula evaluation failed: "),
        text.stderr
      )
      assertEquals(text.stderr.linesIterator.size, 1, text.stderr)
  }

  test(
    "view --format csv --limit: TRUNCATED rides in warnings[]; text mode prints Warning[TRUNCATED]"
  ) {
    val notice = "… showing 2 of 4 rows (use --limit to raise; --limit 0 = no limit)"
    for
      wrapped <- CliHarness.run(
        "-f",
        file("simple.xlsx"),
        "-s",
        "Data",
        "--json",
        "view",
        "A1:C4",
        "--format",
        "csv",
        "--limit",
        "2"
      )
      text <- CliHarness
        .run(
          "-f",
          file("simple.xlsx"),
          "-s",
          "Data",
          "view",
          "A1:C4",
          "--format",
          "csv",
          "--limit",
          "2"
        )
    yield
      assertEquals(wrapped.exit, 0, wrapped.stderr)
      val e = envelope(wrapped)
      assertEquals(
        e("warnings"),
        ujson.Arr(ujson.Obj("code" -> ujson.Str("TRUNCATED"), "message" -> ujson.Str(notice)))
      )
      assertEquals(e("data")("text"), ujson.Str("Hello,10,1/15/24\nWorld,20,6/30/24"))
      assertEquals(wrapped.stderr, "", "warnings are in the envelope, not on stderr")
      assertEquals(text.exit, 0, text.stderr)
      assertEquals(text.stderr, s"Warning[TRUNCATED]: $notice\n")
      assertEquals(text.stdout, "Hello,10,1/15/24\nWorld,20,6/30/24\n")
  }
