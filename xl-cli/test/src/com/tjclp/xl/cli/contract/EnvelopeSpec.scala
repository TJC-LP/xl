package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.cli.{Cli, CliIO}

/**
 * The global `--json` flag end to end (ADR-017 §2.4), through the in-process harness: every verb,
 * success or failure, prints exactly one envelope on stdout; `--format json` payloads ride inside
 * it unchanged (as text — no number is re-parsed); a pass-through verb with no `--format` uses its
 * JSON payload format; failures put one `Error: <message>` line on stderr while signals (findings,
 * gates) print nothing there; and every envelope validates against `envelope.schema.json`.
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
    // reformat, not read-then-write: the comparison must not round a 17-digit integer either
    assertEquals(
      run.stdout,
      ujson.reformat(run.stdout, indent = 2) + "\n",
      "stdout is the envelope alone"
    )
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

  /**
   * The `data` block of an envelope as TEXT: `bare` reformatted to the envelope's indentation and
   * depth — what [[Render.json]] splices — so the comparison is lexeme for lexeme, never through a
   * ujson tree.
   */
  private def dataBlock(bare: String): String =
    "  \"data\": " + ujson.reformat(bare, indent = 2).replace("\n", "\n  ") + ",\n"

  test("view --json: data equals what view --format json prints bare, lexeme for lexeme") {
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
      assert(wrapped.stdout.contains(dataBlock(bare.stdout)), wrapped.stdout)
      assertEquals(wrapped.stderr, "")
  }

  test("view --json without --format: data is the JSON payload, as with --format json") {
    val simple = file("simple.xlsx")
    for
      bare <- CliHarness.run("-f", simple, "-s", "Data", "view", "A1:B2", "--format", "json")
      wrapped <- CliHarness.run("-f", simple, "-s", "Data", "--json", "view", "A1:B2")
      text <- CliHarness.run("-f", simple, "-s", "Data", "view", "A1:B2")
    yield
      val e = envelope(wrapped)
      assertEquals(e("data")("sheet"), ujson.Str("Data"))
      assertEquals(e("data")("range"), ujson.Str("A1:B2"))
      assert(wrapped.stdout.contains(dataBlock(bare.stdout)), wrapped.stdout)
      // and text mode still defaults to the markdown table
      assert(text.stdout.startsWith("|   | A"), text.stdout)
  }

  test("view --json --format csv: an explicit text format still rides as data.text") {
    for
      text <- CliHarness
        .run("-f", file("simple.xlsx"), "-s", "Data", "view", "A1:B2", "--format", "csv")
      wrapped <- CliHarness
        .run("-f", file("simple.xlsx"), "-s", "Data", "--json", "view", "A1:B2", "--format", "csv")
      markdown <- CliHarness
        .run("-f", file("simple.xlsx"), "-s", "Data", "--json", "view", "A1:B2", "--format", "md")
    yield
      val e = envelope(wrapped)
      assertEquals(e("data")("text"), ujson.Str(text.stdout.stripSuffix("\n")))
      assertEquals(e("data")("saved"), ujson.Null)
      assertEquals(e("data")("written"), ujson.False)
      assert(envelope(markdown)("data")("text").str.startsWith("|   | A"), markdown.stdout)
  }

  test("a 17-digit integer keeps every digit in data: view, eval and evala") {
    val precise = file("precise.xlsx")
    val digits = "12345678901234567"
    for
      bare <- CliHarness.run("-f", precise, "view", "A1:A1", "--format", "json")
      view <- CliHarness.run("-f", precise, "--json", "view", "A1:A1")
      eval <- CliHarness.run("-f", precise, "--json", "eval", "=A1")
      constant <- CliHarness.run("--json", "eval", s"=$digits")
      evala <- CliHarness.run("-f", precise, "--json", "evala", "=TRANSPOSE(A1:A1)")
    yield
      assert(bare.stdout.contains(s"\"value\": $digits,"), bare.stdout)
      List("view" -> view, "eval" -> eval, "eval constant" -> constant, "evala" -> evala).foreach {
        (verb, run) =>
          assertEquals(run.exit, 0, s"$verb: ${run.stderr}")
          envelope(run)
          assert(run.stdout.contains(s"\"value\": $digits,"), s"$verb:\n${run.stdout}")
          assert(!run.stdout.contains("12345678901234568"), s"$verb rounded:\n${run.stdout}")
      }
      assertEquals(envelope(eval)("data")("result")("type"), ujson.Str("number"))
      assertEquals(envelope(evala)("data")("spillRange"), ujson.Str("Z1000:Z1000"))
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
      assertEquals(run.stderr, "", "a gate keeps its report as data and prints no Error: line")
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

  test(
    "batch --dry-run --json: data is {ops: [{index, op, summary}]} with 1-based index; parse warnings ride in warnings[]"
  ) {
    val stdin =
      """[{"op":"put","ref":"A1","value":"$1,234.56"},{"op":"merge","range":"A1:B1"},{"op":"put","ref":"A2","value":1,"bogus":true}]"""
    CliHarness.run(List("--json", "batch", "--dry-run", "-"), stdin).map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val e = envelope(run)
      assertEquals(e("verb"), ujson.Str("batch"))
      val d = e("data")
      // `index` is the op's 1-based position — the number every `Object N` message and
      // `location.opIndex` carry (ADR-017 §2.6) — and `data` holds nothing but the ops
      assertEquals(
        d,
        ujson.Obj(
          "ops" -> ujson.Arr(
            ujson.Obj(
              "index" -> ujson.Num(1),
              "op" -> ujson.Str("PUT"),
              "summary" -> ujson.Str("PUT A1 = Number(1234.56) (Currency)")
            ),
            ujson.Obj(
              "index" -> ujson.Num(2),
              "op" -> ujson.Str("MERGE"),
              "summary" -> ujson.Str("MERGE A1:B1")
            ),
            ujson.Obj(
              "index" -> ujson.Num(3),
              "op" -> ujson.Str("PUT"),
              "summary" -> ujson.Str("PUT A2 = Number(1.0)")
            )
          )
        )
      )
      // The parse warning is an envelope warning like any other, located at its op
      val warnings = e("warnings").arr
      assertEquals(warnings.size, 1, warnings.mkString("; "))
      assertEquals(warnings(0)("code"), ujson.Str("UNKNOWN_PROPERTY"))
      assert(warnings(0)("message").str.contains("bogus"), warnings(0)("message").str)
      assertEquals(warnings(0)("location")("opIndex"), ujson.Num(3))
      assertEquals(run.stderr, "", "under --json the warning rides in the envelope, not on stderr")
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
      // no --format under --json: the JSON report, exactly as --format json gives it
      assertEquals(e("data")("identical"), ujson.False)
      assertEquals(e("data"), envelope(typed)("data"))
      assertEquals(differs.stderr, "", "findings are the report, not an Error: line")

      assertEquals(typed.exit, 1)
      assertEquals(envelope(typed)("data")("identical"), ujson.False)
      assertEquals(typed.stderr, "")

      assertEquals(same.exit, 0, same.stderr)
      val s = envelope(same)
      assertEquals(s("ok"), ujson.True)
      assertEquals(s("data")("identical"), ujson.True)
  }

  test("diff --json --format markdown: the explicit text format rides as data.text") {
    CliHarness
      .run(
        "-f",
        file("simple.xlsx"),
        "--json",
        "diff",
        "-g",
        file("changed.xlsx"),
        "--format",
        "markdown"
      )
      .map { run =>
        assertEquals(run.exit, 1)
        val e = envelope(run)
        assert(e("data")("text").str.startsWith("Comparing "), e("data")("text").str)
        assertEquals(run.stderr, "")
      }
  }

  test("lint --json: clean is ok:true with the JSON report as data; --format text rides as text") {
    val simple = file("simple.xlsx")
    for
      plain <- CliHarness.run("--json", "lint", simple)
      typed <- CliHarness.run("--json", "lint", simple, "--format", "json")
      text <- CliHarness.run("--json", "lint", simple, "--format", "text")
    yield
      assertEquals(plain.exit, 0, plain.stderr)
      val e = envelope(plain)
      assertEquals(e("verb"), ujson.Str("lint"))
      assertEquals(e("data")("clean"), ujson.True)
      assertEquals(e("data"), envelope(typed)("data"))
      assertEquals(typed.exit, 0, typed.stderr)
      assertEquals(envelope(text)("data")("text"), ujson.Str(s"$simple: clean (no findings)"))
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

  test("--json with a usage failure still yields an envelope (UNKNOWN_VERB / USAGE, exit 2)") {
    for
      unknown <- CliHarness.run("frob", "--json")
      flagAfterVerb <- CliHarness
        .run("-f", file("simple.xlsx"), "--json", "view", "A1:B2", "-s", "Data")
      // W2.4: `view` no longer needs a range (it defaults to the used range); `cell` still does
      missingArg <- CliHarness.run("-f", file("simple.xlsx"), "-s", "Data", "--json", "cell")
      noArgs <- CliHarness.run("--json")
    yield
      assertEquals(unknown.exit, 2)
      val u = envelope(unknown)
      assertEquals(u("ok"), ujson.False)
      assertEquals(u("exitCode"), ujson.Num(2))
      assertEquals(u("error")("code"), ujson.Str("UNKNOWN_VERB"))
      assertEquals(u("error")("message"), ujson.Str("unknown verb 'frob'"))
      assertEquals(u("data"), ujson.Null)
      oneErrorLine(unknown, "unknown verb 'frob'")

      // ADR-017 §2.2: a global after the verb is hoisted, so this is now a plain success
      assertEquals(flagAfterVerb.exit, 0, flagAfterVerb.stderr)
      val f = envelope(flagAfterVerb)
      assertEquals(f("ok"), ujson.True)
      assertEquals(f("verb"), ujson.Str("view"))

      assertEquals(missingArg.exit, 2)
      val m = envelope(missingArg)
      assertEquals(m("error")("code"), ujson.Str("USAGE"))
      assertEquals(m("verb"), ujson.Str("cell"))
      assertEquals(m("error")("hint"), ujson.Str("run `xl cell --help` for the usage"))
      assertEquals(missingArg.stderr.linesIterator.size, 1, missingArg.stderr)

      assertEquals(noArgs.exit, 2)
      assertEquals(envelope(noArgs)("error")("code"), ujson.Str("USAGE"))
  }

  test("GH-620: `--help` is a result — stdout in text mode, `data.usage` under --json") {
    for
      text <- CliHarness.run("delete-rows", "--help")
      json <- CliHarness.run("--json", "delete-rows", "--help")
      top <- CliHarness.run("--json", "--help")
    yield
      assertEquals(text.exit, 0)
      assertEquals(text.stderr, "")
      assert(text.stdout.startsWith("Usage:"), text.stdout)
      assertEquals(json.exit, 0)
      assertEquals(json.stderr, "")
      val e = envelope(json)
      assertEquals(e("ok"), ujson.True)
      assertEquals(e("verb"), ujson.Str("delete-rows"))
      assertEquals(e("error"), ujson.Null)
      assertEquals(e("data")("usage"), ujson.Str(text.stdout.stripSuffix("\n")))
      assert(e("data")("usage").str.startsWith("Usage:"), e("data")("usage").str)
      val t = envelope(top)
      assertEquals(t("ok"), ujson.True)
      assertEquals(t("verb"), ujson.Str(""))
      assert(t("data")("usage").str.contains("Exit codes:"), t("data")("usage").str)
  }

  test("`--` ends the --json scan: a literal --json after it keeps the text-mode usage error") {
    CliHarness.run("-f", file("simple.xlsx"), "search", "--", "--json", "extra").map { run =>
      assertEquals(run.exit, 2)
      assertEquals(run.stdout, "", "no envelope: that --json was data, not a flag")
      assert(run.stderr.startsWith("Error: Unexpected argument: extra"), run.stderr)
      assert(run.stderr.contains(Cli.usage), run.stderr)
      assert(run.stderr.contains("  code: USAGE"), run.stderr)
    }
  }

  test("Argv.verbOf names the verb a usage envelope reports; an unknown one is not a verb") {
    assertEquals(Argv.verbOf(List("-s", "view", "--json", "cell", "A1")), Some("cell"))
    assertEquals(Argv.verbOf(List("--file=view", "--json", "cell")), Some("cell"))
    assertEquals(Argv.verbOf(List("-f", "x.xlsx", "search", "--", "view")), Some("search"))
    assertEquals(Argv.verbOf(List("-s")), None)
    assertEquals(Argv.verbOf(List("-f", "view")), None)
    for
      missingFile <- CliHarness.run("-s", "view", "--json", "cell", "A1")
      unknown <- CliHarness.run("--json", "frob", "A1")
    yield
      assertEquals(missingFile.exit, 2)
      val e = envelope(missingFile)
      assertEquals(e("error")("code"), ujson.Str("USAGE"))
      assertEquals(e("verb"), ujson.Str("cell"))
      assertEquals(envelope(unknown)("verb"), ujson.Str(""))
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

  test("Argv.verbs is the parser's own verb list, so a usage envelope's verb cannot drift") {
    // decline names one subcommand per usage line: `xl [globals] <verb>`
    val usage = Cli.command(CliIO.system).showHelp.linesIterator.toVector
    val names = usage.takeWhile(_.trim.nonEmpty).drop(1).map(_.trim.split(" ").last).distinct
    assertEquals(names, Argv.verbs)
    // and the no-argument run still reports usage (exit 2) without listing them all
    CliHarness.run().map { run =>
      assertEquals(run.exit, 2)
      assert(run.stderr.contains(Cli.usage), run.stderr)
      assert(!run.stderr.contains("Missing expected command ("), run.stderr)
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
      // no --format under --json: the JSON grid, with the formula's (uncached) cell
      assertEquals(e("data")("rows")(0)("cells")(0)("formula"), ujson.Str("=A1+1"))
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
