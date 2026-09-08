package com.tjclp.xl.cli.contract

import munit.CatsEffectSuite

import com.tjclp.xl.cli.{BuildInfo, Cli, CliIO, Main}
import com.tjclp.xl.cli.batch.OpRegistry
import com.tjclp.xl.cli.read.Capability

/**
 * `xl schema` (ADR-017 §2.13): the machine-readable contract. The verb table is a literal in
 * [[Schema]] until Wave 2 derives it from the command registry, so this suite pins it to the
 * parser: its heads are exactly the subcommand names decline renders (== [[Argv.verbs]]), every
 * batch twin names an [[OpRegistry]] op that points back, every global is one of [[Argv.globals]],
 * and every error code maps to an exit code. Then the three surfaces through the harness:
 * `schema [--json]`, `batch --schema`, `batch --help`, `functions --json`.
 */
class SchemaSpec extends CatsEffectSuite:

  private val version = BuildInfo.version

  private def envelope(run: CliRun): ujson.Obj =
    val parsed = ujson.read(run.stdout)
    EnvelopeSchema.assertValid(parsed)
    assertEquals(run.stdout, ujson.write(parsed, indent = 2) + "\n", "stdout is the envelope alone")
    parsed.obj

  // ---------------------------------------------------------------------------------------------
  // The verb table against the parser
  // ---------------------------------------------------------------------------------------------

  test("verb heads are exactly the subcommand names decline renders, in usage order") {
    val usage = Cli.command(CliIO.system).showHelp.linesIterator.toVector
    val rendered = usage.takeWhile(_.trim.nonEmpty).drop(1).map(_.trim.split(" ").last).distinct
    assertEquals(Schema.verbs.map(_.path.headOption).flatten.distinct, rendered)
    assertEquals(rendered, Argv.verbs)
    assert(Argv.verbs.contains("schema"), Argv.verbs.toString)
  }

  test("verb paths are unique, non-empty, and nested only under sheets/name/cf/chart") {
    val paths = Schema.verbs.map(_.path)
    assertEquals(paths.distinct.size, paths.size)
    assert(paths.forall(_.nonEmpty))
    val nested = Schema.verbs.filter(_.path.size > 1)
    assertEquals(
      nested.map(_.path.mkString(" ")).toSet,
      Set("sheets hide", "sheets show", "name add", "name rm", "cf add", "cf list", "chart add")
    )
    assert(Schema.verbs.forall(_.path.size <= 2))
    assert(Schema.verbs.forall(_.summary.trim.nonEmpty), "every verb has a one-line summary")
    assert(Schema.verbs.forall(v => !v.summary.contains('\n')), "summaries are one line")
  }

  test("every batchTwin names an OpRegistry op whose cliVerb points back at the verb") {
    Schema.verbs.foreach { verb =>
      verb.batchTwin.foreach { twin =>
        val spec = OpRegistry.find(twin)
        assert(spec.exists(_.name == twin), s"${verb.path.mkString(" ")}: no batch op '$twin'")
        assertEquals(
          spec.flatMap(_.cliVerb),
          Some(verb.path.mkString(" ")),
          s"$twin's cliVerb should name ${verb.path.mkString(" ")}"
        )
      }
    }
    // and the other direction: every op that names a CLI verb has that verb in the table
    OpRegistry.all.flatMap(_.cliVerb).foreach { cliVerb =>
      assert(Schema.verbs.exists(_.path.mkString(" ") == cliVerb), s"no verb '$cliVerb'")
    }
  }

  test("every verb can exit 0, 2 and 3; exit 1 only for gates and findings") {
    Schema.verbs.foreach { verb =>
      val clue = verb.path.mkString(" ")
      assert(verb.exit.toSet.subsetOf(Set(0, 1, 2, 3)), clue)
      assert(Set(0, 2, 3).subsetOf(verb.exit.toSet), clue)
      assertEquals(verb.exit, verb.exit.sorted, clue)
    }
    val gates = Schema.verbs.filter(_.exit.contains(1)).map(_.path.mkString(" ")).toSet
    assert(Set("diff", "lint", "audit", "view", "recalc", "put", "batch").subsetOf(gates), gates)
    assert(!gates.contains("sheets") && !gates.contains("style"), gates.toString)
  }

  test("needs: -f for every verb but the standalone ones; -o exactly for the writes") {
    val byPath = Schema.verbs.map(v => v.path.mkString(" ") -> v).toMap
    val standalone = Set("rasterizers", "functions", "schema", "new", "eval")
    Schema.verbs.foreach { v =>
      val path = v.path.mkString(" ")
      assertEquals(v.needs.file, !standalone.contains(path), path)
    }
    assertEquals(byPath.get("view").map(_.needs), Some(Needs(true, true, false, true)))
    assertEquals(byPath.get("put").map(_.needs), Some(Needs(true, true, true, true)))
    assertEquals(byPath.get("sheets hide").map(_.needs), Some(Needs(true, false, true, false)))
    assertEquals(byPath.get("describe").map(_.needs), Some(Needs(true, false, false, true)))
    assertEquals(byPath.get("names").map(_.needs), Some(Needs(true, false, false, true)))
    assertEquals(byPath.get("lint").map(_.needs), Some(Needs(true, false, false, true)))
    assertEquals(byPath.get("schema").map(_.needs), Some(Needs(false, false, false, false)))
  }

  test("needs.streaming is exactly Main's streaming dispatch set (O(1)-memory verbs)") {
    // A verb streams only when Main routes it to the streaming reader/writer or a metadata fast
    // path; every other write verb accepts --stream but loads the book, so it must not claim it.
    assertEquals(
      Schema.verbs.filter(_.needs.streaming).map(_.verb).toSet,
      Main.streamingVerbs
    )
    assert(!Schema.verbs.exists(v => v.verb == "sheets hide" && v.needs.streaming))
    assert(!Schema.verbs.exists(v => v.verb == "merge" && v.needs.streaming))
    assert(Schema.verbs.exists(v => v.verb == "sheets" && v.needs.streaming))
    assert(Schema.verbs.exists(v => v.verb == "names" && v.needs.streaming))
    assert(Schema.verbs.exists(v => v.verb == "lint" && v.needs.streaming))
  }

  test("stream: every verb is o1, backend or refused, from one table (GH-638)") {
    val byVerb = Schema.verbs.map(v => v.verb -> v.stream).toMap
    // o1 ⇔ needs.streaming
    Schema.verbs.foreach { v =>
      assertEquals(v.stream == StreamSupport.O1, v.needs.streaming, v.verb)
    }
    val refused = Schema.verbs.collect { case v if v.stream.name == "refused" => v.verb }.toSet
    assertEquals(
      refused,
      Set("rasterizers", "functions", "schema", "new", "diff", "eval", "evala", "audit", "deps")
    )
    // every verb that reads no workbook refuses the flag; every write accepts it or streams
    Schema.verbs.filterNot(_.needs.file).filterNot(_.verb == "eval").foreach { v =>
      assertEquals(v.stream.name, "refused", v.verb)
    }
    Schema.verbs.filter(_.needs.output).foreach { v =>
      assert(v.stream.name != "refused", s"${v.verb} is a write: backend or o1, never refused")
    }
    assertEquals(byVerb.get("merge"), Some(StreamSupport.BackendOnly))
    assertEquals(byVerb.get("sheets hide"), Some(StreamSupport.BackendOnly))
    assertEquals(byVerb.get("view"), Some(StreamSupport.O1))
    assertEquals(
      byVerb.get("audit"),
      Some(StreamSupport.Refused("the analysis needs the whole workbook"))
    )
    // the flags an O(1) verb's own handler refuses under --stream, as the table publishes them
    val refusedWith = Schema.verbs.map(v => v.verb -> v.refusedWith).toMap
    assertEquals(
      refusedWith.get("view"),
      Some(Vector("--eval", "--format html/svg/png/jpeg/webp/pdf"))
    )
    assertEquals(refusedWith.get("sheets"), Some(Vector("--stats")))
    assertEquals(refusedWith.get("describe"), Some(Vector("--full")))
    assertEquals(refusedWith.get("put"), Some(Vector("--csv", "--strict")))
    assertEquals(refusedWith.get("batch"), Some(Vector("--strict")))
    assertEquals(refusedWith.get("search"), Some(Vector.empty))
    Schema.verbs.filter(_.refusedWith.nonEmpty).foreach { v =>
      assertEquals(v.stream, StreamSupport.O1, s"${v.verb} refuses flags but is not o1")
    }
  }

  test("streamRefusal: a head every form of which refuses --stream, with the alternative as hint") {
    val audit = Schema.streamRefusal("audit")
    assertEquals(
      audit.map(_.message),
      Some("audit is not supported with --stream (the analysis needs the whole workbook)")
    )
    assertEquals(audit.map(_.code), Some(ErrorCode.UNSUPPORTED_IN_STREAM))
    assertEquals(
      audit.flatMap(_.hint),
      Some("omit --stream; use --max-size <MB> to load a large file in memory")
    )
    assertEquals(audit.map(_.exitCode.code), Some(2))
    // a verb with no file has no in-memory alternative to name
    assertEquals(Schema.streamRefusal("schema").flatMap(_.hint), Some("omit --stream"))
    assertEquals(
      Schema.streamRefusal("new").map(_.message),
      Some("new is not supported with --stream (it writes a new file and reads none)")
    )
    // heads with a form that takes the flag are never refused here
    Vector("sheets", "names", "lint", "view", "batch", "name", "cf", "merge", "nope").foreach {
      head => assertEquals(Schema.streamRefusal(head), None, head)
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Globals, exit codes, error codes
  // ---------------------------------------------------------------------------------------------

  test("globals cover Argv.globals exactly, takesValue from Argv") {
    val names = Schema.globals.flatMap(g => g.name :: g.short.toList)
    assertEquals(names.toSet, Argv.globals.keySet)
    assertEquals(names.distinct.size, names.size)
    Schema.globals.foreach { g =>
      g.short.foreach(s => assertEquals(Argv.globals.get(s), Argv.globals.get(g.name), g.name))
      assert(g.doc.trim.nonEmpty, g.name)
    }
  }

  test("the exit-code table is ExitCodes' four codes with a meaning each") {
    assertEquals(
      Schema.exitCodes.map(_.code),
      Vector(ExitCodes.ok, ExitCodes.signal, ExitCodes.usage, ExitCodes.failed).map(_.code)
    )
    assert(Schema.exitCodes.forall(_.meaning.trim.nonEmpty))
  }

  test("every error code maps to an exit code 1..3, once") {
    assertEquals(ErrorCode.all.distinct, ErrorCode.all)
    ErrorCode.all.foreach { code =>
      assert(Set(1, 2, 3).contains(ExitCodes.forCode(code).code), code)
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Schema.json
  // ---------------------------------------------------------------------------------------------

  test("Schema.json has exactly the ten keys, in order, and the version") {
    val json = Schema.json(version)
    assertEquals(
      json.value.keys.toList,
      List(
        "version",
        "exitCodes",
        "errorCodes",
        "warningCodes",
        "globals",
        "verbs",
        "capabilities",
        "batchOps",
        "functions",
        "envelope"
      )
    )
    // W2.4: the read-source capability table — every capability, in memory and streaming
    assertEquals(json("capabilities"), Capability.json)
    val capabilities = json("capabilities").arr.toVector
    assertEquals(capabilities.map(_("name").str), Capability.all.map(_.name))
    assert(capabilities.forall(_("inMemory").bool), "the loaded workbook answers everything")
    assertEquals(
      capabilities.filter(_("streaming").bool).map(_("name").str).toSet,
      Capability.streaming.map(_.name)
    )
    assertEquals(json("version"), ujson.Str(version))
    assertEquals(
      json("exitCodes"),
      ujson.Arr.from(
        Schema.exitCodes.map(e =>
          ujson.Obj("code" -> ujson.Num(e.code), "meaning" -> ujson.Str(e.meaning))
        )
      )
    )
    assertEquals(
      json("errorCodes").arr.map(_("code").str).toVector,
      ErrorCode.all,
      "errorCodes are ErrorCode.all, in order"
    )
    json("errorCodes").arr.foreach { e =>
      assertEquals(e("exit"), ujson.Num(ExitCodes.forCode(e("code").str).code), e.toString)
    }
    assertEquals(json("warningCodes"), ujson.Arr.from(WarningCode.all.map(ujson.Str.apply)))
    assertEquals(json("globals").arr.size, Schema.globals.size)
    json("globals").arr.foreach { g =>
      assertEquals(g.obj.keys.toList, List("name", "short", "takesValue", "doc"))
      assertEquals(g("takesValue"), ujson.Bool(Argv.globals.getOrElse(g("name").str, false)))
    }
    assertEquals(json("verbs").arr.size, Schema.verbs.size)
    json("verbs").arr.foreach { v =>
      assertEquals(
        v.obj.keys.toList,
        List("path", "summary", "needs", "stream", "refusedWith", "exit", "batchTwin", "since")
      )
      assertEquals(v("needs").obj.keys.toList, List("file", "sheet", "output", "streaming"))
      assert(Set("o1", "backend", "refused").contains(v("stream").str), v("stream").str)
      assertEquals(v("stream").str == "o1", v("needs")("streaming").bool, v("path").toString)
      // only an O(1) verb has flags of its own to refuse under --stream
      if v("stream").str != "o1" then assertEquals(v("refusedWith").arr.size, 0, v("path").toString)
    }
    assertEquals(json("functions"), FunctionDoc.toJson(FunctionDoc.all))
  }

  test("batchOps is OpRegistry.jsonSchema with one oneOf entry per op (32)") {
    val json = Schema.json(version)
    assertEquals(json("batchOps"), OpRegistry.jsonSchema(version))
    val oneOf = json("batchOps")("items")("oneOf").arr
    assertEquals(oneOf.size, OpRegistry.all.size)
    assertEquals(oneOf.size, 32)
    assertEquals(oneOf.map(_("title").str).toVector, OpRegistry.all.map(_.name))
  }

  test("the envelope schema the binary serves is the one the contract suite checks against") {
    assertEquals(Schema.json(version)("envelope"), EnvelopeSchema.schema)
    assertEquals(Schema.envelope("title"), ujson.Str("xl envelope"))
    assertEquals(
      Schema.envelope("required"),
      ujson.Arr.from(
        List("ok", "exitCode", "verb", "version", "data", "warnings", "error").map(ujson.Str.apply)
      )
    )
  }

  // ---------------------------------------------------------------------------------------------
  // Through the harness
  // ---------------------------------------------------------------------------------------------

  test("schema --json: the envelope carries Schema.json(version) as data") {
    CliHarness.run("schema", "--json").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val e = envelope(run)
      assertEquals(e("ok"), ujson.True)
      assertEquals(e("verb"), ujson.Str("schema"))
      assertEquals(e("data"), Schema.json(version))
      assertEquals(e("data")("version"), e("version"))
      assertEquals(run.stderr, "")
    }
  }

  test("schema: the text table names every verb path and the exit codes") {
    CliHarness.run("schema").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      assertEquals(run.stdout, Schema.verbTable(version) + "\n")
      Schema.verbs.foreach { v =>
        assert(run.stdout.linesIterator.exists(_.startsWith(v.path.mkString(" "))), v.path.toString)
      }
      assert(run.stdout.contains(version), "the table names the version")
      assert(run.stdout.contains("xl schema --json"), "the table points at its JSON form")
      assertEquals(run.stderr, "")
    }
  }

  test("batch --schema prints the batch JSON Schema alone; --json wraps it as data") {
    for
      bare <- CliHarness.run("batch", "--schema")
      wrapped <- CliHarness.run("--json", "batch", "--schema")
    yield
      assertEquals(bare.exit, 0, bare.stderr)
      val schema = ujson.read(bare.stdout)
      assertEquals(schema, OpRegistry.jsonSchema(version))
      assertEquals(schema("$schema"), ujson.Str("https://json-schema.org/draft/2020-12/schema"))
      assertEquals(schema("type"), ujson.Str("array"))
      assertEquals(schema("items")("oneOf").arr.size, 32)
      assertEquals(
        bare.stdout,
        ujson.write(schema, indent = 2) + "\n",
        "the schema and nothing else"
      )
      assertEquals(bare.stderr, "")
      assertEquals(wrapped.exit, 0, wrapped.stderr)
      val e = envelope(wrapped)
      assertEquals(e("verb"), ujson.Str("batch"))
      assertEquals(e("data"), schema)
  }

  test("batch --help lists all 32 ops with their examples") {
    CliHarness.run("batch", "--help").map { run =>
      assertEquals(run.exit, 0)
      OpRegistry.all.foreach { spec =>
        assert(
          run.stdout.linesIterator.exists(_.trim.startsWith(spec.name + " ")),
          s"batch --help does not list ${spec.name}"
        )
        assert(run.stdout.contains(ujson.write(spec.example)), s"no example for ${spec.name}")
      }
      assert(run.stdout.contains("--schema"), "the help mentions batch --schema")
    }
  }

  test("functions --json: every FunctionDoc row, LET flagged as the special form") {
    CliHarness.run("--json", "functions").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      val e = envelope(run)
      assertEquals(e("verb"), ujson.Str("functions"))
      assertEquals(e("data"), FunctionDoc.toJson(FunctionDoc.all))
      val rows = e("data").arr
      assertEquals(rows.size, FunctionDoc.all.size)
      assertEquals(rows.count(_("specialForm") == ujson.True), 1)
      assertEquals(rows.lastOption.map(_("name")), Some(ujson.Str("LET")))
    }
  }

  test("functions (text) is unchanged by the typed JSON: the same list as before") {
    CliHarness.run("functions").map { run =>
      assertEquals(run.exit, 0, run.stderr)
      assert(run.stdout.startsWith("Supported Excel Functions ("), run.stdout)
      assert(!run.stdout.contains("LET"), "the text listing is the registry, as before")
    }
  }
