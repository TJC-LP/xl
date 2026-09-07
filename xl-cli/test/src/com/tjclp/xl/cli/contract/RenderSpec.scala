package com.tjclp.xl.cli.contract

import cats.effect.ExitCode
import munit.FunSuite

/**
 * The envelope and its text twin (ADR-017 §2.4), rendered from an [[Outcome]] without running a
 * command: seven keys, `ok` ⇔ `error == null`, `exitCode` from the code table, `version` present,
 * `data` from the payload, and every rendering valid against `envelope.schema.json`. `Render.text`
 * is today's bytes: the payload on stdout, the diagnostics block or warnings on stderr.
 */
class RenderSpec extends FunSuite:

  private val version = "1.2.3"
  private val sevenKeys = List("ok", "exitCode", "verb", "version", "data", "warnings", "error")

  /** Parse a JSON rendering, check it against the schema, and hand back the object. */
  private def envelope(rendered: Rendered): ujson.Obj =
    val parsed = ujson.read(rendered.stdout)
    EnvelopeSchema.assertValid(parsed)
    assertEquals(parsed.obj.keys.toList, sevenKeys, "exactly the seven keys, in order")
    parsed.obj

  private val notFound = CliError(
    "SHEET_NOT_FOUND",
    "Sheet not found: Nope. Available: Data, Summary",
    hint = Some("list sheets with `xl -f <file> sheets`"),
    candidates = Vector("Data"),
    location = Some(Location(Some("in.xlsx"), Some("Nope"), None, None))
  )

  // ---------------------------------------------------------------------------------------------
  // Render.json
  // ---------------------------------------------------------------------------------------------

  test("json: a Text payload success is the seven-key envelope with ok:true and error:null") {
    val outcome =
      Outcome.ok(
        "put",
        Payload.Text("Put: A5 = 42 (number)\nSaved: out.xlsx", Some("out.xlsx"), true)
      )
    val rendered = Render.json(outcome, version)
    val e = envelope(rendered)
    assertEquals(e("ok"), ujson.True)
    assertEquals(e("exitCode"), ujson.Num(0))
    assertEquals(e("verb"), ujson.Str("put"))
    assertEquals(e("version"), ujson.Str(version))
    assertEquals(
      e("data"),
      ujson.Obj(
        "text" -> ujson.Str("Put: A5 = 42 (number)\nSaved: out.xlsx"),
        "saved" -> ujson.Str("out.xlsx"),
        "written" -> ujson.True
      )
    )
    assertEquals(e("warnings"), ujson.Arr())
    assertEquals(e("error"), ujson.Null)
    assertEquals(rendered.stderr, "", "a success writes nothing to stderr")
  }

  test("json: a Text payload with nothing committed has saved:null and written:false") {
    val e = envelope(Render.json(Outcome.ok("view", Payload.text("| A |")), version))
    assertEquals(
      e("data"),
      ujson.Obj("text" -> ujson.Str("| A |"), "saved" -> ujson.Null, "written" -> ujson.False)
    )
  }

  test("json: a Json payload is data verbatim") {
    val value = ujson.Obj(
      "sheet" -> ujson.Str("Data"),
      "range" -> ujson.Str("A1:B1"),
      "rows" -> ujson.Arr(ujson.Obj("row" -> ujson.Num(1), "cells" -> ujson.Arr()))
    )
    val e = envelope(Render.json(Outcome.ok("view", Payload.Json(value)), version))
    assertEquals(e("data"), value)
  }

  test(
    "json: a failure is ok:false, data:null, exit per the code table, one Error: line on stderr"
  ) {
    val rendered = Render.json(Outcome.failed("put", notFound), version)
    val e = envelope(rendered)
    assertEquals(e("ok"), ujson.False)
    assertEquals(e("exitCode"), ujson.Num(3))
    assertEquals(e("verb"), ujson.Str("put"))
    assertEquals(e("data"), ujson.Null)
    assertEquals(
      e("error"),
      ujson.Obj(
        "code" -> ujson.Str("SHEET_NOT_FOUND"),
        "message" -> ujson.Str("Sheet not found: Nope. Available: Data, Summary"),
        "hint" -> ujson.Str("list sheets with `xl -f <file> sheets`"),
        "candidates" -> ujson.Arr(ujson.Str("Data")),
        "location" -> ujson.Obj(
          "file" -> ujson.Str("in.xlsx"),
          "sheet" -> ujson.Str("Nope"),
          "ref" -> ujson.Null,
          "opIndex" -> ujson.Null
        )
      )
    )
    assertEquals(rendered.stderr, "Error: Sheet not found: Nope. Available: Data, Summary")
  }

  test("json: an error without hint, candidates or location renders them as null/empty/null") {
    val e =
      envelope(Render.json(Outcome.failed("lint", CliError("IO_READ", "No such file")), version))
    assertEquals(
      e("error"),
      ujson.Obj(
        "code" -> ujson.Str("IO_READ"),
        "message" -> ujson.Str("No such file"),
        "hint" -> ujson.Null,
        "candidates" -> ujson.Arr(),
        "location" -> ujson.Null
      )
    )
  }

  test("json: a gate keeps its report as data, is ok:false and exits 1") {
    val summary = "Recalculated 1 formula\nNOT saved (--strict failure): in.xlsx left untouched"
    val gate =
      CliError(ErrorCode.RECALC_GATE, "STRICT FAILURE (--strict): 1 formula evaluation error(s)")
    val rendered =
      Render.json(Outcome.signal("recalc", Payload.Text(summary, None, false), gate), version)
    val e = envelope(rendered)
    assertEquals(e("ok"), ujson.False)
    assertEquals(e("exitCode"), ujson.Num(1))
    assertEquals(e("error")("code"), ujson.Str("RECALC_GATE"))
    assertEquals(
      e("data"),
      ujson.Obj("text" -> ujson.Str(summary), "saved" -> ujson.Null, "written" -> ujson.False)
    )
    assertEquals(rendered.stderr, s"Error: ${gate.message}")
  }

  test("json: ok ⇔ error == null, and exitCode follows the code table for every code") {
    ErrorCode.all.foreach { code =>
      val e = envelope(Render.json(Outcome.failed("v", CliError(code, "m")), version))
      assertEquals(e("ok"), ujson.False, code)
      assertEquals(e("exitCode"), ujson.Num(ExitCodes.forCode(code).code), code)
      assertEquals(e("error")("code"), ujson.Str(code))
    }
    val ok = envelope(Render.json(Outcome.ok("v", Payload.text("fine")), version))
    assertEquals(ok("ok"), ujson.True)
    assertEquals(ok("exitCode"), ujson.Num(0))
    assertEquals(ok("error"), ujson.Null)
  }

  test("json: warnings carry code and message; location only when known") {
    val warnings = Vector(
      Warning(WarningCode.TRUNCATED, "… showing 2 of 4 rows"),
      Warning(
        WarningCode.READER_WARNING,
        "MissingStylesXml",
        Some(Location(Some("in.xlsx"), None, None, None))
      )
    )
    val e = envelope(Render.json(Outcome.ok("view", Payload.text("| A |"), warnings), version))
    assertEquals(
      e("warnings"),
      ujson.Arr(
        ujson
          .Obj("code" -> ujson.Str("TRUNCATED"), "message" -> ujson.Str("… showing 2 of 4 rows")),
        ujson.Obj(
          "code" -> ujson.Str("READER_WARNING"),
          "message" -> ujson.Str("MissingStylesXml"),
          "location" -> ujson.Obj(
            "file" -> ujson.Str("in.xlsx"),
            "sheet" -> ujson.Null,
            "ref" -> ujson.Null,
            "opIndex" -> ujson.Null
          )
        )
      )
    )
  }

  test("json: the envelope is the whole of stdout — one indented document, no trailer") {
    val rendered = Render.json(Outcome.ok("sheets", Payload.Json(ujson.Arr())), version)
    assertEquals(rendered.stdout, ujson.write(ujson.read(rendered.stdout), indent = 2))
    assert(rendered.stdout.startsWith("{\n  \"ok\": true"), rendered.stdout)
  }

  // ---------------------------------------------------------------------------------------------
  // Render.text — today's bytes
  // ---------------------------------------------------------------------------------------------

  test("text: stdout is the payload text, stderr is empty") {
    assertEquals(
      Render.text(
        Outcome.ok("put", Payload.Text("Put: A5 = 42\nSaved: out.xlsx", Some("out.xlsx"), true))
      ),
      Rendered("Put: A5 = 42\nSaved: out.xlsx", "")
    )
  }

  test("text: warnings are Warning[CODE] lines on stderr, in order") {
    val warnings = Vector(
      Warning(WarningCode.READER_WARNING, "MissingStylesXml"),
      Warning(WarningCode.TRUNCATED, "… showing 2 of 4 rows")
    )
    assertEquals(
      Render.text(Outcome.ok("view", Payload.text("| A |"), warnings)),
      Rendered(
        "| A |",
        "Warning[READER_WARNING]: MissingStylesXml\nWarning[TRUNCATED]: … showing 2 of 4 rows"
      )
    )
  }

  test("text: a failure is the diagnostics block on stderr and nothing on stdout") {
    assertEquals(
      Render.text(Outcome.failed("put", notFound)),
      Rendered("", Diagnostics.render(notFound))
    )
  }

  test("text: a failure's warnings follow the diagnostics block") {
    val warning = Warning(WarningCode.READER_WARNING, "MissingStylesXml")
    assertEquals(
      Render.text(Outcome.failed("put", notFound, Vector(warning))),
      Rendered("", s"${Diagnostics.render(notFound)}\n${Diagnostics.renderWarning(warning)}")
    )
  }

  test("text: a gate's report stays on stdout and stderr carries no Error line") {
    val summary = "Recalculated 1 formula\nSTRICT FAILURE (--strict): 1 formula evaluation error(s)"
    val gate =
      CliError(ErrorCode.RECALC_GATE, "STRICT FAILURE (--strict): 1 formula evaluation error(s)")
    assertEquals(
      Render.text(Outcome.signal("recalc", Payload.Text(summary, Some("out.xlsx"), true), gate)),
      Rendered(summary, "")
    )
  }

  test("text: a Json payload prints as indented JSON") {
    val value = ujson.Obj("identical" -> ujson.Bool(false))
    assertEquals(
      Render.text(Outcome.ok("diff", Payload.Json(value))),
      Rendered(ujson.write(value, indent = 2), "")
    )
  }

  test("Render(mode) selects the renderer") {
    val outcome = Outcome.ok("eval", Payload.text("Formula: =1+1\nResult: 2 (number)"))
    assertEquals(Render(OutputMode.Text)(outcome, version), Render.text(outcome))
    assertEquals(Render(OutputMode.Json)(outcome, version), Render.json(outcome, version))
  }

  // ---------------------------------------------------------------------------------------------
  // Outcome constructors
  // ---------------------------------------------------------------------------------------------

  test("Outcome.ok exits 0 and is complete; failed exits per its code and is incomplete") {
    val ok = Outcome.ok("view", Payload.text("x"))
    assertEquals(ok.exitCode, ExitCode.Success)
    assertEquals(ok.error, None)
    assert(ok.outputComplete)

    val failed = Outcome.failed("put", notFound)
    assertEquals(failed.exitCode, ExitCodes.failed)
    assertEquals(failed.payload, None)
    assert(!failed.outputComplete)

    val usage = Outcome.failed("", CliError.usage("Unexpected argument: frob", None))
    assertEquals(usage.exitCode, ExitCodes.usage)
  }

  test("Outcome.signal exits per its code (1 for a gate or findings) and keeps the payload") {
    val gate = Outcome.signal(
      "recalc",
      Payload.text("summary"),
      CliError(ErrorCode.RECALC_GATE, "strict")
    )
    assertEquals(gate.exitCode, ExitCodes.signal)
    assertEquals(gate.payload, Some(Payload.text("summary")))
    assert(gate.outputComplete)
    val differs = Outcome.signal(
      "diff",
      Payload.Json(ujson.Obj("identical" -> ujson.False)),
      CliError(ErrorCode.DIFFERENCES_FOUND, "Differences found")
    )
    assertEquals(differs.exitCode, ExitCode(1))
  }
