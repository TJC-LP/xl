package com.tjclp.xl.cli.contract

/** How a run's [[Outcome]] reaches the streams: today's text, or the `--json` envelope. */
enum OutputMode derives CanEqual:
  case Text, Json

/** The bytes of one run, per channel. Empty means nothing is printed on that channel. */
final case class Rendered(stdout: String, stderr: String) derives CanEqual

/**
 * The two renderings of an [[Outcome]] (ADR-017 §2.4).
 *
 * [[text]] is today's contract byte for byte: the payload on stdout; on stderr the diagnostics
 * block ([[Diagnostics.render]]) when a failure has no payload, then one `Warning[CODE]: …` line
 * per warning. An error that arrives WITH a payload is a signal — findings or a failed gate — whose
 * report is the payload itself, so the text form prints that report and no `Error:` line, exactly
 * as `diff`, `lint` and a `--strict` write always have.
 *
 * [[json]] is the envelope, identical for every verb, success or failure, and the only thing on
 * stdout:
 * {{{
 * { "ok": true, "exitCode": 0, "verb": "view", "version": "0.20.0",
 *   "data": …, "warnings": [ { "code": "TRUNCATED", "message": "…" } ], "error": null }
 * }}}
 * `ok ⇔ error == null`; `data` is `null` on a failure, `{text, saved, written}` for prose verbs and
 * the verb's own JSON otherwise. Whenever `error` is present, stderr carries the one-line
 * `Error: <message>` so a human tailing a log still sees it.
 */
object Render:

  def apply(mode: OutputMode)(outcome: Outcome, version: String): Rendered = mode match
    case OutputMode.Text => text(outcome)
    case OutputMode.Json => json(outcome, version)

  def text(outcome: Outcome): Rendered =
    val stdout = outcome.payload.fold("") {
      case Payload.Text(text, _, _) => text
      case Payload.Json(value) => ujson.write(value, indent = 2)
    }
    val diagnostics =
      outcome.error.filter(_ => outcome.payload.isEmpty).map(Diagnostics.render).toList
    val warnings = outcome.warnings.map(Diagnostics.renderWarning).toList
    Rendered(stdout, (diagnostics ++ warnings).mkString("\n"))

  def json(outcome: Outcome, version: String): Rendered =
    val envelope = ujson.Obj(
      "ok" -> ujson.Bool(outcome.ok),
      "exitCode" -> ujson.Num(outcome.exitCode.code),
      "verb" -> ujson.Str(outcome.verb),
      "version" -> ujson.Str(version),
      "data" -> outcome.payload.fold[ujson.Value](ujson.Null)(Payload.toJson),
      "warnings" -> ujson.Arr.from(outcome.warnings.map(warningJson)),
      "error" -> outcome.error.fold[ujson.Value](ujson.Null)(errorJson)
    )
    Rendered(
      ujson.write(envelope, indent = 2),
      outcome.error.fold("")(e => s"Error: ${e.message}")
    )

  /** `{code, message, hint, candidates, location}` — every key present, absent ones `null`. */
  def errorJson(error: CliError): ujson.Obj =
    ujson.Obj(
      "code" -> ujson.Str(error.code),
      "message" -> ujson.Str(error.message),
      "hint" -> error.hint.fold[ujson.Value](ujson.Null)(ujson.Str.apply),
      "candidates" -> ujson.Arr.from(error.candidates.map(ujson.Str.apply)),
      "location" -> error.location.fold[ujson.Value](ujson.Null)(locationJson)
    )

  /** `{code, message}`, plus `location` only when the raising site knew one. */
  def warningJson(warning: Warning): ujson.Obj =
    val base = ujson.Obj(
      "code" -> ujson.Str(warning.code),
      "message" -> ujson.Str(warning.message)
    )
    warning.location.foreach(l => base("location") = locationJson(l))
    base

  /** `{file, sheet, ref, opIndex}` — every key present, unknown ones `null`. */
  def locationJson(location: Location): ujson.Obj =
    ujson.Obj(
      "file" -> location.file.fold[ujson.Value](ujson.Null)(ujson.Str.apply),
      "sheet" -> location.sheet.fold[ujson.Value](ujson.Null)(ujson.Str.apply),
      "ref" -> location.ref.fold[ujson.Value](ujson.Null)(ujson.Str.apply),
      "opIndex" -> location.opIndex.fold[ujson.Value](ujson.Null)(i => ujson.Num(i))
    )
