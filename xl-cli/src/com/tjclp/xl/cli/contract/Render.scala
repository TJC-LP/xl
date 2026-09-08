package com.tjclp.xl.cli.contract

import scala.util.Try

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
 * { "ok": true, "exitCode": 0, "verb": "view", "version": "0.21.0",
 *   "data": …, "warnings": [ { "code": "TRUNCATED", "message": "…" } ], "error": null }
 * }}}
 * `ok ⇔ error == null`; `data` is `null` on a failure, `{text, saved, written}` for prose verbs and
 * the verb's own JSON otherwise — a [[Payload.Raw]] is spliced in as the text its renderer
 * produced, re-indented, so no number lexeme is ever rounded through a `Double`. stderr follows the
 * text rule: the one-line `Error: <message>` only when a failure has no payload, so a human tailing
 * a log still sees it; a signal (findings, a failed gate) keeps its report as `data` and prints
 * nothing on stderr, exactly as text mode does for the same run.
 */
object Render:

  def apply(mode: OutputMode)(outcome: Outcome, version: String): Rendered = mode match
    case OutputMode.Text => text(outcome)
    case OutputMode.Json => json(outcome, version)

  def text(outcome: Outcome): Rendered =
    val stdout = outcome.payload.fold("") {
      case Payload.Text(text, _, _) => text
      case Payload.Json(value) => ujson.write(value, indent = 2)
      case Payload.Raw(json) => json
    }
    val diagnostics = errorWithoutPayload(outcome).map(Diagnostics.render).toList
    val warnings = outcome.warnings.map(Diagnostics.renderWarning).toList
    Rendered(stdout, (diagnostics ++ warnings).mkString("\n"))

  def json(outcome: Outcome, version: String): Rendered =
    val stdout = outcome.payload match
      case Some(Payload.Raw(raw)) => spliced(outcome, version, raw)
      case other =>
        ujson.write(
          envelope(outcome, version, other.fold[ujson.Value](ujson.Null)(Payload.toJson)),
          indent = 2
        )
    Rendered(stdout, errorWithoutPayload(outcome).fold("")(e => s"Error: ${e.message}"))

  /** The error a channel reports as such: a failure's; a signal's report is its payload. */
  private def errorWithoutPayload(outcome: Outcome): Option[CliError] =
    outcome.error.filter(_ => outcome.payload.isEmpty)

  /** The seven keys, in order, around a given `data`. */
  private def envelope(outcome: Outcome, version: String, data: ujson.Value): ujson.Obj =
    ujson.Obj(
      "ok" -> ujson.Bool(outcome.ok),
      "exitCode" -> ujson.Num(outcome.exitCode.code),
      "verb" -> ujson.Str(outcome.verb),
      "version" -> ujson.Str(version),
      "data" -> data,
      "warnings" -> ujson.Arr.from(outcome.warnings.map(warningJson)),
      "error" -> outcome.error.fold[ujson.Value](ujson.Null)(errorJson)
    )

  /**
   * A NUL-delimited marker: JSON never carries an unescaped control character, and ujson writes it
   * as `\u0000`, which no other string in the envelope can spell verbatim (a backslash in a message
   * is itself escaped). So the marker's rendering occurs exactly once in the shell.
   */
  private val marker: String = "\u0000xl-raw-payload\u0000"

  /**
   * The envelope with a [[Payload.Raw]] as `data`: render the shell with a marker in the data slot,
   * reformat the raw text with the same indentation (`ujson.reformat` re-emits every lexeme as read
   * — numbers included), indent its continuation lines to the slot's depth, and splice. Total: text
   * that is not JSON (a defect in the renderer that produced it) rides as a string so the envelope
   * stays well-formed.
   */
  private def spliced(outcome: Outcome, version: String, raw: String): String =
    val shell = ujson.write(envelope(outcome, version, ujson.Str(marker)), indent = 2)
    val slot = ujson.write(ujson.Str(marker))
    val at = shell.indexOf(slot)
    val block = Try(ujson.reformat(raw, indent = 2)).getOrElse(ujson.write(ujson.Str(raw)))
    if at < 0 then shell
    else shell.substring(0, at) + block.replace("\n", "\n  ") + shell.substring(at + slot.length)

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
