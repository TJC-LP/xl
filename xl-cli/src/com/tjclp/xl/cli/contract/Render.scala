package com.tjclp.xl.cli.contract

import scala.util.Try

import cats.effect.IO
import fs2.{Chunk, Pull, Stream}

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
 * { "ok": true, "exitCode": 0, "verb": "view", "version": "0.22.0",
 *   "data": …, "warnings": [ { "code": "TRUNCATED", "message": "…" } ], "error": null }
 * }}}
 * `ok ⇔ error == null`; `data` is `null` on a failure, `{text, saved, written}` for prose verbs and
 * the verb's own JSON otherwise — a [[Payload.Raw]] is spliced in as the text its renderer
 * produced, re-indented, so no number lexeme is ever rounded through a `Double`. stderr follows the
 * text rule: the one-line `Error: <message>` only when a failure has no payload, so a human tailing
 * a log still sees it; a signal (findings, a failed gate) keeps its report as `data` and prints
 * nothing on stderr, exactly as text mode does for the same run.
 *
 * [[stream]] is the stdout of either mode written as it is produced (GH-635): the bytes `emit`
 * prints, a fragment at a time for a [[Payload.Streamed]] — the same bytes as [[apply]] would print
 * once the rows were all in.
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
      // The rows are not in hand: [[stream]] writes them; nothing else may ask for this text
      case Payload.Streamed(_) => ""
    }
    Rendered(stdout, stderr(OutputMode.Text)(outcome))

  def json(outcome: Outcome, version: String): Rendered =
    val stdout = outcome.payload match
      case Some(Payload.Raw(raw)) => spliced(outcome, version, raw)
      case Some(Payload.Streamed(_)) => ""
      case other =>
        ujson.write(
          envelope(outcome, version, other.fold[ujson.Value](ujson.Null)(Payload.toJson)),
          indent = 2
        )
    Rendered(stdout, stderr(OutputMode.Json)(outcome))

  /**
   * The stderr of a run in `mode`, which never depends on the payload's text — only on whether
   * there is one: text mode prints the diagnostics block of a failure without a payload, then the
   * warnings; `--json` the one-line `Error:` of such a failure and nothing else.
   */
  def stderr(mode: OutputMode)(outcome: Outcome): String = mode match
    case OutputMode.Text =>
      val diagnostics = errorWithoutPayload(outcome).map(Diagnostics.render).toList
      val warnings = outcome.warnings.map(Diagnostics.renderWarning).toList
      (diagnostics ++ warnings).mkString("\n")
    case OutputMode.Json =>
      errorWithoutPayload(outcome).fold("")(e => s"Error: ${e.message}")

  /**
   * The bytes this run writes on stdout in `mode`, as text fragments in order (GH-635): exactly
   * what `emit` prints for the outcome — [[apply]]'s stdout followed by the newline `println` adds,
   * or nothing at all when that stdout is empty — produced as the rows arrive for a
   * [[Payload.Streamed]]. In text mode that is the table itself (csv and markdown lines, the JSON
   * document in the layout `view --format json` prints); under `--json` the envelope, with a JSON
   * document spliced element by element in the layout `ujson` gives a [[Payload.Raw]] ([[spliced]])
   * and a text table's `data.text` escaped line by line as `ujson` escapes one string — so a
   * consumer reads the same bytes whether the rows were streamed or gathered, and nothing is ever
   * gathered. Nothing is produced before the first row has been read: a source that fails before it
   * has one (the table's ZIP-bomb limit, a worksheet the parser rejects at once) fails before the
   * first byte, and the runner reports it as any other failure.
   */
  def stream(mode: OutputMode)(outcome: Outcome, version: String): Stream[IO, String] =
    outcome.payload match
      case Some(Payload.Streamed(body)) =>
        mode match
          case OutputMode.Text => textFragments(body)
          case OutputMode.Json => envelopeFragments(outcome, version, body)
      case _ =>
        val stdout = apply(mode)(outcome, version).stdout
        if stdout.isEmpty then Stream.empty else Stream.emit(stdout + "\n")

  /** The escaped newline: what `ujson` writes for a line break inside a string. */
  private val escapedNewline: String = "\\n"

  /** `text` as `ujson` writes it inside a JSON string: the escaped characters, quotes excluded. */
  private def escaped(text: String): String =
    val quoted = ujson.write(ujson.Str(text))
    quoted.substring(1, quoted.length - 1)

  /**
   * The text form of a streamed body: the table, then the newline `println` adds — nothing at all
   * for a table whose text is empty (as printing the empty text never printed).
   */
  private def textFragments(body: StreamedBody): Stream[IO, String] = body match
    case StreamedBody.Lines(before, rows, after) =>
      // StreamedBody.lines, a row at a time; the first fragment waits for the first row, and a text
      // that turns out empty (no lines, or one empty line) is no bytes
      def whole(lines: Vector[String]): Pull[IO, String, Unit] =
        val text = lines.mkString("\n")
        if text.isEmpty then Pull.done else Pull.output1(text + "\n")
      rows.pull.uncons1.flatMap {
        case None => whole(before ++ after)
        case Some((first, rest)) =>
          rest.pull.uncons1.flatMap {
            case None => whole((before :+ first) ++ after)
            case Some((second, more)) =>
              Pull.output1((before :+ first).mkString("\n")) >>
                Pull.output1("\n" + second) >>
                more.map("\n" + _).pull.echo >>
                Pull.output1(after.map("\n" + _).mkString + "\n")
          }
      }.stream
    case StreamedBody.JsonArray(head, rows, tail) =>
      // StreamedBody.jsonArray, an element at a time
      rows.pull.uncons1.flatMap {
        case None => Pull.output1(head + tail + "\n")
        case Some((first, rest)) =>
          Pull.output1(head + "\n" + first) >>
            rest.map(",\n" + _).pull.echo >>
            Pull.output1("\n  " + tail + "\n")
      }.stream

  /**
   * The envelope with a streamed body as `data`, a fragment at a time, the newline `println` adds
   * last. A [[StreamedBody.Lines]] is `data.text`: the envelope's shell is rendered around a marker
   * where the text goes and split there, and the lines are written escaped as `ujson` escapes a
   * string — character by character, so the pieces spell what one write of the whole would — with
   * the escaped newline between them. A [[StreamedBody.JsonArray]] is the spliced document
   * ([[spliced]]) piecewise: the document with no elements is re-laid-out by `ujson` with a marker
   * in the array's slot, each element is re-laid-out on its own and indented to the array's depth,
   * and every piece is shifted by the slot's depth exactly as [[spliced]] shifts the whole block.
   * Either way the first fragment waits for the first row.
   */
  private def envelopeFragments(
    outcome: Outcome,
    version: String,
    body: StreamedBody
  ): Stream[IO, String] =
    body match
      case StreamedBody.Lines(before, rows, after) =>
        val (shellHead, shellTail) = textShell(outcome, version)
        rows.pull.uncons1.flatMap {
          case None =>
            val text = (before ++ after).map(escaped).mkString(escapedNewline)
            Pull.output1(shellHead + "\"" + text + "\"" + shellTail + "\n")
          case Some((first, rest)) =>
            Pull.output1(
              shellHead + "\"" + (before :+ first).map(escaped).mkString(escapedNewline)
            ) >>
              rest.map(row => escapedNewline + escaped(row)).pull.echo >>
              Pull.output1(
                after.map(line => escapedNewline + escaped(line)).mkString + "\"" + shellTail +
                  "\n"
              )
        }.stream
      case StreamedBody.JsonArray(head, rows, tail) =>
        val (shellHead, shellTail) = shell(outcome, version)
        val (dataHead, dataTail) = emptyDocument(head + tail)
        def shift(s: String): String = s.replace("\n", "\n  ")
        rows.pull.uncons1.flatMap {
          case None =>
            Pull.output1(shellHead + shift(dataHead + "[]" + dataTail) + shellTail + "\n")
          case Some((first, rest)) =>
            Pull.output1(shellHead + shift(dataHead) + "[") >>
              (Stream.emit(first) ++ rest).zipWithNext
                .map { (element, next) =>
                  val laidOut = ujson.reformat(element, indent = 2).replace("\n", "\n    ")
                  shift("\n    " + laidOut) + (if next.isDefined then "," else "")
                }
                .pull
                .echo >>
              Pull.output1(shift("\n  ]" + dataTail) + shellTail + "\n")
        }.stream

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

  /** The envelope's text around its `data` slot: what precedes the slot and what follows it. */
  private def shell(outcome: Outcome, version: String): (String, String) =
    splitAtMarker(ujson.write(envelope(outcome, version, ujson.Str(marker)), indent = 2))

  /**
   * The envelope's text around the `data.text` of a read's prose payload (`{text, saved: null,
   * written: false}`): what precedes the string's opening quote and what follows its closing one.
   */
  private def textShell(outcome: Outcome, version: String): (String, String) =
    val data = Payload.toJson(Payload.Text(marker, None, false))
    splitAtMarker(ujson.write(envelope(outcome, version, data), indent = 2))

  /**
   * A JSON object's text, laid out by `ujson`, around the value of its LAST member: the document
   * `view --format json` prints with its array empty, re-laid-out with the marker in the array's
   * place. What precedes the marker ends with `"rows": ` (or `"records": `); what follows is `\n}`.
   */
  private def emptyDocument(document: String): (String, String) =
    // `.obj` is the members' LinkedHashMap itself (a `.value` on it would be a converted copy)
    val members = ujson.read(document).obj
    members.keys.lastOption.foreach(key => members.update(key, ujson.Str(marker)))
    splitAtMarker(ujson.write(ujson.Obj(members), indent = 2))

  private def splitAtMarker(text: String): (String, String) =
    val slot = ujson.write(ujson.Str(marker))
    val at = text.indexOf(slot)
    if at < 0 then (text, "") else (text.substring(0, at), text.substring(at + slot.length))

  /**
   * The envelope with a [[Payload.Raw]] as `data`: render the shell with a marker in the data slot,
   * reformat the raw text with the same indentation (`ujson.reformat` re-emits every lexeme as read
   * — numbers included), indent its continuation lines to the slot's depth, and splice. Total: text
   * that is not JSON (a defect in the renderer that produced it) rides as a string so the envelope
   * stays well-formed.
   */
  private def spliced(outcome: Outcome, version: String, raw: String): String =
    val (before, after) = shell(outcome, version)
    val block = Try(ujson.reformat(raw, indent = 2)).getOrElse(ujson.write(ujson.Str(raw)))
    before + block.replace("\n", "\n  ") + after

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
