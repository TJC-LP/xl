package com.tjclp.xl.cli.contract

import cats.effect.IO
import fs2.Stream

/**
 * What a verb produced (ADR-017 §2.4): the `data` of the `--json` envelope, and the bytes text mode
 * prints on stdout.
 */
enum Payload derives CanEqual:
  /**
   * The legacy bridge for prose verbs: the text exactly as text mode prints it, plus the two facts
   * every write knows — the path the run committed to (the user-visible one, never the staging
   * temp) and whether it committed at all. `saved = None`, `written = false` when nothing was
   * committed: a read verb, or an `-i` run whose strict gate discarded the staged output.
   */
  case Text(text: String, saved: Option[String], written: Boolean)

  /**
   * Typed verbs whose data holds only strings, booleans and small integers (`sheets`, `names`,
   * `bounds`, `functions`, `rasterizers`, `batch --dry-run`, `describe`, `audit`, `deps`) build
   * ujson directly.
   */
  case Json(value: ujson.Value)

  /**
   * Valid JSON text produced by a renderer of ours ([[com.tjclp.xl.cli.output.JsonRenderer]],
   * `DiffCommands.renderJson`, …), spliced into the envelope AS TEXT — never through a ujson tree,
   * whose numbers are `Double`-backed: `12345678901234567` would come out as `12345678901234568`.
   * The pass-through verbs (`view`/`filter`/`diff`/`lint` with a JSON payload format) and the
   * number-carrying typed verbs (`eval`, `evala`) use this so every number lexeme in `data` is
   * exactly what the bare `--format json` output prints.
   */
  case Raw(json: String)

  /**
   * A table written row by row (GH-635): what `view` yields for csv, json and markdown from either
   * source, so a read never holds its window — the loaded workbook's rows stream from memory, the
   * streaming reader's from the file, one row at a time. Both modes write the rows as they arrive
   * ([[Render.stream]]: the table itself, or the envelope with the document spliced element by
   * element and a text table's `data.text` escaped line by line); [[Payload.materialise]] is the
   * same bytes as one [[Text]] or [[Raw]], which is how the parity law sees it. Every warning of a
   * streamed read is raised before its first row, so the outcome that carries the payload is
   * complete when it is built; nothing is written before the first row has been read, so a source
   * that fails before it has one fails as any other run does, and a failure after that stops the
   * output where it is and is reported on stderr with its exit code.
   */
  case Streamed(body: StreamedBody)

/** How a [[Payload.Streamed]] composes: what is written before the rows, per row, and after. */
enum StreamedBody:
  /**
   * Lines before the rows, one line per row, lines after: the payload is their `\n`-join (csv,
   * markdown). A row may carry its own line breaks (a quoted CSV field); it is still one line here.
   */
  case Lines(before: Vector[String], rows: Stream[IO, String], after: Vector[String])

  /**
   * A JSON document whose last member is an array of the rows (`view --format json`): `head` is the
   * text up to and including the array's `[`, `tail` the text from its `]` on, and every row one
   * element as JSON text on its own line, indented as the document lays it out.
   */
  case JsonArray(head: String, rows: Stream[IO, String], tail: String)

object StreamedBody:

  /** The text of a [[StreamedBody.Lines]] once its rows are all in. */
  def lines(before: Vector[String], rows: Vector[String], after: Vector[String]): String =
    (before ++ rows ++ after).mkString("\n")

  /**
   * The text of a [[StreamedBody.JsonArray]] once its elements are all in — the layout `view
   * --format json` has always printed: the elements on their own lines, comma-separated, the
   * closing bracket indented under its key; an empty array closes on the same line (`[]`).
   */
  def jsonArray(head: String, elements: Vector[String], tail: String): String =
    val block = if elements.isEmpty then "" else "\n" + elements.mkString(",\n") + "\n  "
    head + block + tail

object Payload:

  /** Prose with nothing committed (yet). */
  def text(text: String): Payload = Payload.Text(text, None, false)

  /**
   * A [[Payload.Streamed]] as the [[Text]] or [[Raw]] payload its rows compose ([[StreamedBody]]);
   * every other payload unchanged. For the parity law and the tests: the runner never gathers a
   * streamed table.
   */
  def materialise(payload: Payload): IO[Payload] = payload match
    case Streamed(StreamedBody.Lines(before, rows, after)) =>
      rows.compile.toVector.map(rs => Payload.text(StreamedBody.lines(before, rs, after)))
    case Streamed(StreamedBody.JsonArray(head, rows, tail)) =>
      rows.compile.toVector.map(rs => Payload.Raw(StreamedBody.jsonArray(head, rs, tail)))
    case other => IO.pure(other)

  /**
   * The envelope's `data` as a ujson tree — for callers that inspect it, not for rendering: a
   * [[Payload.Raw]] read this way loses number precision, which is why [[Render.json]] splices its
   * text instead. A [[Payload.Streamed]] has no tree until it is materialised ([[materialise]]) and
   * [[Render]] never asks for one: reaching this with one is a defect, reported as such.
   */
  def toJson(payload: Payload): ujson.Value = payload match
    case Text(text, saved, written) =>
      ujson.Obj(
        "text" -> ujson.Str(text),
        "saved" -> saved.fold[ujson.Value](ujson.Null)(ujson.Str.apply),
        "written" -> ujson.Bool(written)
      )
    case Json(value) => value
    case Raw(json) => ujson.read(json)
    case Streamed(_) =>
      throw new IllegalStateException(
        "a streamed payload has no JSON tree until it is materialised (Payload.materialise)"
      )

  /**
   * Record what the staging step actually did: a committed run names its target; a run that
   * committed nothing has `saved = None` and `written = false`. JSON payloads carry no such facts,
   * and a streamed table is a read's.
   */
  def committed(payload: Payload, target: Option[String]): Payload = payload match
    case Text(text, _, _) => Text(text, target, target.isDefined)
    case json: Json => json
    case raw: Raw => raw
    case streamed: Streamed => streamed
