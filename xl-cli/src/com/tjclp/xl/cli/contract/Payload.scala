package com.tjclp.xl.cli.contract

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

object Payload:

  /** Prose with nothing committed (yet). */
  def text(text: String): Payload = Payload.Text(text, None, false)

  /**
   * The envelope's `data` as a ujson tree — for callers that inspect it, not for rendering: a
   * [[Payload.Raw]] read this way loses number precision, which is why [[Render.json]] splices its
   * text instead.
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

  /**
   * Record what the staging step actually did: a committed run names its target; a run that
   * committed nothing has `saved = None` and `written = false`. JSON payloads carry no such facts.
   */
  def committed(payload: Payload, target: Option[String]): Payload = payload match
    case Text(text, _, _) => Text(text, target, target.isDefined)
    case json: Json => json
    case raw: Raw => raw
