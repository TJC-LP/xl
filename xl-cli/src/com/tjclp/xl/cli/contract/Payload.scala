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
   * Verbs that already emit JSON (`view`/`filter`/`diff`/`lint` under their own `--format json`)
   * pass their payload through unchanged; typed verbs (`sheets`, `names`, `bounds`, `eval`,
   * `evala`, `functions`, `rasterizers`, `batch --dry-run`) build ujson directly.
   */
  case Json(value: ujson.Value)

object Payload:

  /** Prose with nothing committed (yet). */
  def text(text: String): Payload = Payload.Text(text, None, false)

  /** The envelope's `data`. */
  def toJson(payload: Payload): ujson.Value = payload match
    case Text(text, saved, written) =>
      ujson.Obj(
        "text" -> ujson.Str(text),
        "saved" -> saved.fold[ujson.Value](ujson.Null)(ujson.Str.apply),
        "written" -> ujson.Bool(written)
      )
    case Json(value) => value

  /**
   * Record what the staging step actually did: a committed run names its target; a run that
   * committed nothing has `saved = None` and `written = false`. JSON payloads carry no such facts.
   */
  def committed(payload: Payload, target: Option[String]): Payload = payload match
    case Text(text, _, _) => Text(text, target, target.isDefined)
    case json: Json => json
