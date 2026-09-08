package com.tjclp.xl.cli.read

/**
 * What a [[SheetSource]] can answer (W2.4). A read verb declares what it needs
 * ([[ReadQuery.needs]]); a source declares what it has; the difference is refused in band as
 * `UNSUPPORTED_IN_STREAM` before anything is read. `schema --json` publishes the table.
 */
enum Capability derives CanEqual:
  /** Cell values, kinds and their display text. */
  case Values

  /** Number formats and the rest of a cell's style. */
  case Styles

  /** Formula expressions with their cached values. */
  case Formulas

  /** Cell comments. */
  case Comments

  /** Cell hyperlinks. */
  case Hyperlinks

  /** Hidden rows and columns, so `--skip-hidden` and the hidden-line markers can be honoured. */
  case Hidden

  /** Merged ranges. */
  case Merges

  /** The cross-sheet dependency graph: exact precedents and dependents. */
  case Graph

  /** Formula evaluation (`view --eval`). */
  case Eval

  /** The styled renderings: html, svg and the raster formats. */
  case Render

  /** The name `schema --json` publishes. */
  def name: String = this match
    case Values => "values"
    case Styles => "styles"
    case Formulas => "formulas"
    case Comments => "comments"
    case Hyperlinks => "hyperlinks"
    case Hidden => "hidden"
    case Merges => "merges"
    case Graph => "graph"
    case Eval => "eval"
    case Render => "render"

  /**
   * What the capability answers, and how its fields read from a source without it: `null`, or
   * `(not available in streaming mode)` in text — never a guess.
   */
  def doc: String = this match
    case Values => "cell values, kinds and display text (view, search, stats, filter)"
    case Styles => "number formats and cell styles"
    case Formulas => "formula expressions with their cached values"
    case Comments => "cell comments (cell)"
    case Hyperlinks => "cell hyperlinks (cell: hyperlink; null without it)"
    case Hidden =>
      "hidden rows and columns: --skip-hidden, the hidden-line markers (view) and the record's " +
        "hidden flag (null without it)"
    case Merges => "merged ranges (mergedInto; null without it)"
    case Graph =>
      "exact precedents and dependents across sheets (cell: dependencies and dependents; " +
        "null / '(not available in streaming mode)' without it)"
    case Eval => "formula evaluation (view --eval)"
    case Render => "styled renderings: html, svg, png, jpeg, webp, pdf (view --format)"

object Capability:

  /** Every capability, in table order. */
  val all: Vector[Capability] = Capability.values.toVector

  /** What a loaded workbook answers: everything. */
  val inMemory: Set[Capability] = all.toSet

  /**
   * What the streaming reader answers: values, styles (from `styles.xml`), formulas with their
   * cached values and comments. It never parses row/column properties, merges, hyperlinks or the
   * other sheets' formulas — the fields of those capabilities read `null` — and it cannot evaluate
   * or render, so `--eval` and the styled formats are refused.
   */
  val streaming: Set[Capability] = Set(Values, Styles, Formulas, Comments)

  /** The table `schema --json` publishes: `[{name, doc, inMemory, streaming}]`. */
  def json: ujson.Value =
    ujson.Arr.from(all.map { c =>
      ujson.Obj(
        "name" -> ujson.Str(c.name),
        "doc" -> ujson.Str(c.doc),
        "inMemory" -> ujson.Bool(inMemory.contains(c)),
        "streaming" -> ujson.Bool(streaming.contains(c))
      )
    })
