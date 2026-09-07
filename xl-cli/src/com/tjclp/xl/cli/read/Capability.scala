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

  def doc: String = this match
    case Values => "cell values, kinds and display text (view, search, stats, filter)"
    case Styles => "number formats and cell styles"
    case Formulas => "formula expressions with their cached values"
    case Comments => "cell comments"
    case Hyperlinks => "cell hyperlinks"
    case Hidden =>
      "hidden rows and columns: --skip-hidden and the hidden-line markers (view)"
    case Merges => "merged ranges (cell --json mergedInto)"
    case Graph => "exact precedents and dependents across sheets (cell)"
    case Eval => "formula evaluation (view --eval)"
    case Render => "styled renderings: html, svg, png, jpeg, webp, pdf (view --format)"

object Capability:

  /** Every capability, in table order. */
  val all: Vector[Capability] = Capability.values.toVector

  /** What a loaded workbook answers: everything. */
  val inMemory: Set[Capability] = all.toSet

  /**
   * What the streaming reader answers: values, styles (from `styles.xml`), cached formulas and
   * comments. It never parses row/column properties, merges, hyperlinks or the other sheets'
   * formulas, and it cannot evaluate or render.
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
