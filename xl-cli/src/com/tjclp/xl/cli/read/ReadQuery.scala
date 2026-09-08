package com.tjclp.xl.cli.read

import java.nio.file.Path

import com.tjclp.xl.cli.{CliCommand, FilterFormat, ViewFormat}
import com.tjclp.xl.cli.contract.OutputMode

/**
 * A read verb with its options resolved (W2.4): the pass-through formats already chosen for the
 * run's mode ([[CliCommand.viewFormat]]), so [[Reads.run]] is a function of the query and a
 * [[SheetSource]] alone. Sheet selection (`-s`) travels beside the query, not inside it: it is the
 * run's, not the verb's.
 */
enum ReadQuery derives CanEqual:
  /**
   * `view [range]`: the range, or the sheet's used range when none is given; `offset` skips rows
   * from the top, `limit` caps the rows shown (0 = no limit), `maxCols` caps the columns (0 = no
   * limit).
   */
  case View(
    range: Option[String],
    showFormulas: Boolean,
    evalFormulas: Boolean,
    strict: Boolean,
    limit: Int,
    offset: Int,
    maxCols: Int,
    format: ViewFormat,
    printScale: Boolean,
    showGridlines: Boolean,
    showLabels: Boolean,
    dpi: Int,
    quality: Int,
    rasterOutput: Option[Path],
    skipEmpty: Boolean,
    headerRow: Option[Int],
    rasterizer: Option[String],
    skipHidden: Boolean
  )
  case Cell(ref: String, noStyle: Boolean)

  /**
   * `search <pattern>`: `limit` caps the matches listed (0 = no limit); `exactTotal` (`--total`)
   * scans every cell for the exact total, else the scan stops one match past the limit (GH-637).
   */
  case Search(pattern: String, limit: Int, sheetsFilter: Option[String], exactTotal: Boolean)
  case Stats(ref: String)
  case Filter(
    where: String,
    columns: Option[String],
    limit: Int,
    format: FilterFormat,
    header: Boolean
  )

  /** The envelope's `verb`. */
  def verb: String = this match
    case _: View => "view"
    case _: Cell => "cell"
    case _: Search => "search"
    case _: Stats => "stats"
    case _: Filter => "filter"

  /**
   * What the query needs from its source beyond values: `eval` for `--eval`, `render` for the
   * styled formats. A source lacking one refuses the query in band ([[Reads.run]]). `--skip-hidden`
   * is not a need: without `hidden` it is reported as ignored, as it always was under `--stream`.
   */
  def needs: Set[Capability] = this match
    case v: View =>
      val eval = Option.when(v.evalFormulas)(Capability.Eval)
      val render = Option.when(ReadQuery.isStyled(v.format))(Capability.Render)
      Set(Capability.Values) ++ eval ++ render
    case _ => Set(Capability.Values)

object ReadQuery:

  /** The formats drawn from styles rather than from records: html, svg and the raster exports. */
  def isStyled(format: ViewFormat): Boolean = format match
    case ViewFormat.Markdown | ViewFormat.Json | ViewFormat.Csv => false
    case ViewFormat.Html | ViewFormat.Svg | ViewFormat.Png | ViewFormat.Jpeg | ViewFormat.WebP |
        ViewFormat.Pdf =>
      true

  /**
   * The read query a parsed command denotes, with its pass-through format resolved for `mode`;
   * `None` for every other verb.
   */
  def of(cmd: CliCommand, mode: OutputMode): Option[ReadQuery] = cmd match
    case v: CliCommand.View =>
      Some(
        View(
          v.range,
          v.showFormulas,
          v.evalFormulas,
          v.strict,
          v.limit,
          v.offset,
          v.maxCols,
          CliCommand.viewFormat(v.format, mode),
          v.printScale,
          v.showGridlines,
          v.showLabels,
          v.dpi,
          v.quality,
          v.rasterOutput,
          v.skipEmpty,
          v.headerRow,
          v.rasterizer,
          v.skipHidden
        )
      )
    case CliCommand.Cell(ref, noStyle) => Some(Cell(ref, noStyle))
    case CliCommand.Search(pattern, limit, sheetsFilter, exactTotal) =>
      Some(Search(pattern, limit, sheetsFilter, exactTotal))
    case CliCommand.Stats(ref) => Some(Stats(ref))
    case f: CliCommand.Filter =>
      Some(Filter(f.where, f.columns, f.limit, CliCommand.filterFormat(f.format, mode), f.header))
    case _ => None
