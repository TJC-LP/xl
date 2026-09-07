package com.tjclp.xl.cli.batch

import com.tjclp.xl.addressing.{RefType, SheetName}
import com.tjclp.xl.cli.helpers.BatchParser.BatchOp

/**
 * The single table behind batch parsing (ADR-017 §2.6): one [[OpSpec]] per op name, in the order
 * the unknown-op error has always listed them. `find` resolves names, aliases and spellings;
 * `specOf`/`nameOf` map a parsed `BatchOp` back to its row; `jsonSchema` publishes the shape;
 * `helpText` renders the op table for `batch --help`.
 */
object OpRegistry:

  import FieldKind.*

  private def req(name: String, kind: FieldKind, doc: String): Field =
    Field(name, kind, required = true, doc = doc)

  private def opt(name: String, kind: FieldKind, doc: String, aliases: String*): Field =
    Field(name, kind, required = false, aliases.toVector, doc)

  private val alignValues = Vector("left", "center", "right")
  private val valignValues = Vector("top", "middle", "center", "bottom")
  private val borderValues = Vector("none", "thin", "medium", "thick", "dashed", "dotted", "double")
  private val orientationValues = Vector("portrait", "landscape")
  private val chartTypes = Vector("column", "bar", "line", "pie")
  private val groupings = Vector("clustered", "stacked", "percent-stacked")
  private val legendValues = Vector("right", "left", "top", "bottom", "top-right", "none")

  private val formatDoc =
    "Number format: a name (general, integer, decimal, currency, percent, date, datetime, time, " +
      "text) or an Excel format code. Explicit: it REPLACES the cell's number format (GH-560)."

  private def borderField(name: String, side: String): Field =
    opt(name, Enum(borderValues), s"$side border style")

  private def styleFields: Vector[Field] = Vector(
    req("range", Range, "Cells to style (a single cell is a 1x1 range)"),
    opt("bold", Bool, "Bold font"),
    opt("italic", Bool, "Italic font"),
    opt("underline", Bool, "Underlined font"),
    opt("bg", Color, "Fill color: named, #hex, rgb(r,g,b) or theme:accent1[:tint]"),
    opt("fg", Color, "Font color"),
    opt("fontSize", Num, "Font size in points"),
    opt("fontName", Str, "Font name"),
    opt("align", Enum(alignValues), "Horizontal alignment", "halign"),
    opt("valign", Enum(valignValues), "Vertical alignment"),
    opt("wrap", Bool, "Wrap text"),
    opt(
      "numFormat",
      NumFmt,
      "Number format name or Excel format code (applied as given, a suspect string warns)",
      "format"
    ),
    borderField("border", "All-sides"),
    borderField("borderTop", "Top"),
    borderField("borderRight", "Right"),
    borderField("borderBottom", "Bottom"),
    borderField("borderLeft", "Left"),
    opt("borderColor", Color, "Color of every border named in the op"),
    opt("replace", Bool, "Replace the whole cell style instead of merging into it")
  )

  private val put = OpSpec(
    name = "put",
    aliases = Vector.empty,
    fields = Vector(
      req("ref", Ref, "Target cell, or a range when `values` is given"),
      opt(
        "value",
        Value,
        "The value: JSON number/boolean/null as-is; strings are smart-detected (currency, percent, ISO date) unless `detect` is false"
      ),
      opt("values", Values, "Row-major values for the `ref` range, one per cell"),
      opt("format", NumFmt, formatDoc, "numFormat"),
      opt("detect", Bool, "Smart string detection (default true); false stores strings as text")
    ),
    oneOf = Vector(Set("value"), Set("values")),
    sheetScoped = true,
    cellMutating = true,
    structural = false,
    needsFormula = false,
    streamable = true,
    cliVerb = Some("put"),
    since = "0.1.0",
    doc = "Write a value (or a row-major array of values) with optional number format.",
    example = ujson.Obj(
      "op" -> ujson.Str("put"),
      "ref" -> ujson.Str("A1"),
      "value" -> ujson.Num(1234.5),
      "format" -> ujson.Str("currency")
    )
  )

  private val putf = OpSpec(
    name = "putf",
    aliases = Vector.empty,
    fields = Vector(
      req("ref", Ref, "Target cell, or a range when dragging (`from`) or listing `values`"),
      opt("value", Formula, "The formula (leading = optional)", "formula"),
      opt("values", Formulas, "Row-major formulas for the `ref` range, one per cell, as-is"),
      opt(
        "from",
        Ref,
        "Anchor cell: the formula is dragged across `ref` from here, shifting relative refs like Excel fill-down",
        "anchor"
      ),
      opt("format", NumFmt, formatDoc, "numFormat")
    ),
    oneOf = Vector(Set("value"), Set("values")),
    sheetScoped = true,
    cellMutating = true,
    structural = false,
    needsFormula = true,
    streamable = true,
    cliVerb = Some("putf"),
    since = "0.1.0",
    doc = "Write a formula: one cell, dragged across a range from an anchor, or explicit per cell.",
    example = ujson.Obj(
      "op" -> ujson.Str("putf"),
      "ref" -> ujson.Str("B2:B10"),
      "value" -> ujson.Str("=A2*2"),
      "from" -> ujson.Str("B2"),
      "format" -> ujson.Str("#,##0.0")
    )
  )

  private val style = OpSpec(
    name = "style",
    aliases = Vector.empty,
    fields = styleFields,
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = true,
    cliVerb = Some("style"),
    since = "0.1.0",
    doc = "Style a range: font, fill, alignment, number format, borders (merged unless `replace`).",
    example = ujson.Obj(
      "op" -> ujson.Str("style"),
      "range" -> ujson.Str("A1:D1"),
      "bold" -> ujson.Bool(true),
      "bg" -> ujson.Str("#4472C4"),
      "fg" -> ujson.Str("#FFFFFF"),
      "align" -> ujson.Str("center")
    )
  )

  private def rangeOnly(
    name: String,
    doc: String,
    example: String,
    verb: String
  ): OpSpec = OpSpec(
    name = name,
    aliases = Vector.empty,
    fields = Vector(req("range", Range, "The range")),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = true,
    cliVerb = Some(verb),
    since = "0.1.0",
    doc = doc,
    example = ujson.Obj("op" -> ujson.Str(name), "range" -> ujson.Str(example))
  )

  private val merge = rangeOnly("merge", "Merge a range into one cell.", "A1:D1", "merge")
  private val unmerge = rangeOnly("unmerge", "Unmerge a merged range.", "A1:D1", "unmerge")

  private val colwidth = OpSpec(
    name = "colwidth",
    aliases = Vector.empty,
    fields = Vector(
      req("col", Str, "Column letter"),
      req("width", Num, "Width in character units")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = true,
    cliVerb = Some("col"),
    since = "0.1.0",
    doc = "Set a column's width.",
    example = ujson.Obj(
      "op" -> ujson.Str("colwidth"),
      "col" -> ujson.Str("A"),
      "width" -> ujson.Num(15.5)
    )
  )

  private val rowheight = OpSpec(
    name = "rowheight",
    aliases = Vector.empty,
    fields = Vector(
      req("row", Int, "1-based row number"),
      req("height", Num, "Height in points")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = true,
    cliVerb = Some("row"),
    since = "0.1.0",
    doc = "Set a row's height.",
    example = ujson.Obj(
      "op" -> ujson.Str("rowheight"),
      "row" -> ujson.Num(1),
      "height" -> ujson.Num(30)
    )
  )

  private val comment = OpSpec(
    name = "comment",
    aliases = Vector.empty,
    fields = Vector(
      req("ref", Ref, "The cell"),
      req("text", Str, "Comment text"),
      opt("author", Str, "Comment author")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("comment"),
    since = "0.9.6",
    doc = "Add or replace a cell comment.",
    example = ujson.Obj(
      "op" -> ujson.Str("comment"),
      "ref" -> ujson.Str("A1"),
      "text" -> ujson.Str("Note"),
      "author" -> ujson.Str("User")
    )
  )

  private def refOnly(
    name: String,
    doc: String,
    example: String,
    verb: String,
    since: String,
    streamable: Boolean
  ): OpSpec = OpSpec(
    name = name,
    aliases = Vector.empty,
    fields = Vector(req("ref", Ref, "The cell")),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = streamable,
    cliVerb = Some(verb),
    since = since,
    doc = doc,
    example = ujson.Obj("op" -> ujson.Str(name), "ref" -> ujson.Str(example))
  )

  private val removeComment =
    refOnly(
      "remove-comment",
      "Remove a cell's comment.",
      "A1",
      "remove-comment",
      "0.9.6",
      streamable = false
    )

  private val hyperlink = OpSpec(
    name = "hyperlink",
    aliases = Vector.empty,
    fields = Vector(
      req("ref", Ref, "The cell"),
      opt("target", Str, "URL or internal location; omit to clear the hyperlink", "url")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = None,
    since = "0.10.0",
    doc = "Set or clear a cell hyperlink.",
    example = ujson.Obj(
      "op" -> ujson.Str("hyperlink"),
      "ref" -> ujson.Str("A1"),
      "target" -> ujson.Str("https://example.com")
    )
  )

  private val clear = OpSpec(
    name = "clear",
    aliases = Vector.empty,
    fields = Vector(
      req("range", Range, "The range"),
      opt("all", Bool, "Clear contents, styles and comments"),
      opt("styles", Bool, "Clear styles only"),
      opt("comments", Bool, "Clear comments only")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = true,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("clear"),
    since = "0.9.6",
    doc = "Clear a range: contents by default (unmerging overlaps), or styles/comments/all.",
    example = ujson.Obj(
      "op" -> ujson.Str("clear"),
      "range" -> ujson.Str("A1:B10"),
      "all" -> ujson.Bool(true)
    )
  )

  private def colVisibility(name: String, doc: String): OpSpec = OpSpec(
    name = name,
    aliases = Vector.empty,
    fields = Vector(req("col", Str, "Column letter")),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = true,
    cliVerb = Some("col"),
    since = "0.9.6",
    doc = doc,
    example = ujson.Obj("op" -> ujson.Str(name), "col" -> ujson.Str("C"))
  )

  private def rowVisibility(name: String, doc: String): OpSpec = OpSpec(
    name = name,
    aliases = Vector.empty,
    fields = Vector(req("row", Int, "1-based row number")),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = true,
    cliVerb = Some("row"),
    since = "0.9.6",
    doc = doc,
    example = ujson.Obj("op" -> ujson.Str(name), "row" -> ujson.Num(5))
  )

  private val colHide = colVisibility("col-hide", "Hide a column.")
  private val colShow = colVisibility("col-show", "Show a hidden column.")
  private val rowHide = rowVisibility("row-hide", "Hide a row.")
  private val rowShow = rowVisibility("row-show", "Show a hidden row.")

  private def groupOp(
    name: String,
    axis: String,
    axisDoc: String,
    example: String,
    doc: String
  ): OpSpec = OpSpec(
    name = name,
    aliases = Vector.empty,
    fields = Vector(
      req(axis, Str, axisDoc),
      opt("level", Int, "Outline level (default 1)"),
      opt("collapsed", Bool, "Hide the members and mark the summary row/column after the group")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some(name),
    since = "0.18.0",
    doc = doc,
    example = ujson.Obj(
      "op" -> ujson.Str(name),
      axis -> ujson.Str(example),
      "level" -> ujson.Num(1),
      "collapsed" -> ujson.Bool(false)
    )
  )

  private def ungroupOp(name: String, axis: String, axisDoc: String, example: String): OpSpec =
    OpSpec(
      name = name,
      aliases = Vector.empty,
      fields = Vector(req(axis, Str, axisDoc)),
      oneOf = Vector.empty,
      sheetScoped = true,
      cellMutating = false,
      structural = false,
      needsFormula = false,
      streamable = false,
      cliVerb = Some(name),
      since = "0.18.0",
      doc = s"Clear the outline level and collapse markers of the $axis.",
      example = ujson.Obj("op" -> ujson.Str(name), axis -> ujson.Str(example))
    )

  private val groupRows =
    groupOp(
      "group-rows",
      "rows",
      "Row span, e.g. \"10:20\"",
      "10:20",
      "Group rows into an outline."
    )
  private val groupCols =
    groupOp(
      "group-cols",
      "cols",
      "Column span, e.g. \"E:H\"",
      "E:H",
      "Group columns into an outline."
    )
  private val ungroupRows = ungroupOp("ungroup-rows", "rows", "Row span, e.g. \"10:20\"", "10:20")
  private val ungroupCols = ungroupOp("ungroup-cols", "cols", "Column span, e.g. \"E:H\"", "E:H")

  private val autofit = OpSpec(
    name = "autofit",
    aliases = Vector.empty,
    fields = Vector(
      opt("columns", Str, "A column (\"A\") or span (\"A:F\"); omit for every used column")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("autofit"),
    since = "0.9.6",
    doc = "Auto-fit column widths to their content.",
    example = ujson.Obj("op" -> ujson.Str("autofit"), "columns" -> ujson.Str("A:F"))
  )

  private val addSheet = OpSpec(
    name = "add-sheet",
    aliases = Vector.empty,
    fields = Vector(
      req("name", Sheet, "Name of the new sheet"),
      opt("after", Sheet, "Insert after this sheet (default: at the end)")
    ),
    oneOf = Vector.empty,
    sheetScoped = false,
    cellMutating = false,
    structural = true,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("add-sheet"),
    since = "0.9.6",
    doc = "Add an empty sheet.",
    example = ujson.Obj(
      "op" -> ujson.Str("add-sheet"),
      "name" -> ujson.Str("Summary"),
      "after" -> ujson.Str("Sheet1")
    )
  )

  private val renameSheet = OpSpec(
    name = "rename-sheet",
    aliases = Vector.empty,
    fields = Vector(
      req("from", Sheet, "Current sheet name"),
      req("to", Sheet, "New sheet name")
    ),
    oneOf = Vector.empty,
    sheetScoped = false,
    cellMutating = false,
    structural = true,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("rename-sheet"),
    since = "0.9.6",
    doc =
      "Rename a sheet, rewriting every formula, defined name and rule that refers to it; a rename of the batch's default sheet retargets the ops after it.",
    example = ujson.Obj(
      "op" -> ujson.Str("rename-sheet"),
      "from" -> ujson.Str("Old"),
      "to" -> ujson.Str("New")
    )
  )

  private val freeze =
    refOnly(
      "freeze",
      "Freeze panes above and left of the cell.",
      "B2",
      "freeze",
      "0.10.0",
      streamable = false
    )

  private val unfreeze = OpSpec(
    name = "unfreeze",
    aliases = Vector.empty,
    fields = Vector.empty,
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("unfreeze"),
    since = "0.10.0",
    doc = "Remove the sheet's freeze panes.",
    example = ujson.Obj("op" -> ujson.Str("unfreeze"))
  )

  private val copy = OpSpec(
    name = "copy",
    aliases = Vector.empty,
    fields = Vector(
      req("source", Range, "Source cell or range; may be sheet-qualified"),
      req(
        "target",
        Range,
        "Target cell (expanded to the source's size) or range; may be sheet-qualified"
      ),
      opt("valuesOnly", Bool, "Paste cached values instead of formulas")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = true,
    structural = false,
    needsFormula = true,
    streamable = false,
    cliVerb = Some("copy"),
    since = "0.10.0",
    doc =
      "Copy a range, shifting relative references like Excel; each side may name its own sheet.",
    example = ujson.Obj(
      "op" -> ujson.Str("copy"),
      "source" -> ujson.Str("A1:B2"),
      "target" -> ujson.Str("D1"),
      "valuesOnly" -> ujson.Bool(false)
    )
  )

  private val chart = OpSpec(
    name = "chart",
    aliases = Vector.empty,
    fields = Vector(
      req("type", Enum(chartTypes), "Chart type"),
      opt("grouping", Enum(groupings), "Bar/column grouping"),
      req("data", Range, "Values range; may be sheet-qualified"),
      opt("categories", Range, "Categories range; may be sheet-qualified"),
      opt("seriesNames", Str, "Comma-separated series names, positional"),
      opt("seriesColors", Str, "Comma-separated series colors, positional"),
      opt("title", Str, "Chart title"),
      opt("legend", Enum(legendValues), "Legend position"),
      req("at", Range, "Placement: an anchor cell or the range the chart covers")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("chart add"),
    since = "0.15.0",
    doc = "Add a chart, mirroring `chart add`.",
    example = ujson.Obj(
      "op" -> ujson.Str("chart"),
      "type" -> ujson.Str("column"),
      "data" -> ujson.Str("B2:C4"),
      "categories" -> ujson.Str("A2:A4"),
      "seriesNames" -> ujson.Str("North,South"),
      "title" -> ujson.Str("Sales"),
      "legend" -> ujson.Str("right"),
      "at" -> ujson.Str("E2:K15")
    )
  )

  private val sheetView = OpSpec(
    name = "sheet-view",
    aliases = Vector.empty,
    fields = Vector(
      opt("gridlines", Bool, "Show gridlines"),
      opt("zoom", Int, "Zoom percentage"),
      opt("tabSelected", Bool, "Select this sheet's tab")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("sheet-view"),
    since = "0.13.0",
    doc = "Sheet view settings: gridlines, zoom, selected tab.",
    example = ujson.Obj(
      "op" -> ujson.Str("sheet-view"),
      "gridlines" -> ujson.Bool(false),
      "zoom" -> ujson.Num(85)
    )
  )

  private val tabColor = OpSpec(
    name = "tab-color",
    aliases = Vector.empty,
    fields = Vector(
      opt("color", Color, "Tab color: named, #hex, rgb(r,g,b) or theme:accent1[:tint]"),
      opt("clear", Bool, "Remove the tab color")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("tab-color"),
    since = "0.13.0",
    doc = "Set or clear the sheet tab color.",
    example = ujson.Obj("op" -> ujson.Str("tab-color"), "color" -> ujson.Str("#1F4E79"))
  )

  private val autofilter = OpSpec(
    name = "autofilter",
    aliases = Vector.empty,
    fields = Vector(
      opt("range", Range, "The filter range; may be sheet-qualified"),
      opt("clear", Bool, "Strip the sheet's autoFilter, even one preserved from the source file")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("autofilter"),
    since = "0.18.0",
    doc = "Set or clear the sheet-level autoFilter.",
    example = ujson.Obj("op" -> ujson.Str("autofilter"), "range" -> ujson.Str("A1:M29"))
  )

  private val pageSetup = OpSpec(
    name = "page-setup",
    aliases = Vector.empty,
    fields = Vector(
      opt("orientation", Enum(orientationValues), "Page orientation"),
      opt("scale", Int, "Print scale, 10-400"),
      opt("fitToWidth", Int, "Pages wide; 0 = automatic"),
      opt("fitToHeight", Int, "Pages tall; 0 = automatic (GH-463)"),
      opt("fitToPage", Bool, "Force the fitToPage flag on or off; omit to derive it")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("page-setup"),
    since = "0.13.0",
    doc = "Print setup: orientation, scale, fit-to pages.",
    example = ujson.Obj(
      "op" -> ujson.Str("page-setup"),
      "orientation" -> ujson.Str("landscape"),
      "fitToWidth" -> ujson.Num(1),
      "fitToHeight" -> ujson.Num(0)
    )
  )

  private val headerFooter = OpSpec(
    name = "header-footer",
    aliases = Vector.empty,
    fields = Vector(
      opt(
        "oddHeader",
        Str,
        "Header (&L/&C/&R sections; &P page, &N total, &D date, &F file, &A sheet)"
      ),
      opt("oddFooter", Str, "Footer"),
      opt("evenHeader", Str, "Even-page header (sets differentOddEven)"),
      opt("evenFooter", Str, "Even-page footer (sets differentOddEven)"),
      opt("firstHeader", Str, "First-page header (sets differentFirst)"),
      opt("firstFooter", Str, "First-page footer (sets differentFirst)"),
      opt("differentOddEven", Bool, "Force the odd/even flag on"),
      opt("differentFirst", Bool, "Force the first-page flag on")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("header-footer"),
    since = "0.13.0",
    doc = "Print headers and footers.",
    example = ujson.Obj(
      "op" -> ujson.Str("header-footer"),
      "oddFooter" -> ujson.Str("&LConfidential&RPage &P of &N")
    )
  )

  private val cf = OpSpec(
    name = "cf",
    aliases = Vector.empty,
    fields = Vector(
      req("range", Range, "The range the rule applies to"),
      req(
        "rule",
        Str,
        "Rule DSL as `cf add`: cellIs:greaterThan:100, between:1:10, expression:…, colorScale:…, dataBar:…, top10:…, text:…"
      ),
      opt("bold", Bool, "Differential style: bold"),
      opt("italic", Bool, "Differential style: italic"),
      opt("underline", Bool, "Differential style: underline"),
      opt("strike", Bool, "Differential style: strikethrough"),
      opt("bg", Color, "Differential style: fill color"),
      opt("fg", Color, "Differential style: font color")
    ),
    oneOf = Vector.empty,
    sheetScoped = true,
    cellMutating = false,
    structural = false,
    needsFormula = false,
    streamable = false,
    cliVerb = Some("cf add"),
    since = "0.13.0",
    doc = "Add a conditional-formatting rule; priorities are assigned in order.",
    example = ujson.Obj(
      "op" -> ujson.Str("cf"),
      "range" -> ujson.Str("A1:A10"),
      "rule" -> ujson.Str("cellIs:greaterThan:100"),
      "bold" -> ujson.Bool(true),
      "bg" -> ujson.Str("#FFC7CE"),
      "fg" -> ujson.Str("#9C0006")
    )
  )

  /** The 32 ops, in the order the unknown-op error has always listed them. */
  val all: Vector[OpSpec] = Vector(
    put,
    putf,
    style,
    merge,
    unmerge,
    colwidth,
    rowheight,
    comment,
    removeComment,
    hyperlink,
    clear,
    colHide,
    colShow,
    rowHide,
    rowShow,
    groupRows,
    groupCols,
    ungroupRows,
    ungroupCols,
    autofit,
    addSheet,
    renameSheet,
    freeze,
    unfreeze,
    copy,
    chart,
    sheetView,
    tabColor,
    autofilter,
    pageSetup,
    headerFooter,
    cf
  )

  private val byNormalizedName: Map[String, OpSpec] =
    all.flatMap(spec => (spec.name +: spec.aliases).map(n => Spelling.normalize(n) -> spec)).toMap

  /** Resolve an op name: exact, alias, or any case/separator spelling (`removeComment`). */
  def find(name: String): Option[OpSpec] = byNormalizedName.get(Spelling.normalize(name))

  /**
   * The row for a parsed op. Total and exhaustive on purpose: a new `BatchOp` case must be placed
   * in the table here, or `OpRegistrySpec` fails on it.
   */
  def specOf(op: BatchOp): OpSpec = op match
    case _: BatchOp.Put | _: BatchOp.PutValues => put
    case _: BatchOp.PutFormula | _: BatchOp.PutFormulaDragging | _: BatchOp.PutFormulas => putf
    case _: BatchOp.Style => style
    case _: BatchOp.Merge => merge
    case _: BatchOp.Unmerge => unmerge
    case _: BatchOp.ColWidth => colwidth
    case _: BatchOp.RowHeight => rowheight
    case _: BatchOp.AddComment => comment
    case _: BatchOp.RemoveComment => removeComment
    case _: BatchOp.Hyperlink => hyperlink
    case _: BatchOp.Clear => clear
    case _: BatchOp.ColHide => colHide
    case _: BatchOp.ColShow => colShow
    case _: BatchOp.RowHide => rowHide
    case _: BatchOp.RowShow => rowShow
    case _: BatchOp.GroupRows => groupRows
    case _: BatchOp.GroupCols => groupCols
    case _: BatchOp.UngroupRows => ungroupRows
    case _: BatchOp.UngroupCols => ungroupCols
    case _: BatchOp.AutoFit => autofit
    case _: BatchOp.AddSheet => addSheet
    case _: BatchOp.RenameSheet => renameSheet
    case _: BatchOp.Freeze => freeze
    case BatchOp.Unfreeze => unfreeze
    case _: BatchOp.CopyRange => copy
    case _: BatchOp.AddChart => chart
    case _: BatchOp.SetSheetView => sheetView
    case _: BatchOp.SetTabColor => tabColor
    case _: BatchOp.SetAutoFilter => autofilter
    case _: BatchOp.SetPageSetup => pageSetup
    case _: BatchOp.SetHeaderFooter => headerFooter
    case _: BatchOp.AddConditionalFormat => cf

  /** The op name of a parsed op (`PutFormulaDragging` → `putf`). */
  def nameOf(op: BatchOp): String = specOf(op).name

  /**
   * The ref strings that name the op's own target cells — the ones a `sheet` key must agree with.
   * `copy` and `chart` are absent on purpose: their qualified sides may legitimately name other
   * sheets (a cross-sheet copy, chart data on another sheet).
   */
  def targetRefs(op: BatchOp): Vector[String] = op match
    case BatchOp.Put(ref, _, _) => Vector(ref)
    case BatchOp.PutFormula(ref, _, _) => Vector(ref)
    case BatchOp.PutFormulaDragging(range, _, _, _) => Vector(range)
    case BatchOp.PutFormulas(range, _, _) => Vector(range)
    case BatchOp.PutValues(range, _) => Vector(range)
    case BatchOp.Style(range, _) => Vector(range)
    case BatchOp.Merge(range) => Vector(range)
    case BatchOp.Unmerge(range) => Vector(range)
    case BatchOp.AddComment(ref, _, _) => Vector(ref)
    case BatchOp.RemoveComment(ref) => Vector(ref)
    case BatchOp.Hyperlink(ref, _) => Vector(ref)
    case BatchOp.Clear(range, _, _, _) => Vector(range)
    case BatchOp.Freeze(ref) => Vector(ref)
    case BatchOp.SetAutoFilter(range, _) => range.toList.toVector
    case BatchOp.AddConditionalFormat(range, _, _, _, _, _, _, _) => Vector(range)
    case _ => Vector.empty

  /** The sheet a ref string is qualified with, if any (`Other!A1` → `Other`). */
  def qualifiedSheet(refStr: String): Option[SheetName] =
    RefType.parse(refStr).toOption.collect {
      case RefType.QualifiedCell(sheet, _) => sheet
      case RefType.QualifiedRange(sheet, _) => sheet
    }

  /**
   * The same op with the sheet qualifier stripped from its target refs, for a writer that already
   * knows its worksheet (the streaming path). Only the streamable ref-carrying cases change.
   */
  def unqualified(op: BatchOp): BatchOp =
    def bare(refStr: String): String = RefType.parse(refStr) match
      case Right(RefType.QualifiedCell(_, ref)) => ref.toA1
      case Right(RefType.QualifiedRange(_, range)) => range.toA1
      case _ => refStr
    op match
      case o: BatchOp.Put => o.copy(ref = bare(o.ref))
      case o: BatchOp.PutFormula => o.copy(ref = bare(o.ref))
      case o: BatchOp.PutFormulaDragging => o.copy(range = bare(o.range), from = bare(o.from))
      case o: BatchOp.PutFormulas => o.copy(range = bare(o.range))
      case o: BatchOp.PutValues => o.copy(range = bare(o.range))
      case o: BatchOp.Style => o.copy(range = bare(o.range))
      case o: BatchOp.Merge => o.copy(range = bare(o.range))
      case o: BatchOp.Unmerge => o.copy(range = bare(o.range))
      case other => other

  // ========== JSON Schema ==========

  private def kindSchema(kind: FieldKind): ujson.Obj = kind match
    case Str | Ref | Range | Sheet | Color | NumFmt | Formula =>
      ujson.Obj("type" -> ujson.Str("string"), "x-kind" -> ujson.Str(kind.toString))
    case Num => ujson.Obj("type" -> ujson.Str("number"))
    case Int => ujson.Obj("type" -> ujson.Str("integer"))
    case Bool => ujson.Obj("type" -> ujson.Str("boolean"))
    case Value =>
      ujson.Obj(
        "type" -> ujson.Arr(
          ujson.Str("string"),
          ujson.Str("number"),
          ujson.Str("boolean"),
          ujson.Str("null")
        )
      )
    case Values => ujson.Obj("type" -> ujson.Str("array"), "items" -> kindSchema(Value))
    case Formulas =>
      ujson.Obj("type" -> ujson.Str("array"), "items" -> ujson.Obj("type" -> ujson.Str("string")))
    case Enum(values) =>
      ujson.Obj("type" -> ujson.Str("string"), "enum" -> ujson.Arr.from(values.map(ujson.Str(_))))
    case Arr(of) => ujson.Obj("type" -> ujson.Str("array"), "items" -> kindSchema(of))

  private def withEntries(base: ujson.Obj, extra: Vector[(String, ujson.Value)]): ujson.Obj =
    ujson.Obj.from(base.value.toVector ++ extra)

  private val sheetProperty: ujson.Obj = ujson.Obj(
    "type" -> ujson.Str("string"),
    "description" -> ujson.Str(
      "Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet"
    )
  )

  private def fieldProperties(field: Field): Vector[(String, ujson.Value)] =
    val docEntries =
      if field.doc.isEmpty then Vector.empty
      else Vector("description" -> ujson.Str(field.doc))
    val aliasEntries =
      if field.aliases.isEmpty then Vector.empty
      else Vector("x-aliases" -> ujson.Arr.from(field.aliases.map(ujson.Str(_))))
    val canonical = field.name -> withEntries(kindSchema(field.kind), docEntries ++ aliasEntries)
    val aliases = field.lookupOrder.drop(1).map { alt =>
      alt -> withEntries(kindSchema(field.kind), Vector("x-alias-of" -> ujson.Str(field.name)))
    }
    canonical +: aliases

  /** One `oneOf` group: exactly one of its fields (through any spelling) must be present. */
  private def groupSchema(spec: OpSpec, group: Set[String]): ujson.Obj =
    val names = group.toVector.sorted
    if names.size == 1 then
      val spellings = names.flatMap(n => spec.field(n).toList).flatMap(_.lookupOrder)
      ujson.Obj(
        "anyOf" -> ujson.Arr.from(
          spellings.map(s => ujson.Obj("required" -> ujson.Arr(ujson.Str(s))))
        )
      )
    else ujson.Obj("required" -> ujson.Arr.from(names.map(ujson.Str(_))))

  private def opSchema(spec: OpSpec): ujson.Obj =
    val opProperty = Vector("op" -> (ujson.Obj("const" -> ujson.Str(spec.name)): ujson.Value))
    val scope: Vector[(String, ujson.Value)] =
      if spec.sheetScoped then Vector("sheet" -> sheetProperty) else Vector.empty
    val properties = opProperty ++ scope ++ spec.fields.flatMap(fieldProperties)
    val required = "op" +: spec.fields.filter(_.required).map(_.name)
    val base: Vector[(String, ujson.Value)] = Vector(
      "title" -> ujson.Str(spec.name),
      "description" -> ujson.Str(spec.doc),
      "type" -> ujson.Str("object"),
      "properties" -> ujson.Obj.from(properties),
      "required" -> ujson.Arr.from(required.map(ujson.Str(_))),
      "x-aliases" -> ujson.Arr.from(spec.aliases.map(ujson.Str(_))),
      "x-streamable" -> ujson.Bool(spec.streamable),
      "x-cellMutating" -> ujson.Bool(spec.cellMutating),
      "x-sheetScoped" -> ujson.Bool(spec.sheetScoped),
      "x-structural" -> ujson.Bool(spec.structural),
      "x-needsFormula" -> ujson.Bool(spec.needsFormula),
      "x-cliVerb" -> spec.cliVerb.fold[ujson.Value](ujson.Null)(ujson.Str(_)),
      "x-since" -> ujson.Str(spec.since),
      "x-example" -> spec.example
    )
    val oneOf: Vector[(String, ujson.Value)] =
      if spec.oneOf.isEmpty then Vector.empty
      else Vector("oneOf" -> ujson.Arr.from(spec.oneOf.map(groupSchema(spec, _))))
    ujson.Obj.from(base ++ oneOf)

  /**
   * The batch document's JSON Schema (draft 2020-12): an array whose items are exactly one of the
   * per-op object schemas (discriminated by the `op` const), enums inlined, with `x-aliases`,
   * `x-streamable`, `x-cellMutating`, `x-cliVerb` and `x-example` annotations per op.
   */
  def jsonSchema(version: String): ujson.Value =
    ujson.Obj(
      "$schema" -> ujson.Str("https://json-schema.org/draft/2020-12/schema"),
      "$id" -> ujson.Str(s"https://github.com/TJC-LP/xl/schema/batch-$version.json"),
      "title" -> ujson.Str("xl batch operations"),
      "description" -> ujson.Str(
        "A JSON array of operations `xl batch` applies in order. Every sheet-scoped op accepts an " +
          "optional \"sheet\" key; property names are accepted in camelCase or kebab-case."
      ),
      "x-version" -> ujson.Str(version),
      "type" -> ujson.Str("array"),
      "items" -> ujson.Obj("oneOf" -> ujson.Arr.from(all.map(opSchema)))
    )

  // ========== Help ==========

  /** The op table for `batch --help`: one line per op, its example, then the shared rules. */
  def helpText: String =
    val width = all.foldLeft(0)((w, spec) => math.max(w, spec.name.length))
    val rows = all.map { spec =>
      val name = spec.name.padTo(width, ' ')
      val aliases = spec.fields.flatMap(f => f.aliases.map(a => s"${f.name}|$a"))
      val aliasNote = if aliases.isEmpty then "" else s"  (aliases: ${aliases.mkString(", ")})"
      val stream = if spec.streamable then "" else "  [not with --stream]"
      s"  $name  ${ujson.write(spec.example)}$aliasNote$stream"
    }
    (Vector("OPERATIONS:") ++ rows ++ Vector(
      "",
      "Every op except add-sheet/rename-sheet accepts \"sheet\": the sheet for its unqualified refs",
      "(a sheet-qualified ref wins, then \"sheet\", then -s/--sheet). Property names are accepted in",
      "camelCase or kebab-case. put/putf \"format\" is explicit and replaces the cell's number format."
    )).mkString("\n")
