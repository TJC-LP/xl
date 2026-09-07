package com.tjclp.xl.ops

import java.util.Locale

/** The wire shape of one [[EditField]] (the `FieldKind` of ADR-017 §2.6, extended for spans). */
enum FieldKind derives CanEqual:
  case Str, Num, Int, Bool, Ref, Range, Sheet, Column, Row, ColSpan, RowSpan, Color, NumFmt,
    Formula, Value, Values, Formulas
  case Enum(values: Vector[String])
  case Arr(of: FieldKind)

/**
 * One property of an edit: `name` is the canonical camelCase key, `aliases` the other accepted
 * names; every name is also accepted in its kebab-case spelling (the `OpSpec.Field` shape).
 */
final case class EditField(
  name: String,
  kind: FieldKind,
  required: Boolean,
  aliases: Vector[String] = Vector.empty,
  doc: String = ""
) derives CanEqual:

  /**
   * Every key this field answers to: canonical and aliases, each in given, kebab and camel form.
   */
  def spellings: Set[String] = (name +: aliases).flatMap(EditSchema.variants).toSet

/**
 * The metadata row of one [[Edit]] case — the `OpSpec` shape of ADR-017 §2.6 without the JSON
 * example (xl-core is JSON-free; the batch codec of Wave 2.2 owns examples). `name` is the kebab
 * form of the case; `batchOp` the Wave 1 batch op the case is a form of (`put-values` → `put`);
 * `cliVerb` the verb that lowers to it. `idempotent` declares the class the idempotence law is
 * asserted for.
 */
final case class EditSpec(
  name: String,
  aliases: Vector[String],
  batchOp: Option[String],
  cliVerb: Option[String],
  fields: Vector[EditField],
  oneOf: Vector[Set[String]],
  sheetScoped: Boolean,
  cellMutating: Boolean,
  structural: Boolean,
  needsFormula: Boolean,
  streamable: Boolean,
  idempotent: Boolean,
  since: String,
  doc: String
) derives CanEqual:

  def field(name: String): Option[EditField] = fields.find(_.name == name)

  /** The canonical property names, the `sheet` key included when the edit takes one. */
  def knownProperties: Vector[String] =
    val scope = if sheetScoped then Vector("sheet") else Vector.empty
    ("op" +: scope) ++ fields.map(_.name)

/**
 * The single metadata table of the algebra: one [[EditSpec]] per [[Edit]] case, in declaration
 * order (`EditModelSpec` pins `all.size` to the case count and the names to the case labels).
 */
object EditSchema:

  import FieldKind.{
    Column as ColumnKind,
    Row as RowKind,
    Sheet as SheetKind,
    ColSpan as ColSpanKind,
    RowSpan as RowSpanKind,
    *
  }

  /** `PutFormula` → `put-formula`; already-kebab names are unchanged. */
  def kebab(s: String): String =
    s.replaceAll("([a-z0-9])([A-Z])", "$1-$2").toLowerCase(Locale.ROOT)

  /** `put-formula` → `putFormula`; names without a hyphen are unchanged. */
  def camel(s: String): String = s.split('-').toList match
    case head :: tail => head + tail.map(_.capitalize).mkString
    case Nil => s

  /** The name itself plus its kebab and camel forms, de-duplicated. */
  def variants(s: String): Vector[String] = Vector(s, kebab(s), camel(s)).distinct

  /** Case- and separator-insensitive key for [[find]]. */
  def normalize(s: String): String =
    s.toLowerCase(Locale.ROOT).replace("-", "").replace("_", "")

  private def req(name: String, kind: FieldKind, doc: String, aliases: String*): EditField =
    EditField(name, kind, required = true, aliases.toVector, doc)

  private def opt(name: String, kind: FieldKind, doc: String, aliases: String*): EditField =
    EditField(name, kind, required = false, aliases.toVector, doc)

  private val formatDoc =
    "Number format: a name (general, integer, decimal, currency, percent, date, datetime, time, " +
      "text) or an Excel format code. Explicit: it REPLACES the cell's number format (GH-560)."

  private val alignValues = Vector("left", "center", "right")
  private val valignValues = Vector("top", "middle", "center", "bottom")
  private val borderValues = Vector("none", "thin", "medium", "thick", "dashed", "dotted", "double")
  private val orientationValues = Vector("portrait", "landscape")

  private def spec(
    name: String,
    fields: Vector[EditField],
    doc: String,
    since: String,
    batchOp: Option[String] = None,
    cliVerb: Option[String] = None,
    oneOf: Vector[Set[String]] = Vector.empty,
    sheetScoped: Boolean = true,
    cellMutating: Boolean = false,
    structural: Boolean = false,
    needsFormula: Boolean = false,
    streamable: Boolean = false,
    idempotent: Boolean = true
  ): EditSpec =
    EditSpec(
      name = name,
      aliases = Vector.empty,
      batchOp = batchOp,
      cliVerb = cliVerb,
      fields = fields,
      oneOf = oneOf,
      sheetScoped = sheetScoped,
      cellMutating = cellMutating,
      structural = structural,
      needsFormula = needsFormula,
      streamable = streamable,
      idempotent = idempotent,
      since = since,
      doc = doc
    )

  private val put = spec(
    "put",
    Vector(
      req("ref", Ref, "Target cell"),
      req("value", Value, "The value"),
      opt("format", NumFmt, formatDoc, "numFormat")
    ),
    "Write one value with an optional number-format hint.",
    since = "0.1.0",
    batchOp = Some("put"),
    cliVerb = Some("put"),
    cellMutating = true,
    streamable = true
  )

  private val putValues = spec(
    "put-values",
    Vector(
      req("ref", Range, "Target range"),
      req("values", Values, "Row-major values for the range, one per cell"),
      opt("format", NumFmt, formatDoc, "numFormat")
    ),
    "Write a row-major array of values over a range.",
    since = "0.1.0",
    batchOp = Some("put"),
    cliVerb = Some("put"),
    cellMutating = true,
    streamable = true
  )

  private val putFormula = spec(
    "put-formula",
    Vector(
      req("ref", Ref, "Target cell"),
      req("value", Formula, "The formula (leading = optional)", "formula"),
      opt("format", NumFmt, formatDoc, "numFormat")
    ),
    "Write one formula, stored uncached.",
    since = "0.1.0",
    batchOp = Some("putf"),
    cliVerb = Some("putf"),
    cellMutating = true,
    needsFormula = true,
    streamable = true
  )

  private val putFormulas = spec(
    "put-formulas",
    Vector(
      req("ref", Range, "Target range"),
      req("values", Formulas, "Row-major formulas for the range, one per cell, as-is"),
      opt("format", NumFmt, formatDoc, "numFormat")
    ),
    "Write explicit per-cell formulas over a range.",
    since = "0.1.0",
    batchOp = Some("putf"),
    cliVerb = Some("putf"),
    cellMutating = true,
    needsFormula = true,
    streamable = true
  )

  private val dragFormula = spec(
    "drag-formula",
    Vector(
      req("ref", Range, "Target range"),
      req("value", Formula, "The formula to drag (leading = optional)", "formula"),
      req("from", Ref, "Anchor cell: relative refs shift like Excel fill-down", "anchor"),
      opt("format", NumFmt, formatDoc, "numFormat")
    ),
    "Drag one formula across a range from an anchor cell.",
    since = "0.1.0",
    batchOp = Some("putf"),
    cliVerb = Some("putf"),
    cellMutating = true,
    needsFormula = true,
    streamable = true
  )

  private val fill = spec(
    "fill",
    Vector(
      req("source", Range, "Source cell or range"),
      req("target", Range, "Target range on the same sheet"),
      opt("direction", Enum(Vector("down", "right")), "Fill direction (default down)")
    ),
    "Repeat the source down or right through the target, shifting relative references.",
    since = "0.9.6",
    cliVerb = Some("fill"),
    cellMutating = true,
    needsFormula = true
  )

  private val copy = spec(
    "copy",
    Vector(
      req("source", Range, "Source cell or range; may be sheet-qualified"),
      req("target", Ref, "Target top-left cell; may be sheet-qualified"),
      opt("valuesOnly", Bool, "Paste cached values instead of formulas")
    ),
    "Copy a range, shifting relative references like Excel; each side may name its own sheet.",
    since = "0.10.0",
    batchOp = Some("copy"),
    cliVerb = Some("copy"),
    cellMutating = true,
    needsFormula = true,
    idempotent = false
  )

  private val sort = spec(
    "sort",
    Vector(
      req("range", Range, "The range to sort"),
      req(
        "keys",
        Arr(Str),
        "Sort keys, e.g. \"A\", \"B:desc\", \"C:asc:numeric\" — columns inside the range"
      ),
      opt("header", Bool, "Keep the first row in place")
    ),
    "Sort the rows of a range by one or more columns.",
    since = "0.9.6",
    cliVerb = Some("sort"),
    cellMutating = true,
    needsFormula = true
  )

  private val clear = spec(
    "clear",
    Vector(
      req("range", Range, "The range"),
      opt("all", Bool, "Clear contents, styles and comments"),
      opt("styles", Bool, "Clear styles only"),
      opt("comments", Bool, "Clear comments only")
    ),
    "Clear a range: contents by default (unmerging overlaps), or styles/comments/all.",
    since = "0.9.6",
    batchOp = Some("clear"),
    cliVerb = Some("clear"),
    cellMutating = true
  )

  private def borderField(name: String, side: String): EditField =
    opt(name, Enum(borderValues), s"$side border style")

  private val style = spec(
    "style",
    Vector(
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
      opt("numFormat", NumFmt, "Number format name or Excel format code", "format"),
      borderField("border", "All-sides"),
      borderField("borderTop", "Top"),
      borderField("borderRight", "Right"),
      borderField("borderBottom", "Bottom"),
      borderField("borderLeft", "Left"),
      opt("borderColor", Color, "Color of every border named in the op"),
      opt("replace", Bool, "Replace the whole cell style instead of merging into it")
    ),
    "Style a range: font, fill, alignment, number format, borders (merged unless `replace`).",
    since = "0.1.0",
    batchOp = Some("style"),
    cliVerb = Some("style"),
    streamable = true
  )

  private def rangeOnly(name: String, doc: String, verb: String): EditSpec =
    spec(
      name,
      Vector(req("range", Range, "The range")),
      doc,
      since = "0.1.0",
      batchOp = Some(name),
      cliVerb = Some(verb),
      streamable = true
    )

  private val merge = rangeOnly("merge", "Merge a range into one cell.", "merge")
  private val unmerge = rangeOnly("unmerge", "Unmerge a merged range.", "unmerge")

  private val colWidth = spec(
    "col-width",
    Vector(
      req("col", ColSpanKind, "Column letter or span"),
      req("width", Num, "Width in character units")
    ),
    "Set a column's width.",
    since = "0.1.0",
    batchOp = Some("colwidth"),
    cliVerb = Some("col"),
    streamable = true
  )

  private val rowHeight = spec(
    "row-height",
    Vector(
      req("row", RowSpanKind, "1-based row number or span"),
      req("height", Num, "Height in points")
    ),
    "Set a row's height.",
    since = "0.1.0",
    batchOp = Some("rowheight"),
    cliVerb = Some("row"),
    streamable = true
  )

  private def colVisibility(name: String, batch: String, doc: String): EditSpec =
    spec(
      name,
      Vector(req("col", ColSpanKind, "Column letter or span")),
      doc,
      since = "0.9.6",
      batchOp = Some(batch),
      cliVerb = Some("col"),
      streamable = true
    )

  private def rowVisibility(name: String, batch: String, doc: String): EditSpec =
    spec(
      name,
      Vector(req("row", RowSpanKind, "1-based row number or span")),
      doc,
      since = "0.9.6",
      batchOp = Some(batch),
      cliVerb = Some("row"),
      streamable = true
    )

  private val hideCols = colVisibility("hide-cols", "col-hide", "Hide columns.")
  private val showCols = colVisibility("show-cols", "col-show", "Show hidden columns.")
  private val hideRows = rowVisibility("hide-rows", "row-hide", "Hide rows.")
  private val showRows = rowVisibility("show-rows", "row-show", "Show hidden rows.")

  private def groupSpec(name: String, axis: String, kind: FieldKind, axisDoc: String): EditSpec =
    spec(
      name,
      Vector(
        req(axis, kind, axisDoc),
        opt("level", Int, "Outline level (default 1)"),
        opt("collapsed", Bool, "Hide the members and mark the summary row/column after the group")
      ),
      s"Group ${if axis == "rows" then "rows" else "columns"} into an outline.",
      since = "0.18.0",
      batchOp = Some(name),
      cliVerb = Some(name)
    )

  private def ungroupSpec(name: String, axis: String, kind: FieldKind, axisDoc: String): EditSpec =
    spec(
      name,
      Vector(req(axis, kind, axisDoc)),
      s"Clear the outline level and collapse markers of the $axis.",
      since = "0.18.0",
      batchOp = Some(name),
      cliVerb = Some(name)
    )

  private val groupRows = groupSpec("group-rows", "rows", RowSpanKind, "Row span, e.g. \"10:20\"")
  private val groupCols = groupSpec("group-cols", "cols", ColSpanKind, "Column span, e.g. \"E:H\"")
  private val ungroupRows =
    ungroupSpec("ungroup-rows", "rows", RowSpanKind, "Row span, e.g. \"10:20\"")
  private val ungroupCols =
    ungroupSpec("ungroup-cols", "cols", ColSpanKind, "Column span, e.g. \"E:H\"")

  private val autoFit = spec(
    "auto-fit",
    Vector(
      opt("columns", ColSpanKind, "A column (\"A\") or span (\"A:F\"); omit for every used column")
    ),
    "Auto-fit column widths to their content.",
    since = "0.9.6",
    batchOp = Some("autofit"),
    cliVerb = Some("autofit")
  )

  private val setComment = spec(
    "set-comment",
    Vector(
      req("ref", Ref, "The cell"),
      req("text", Str, "Comment text"),
      opt("author", Str, "Comment author")
    ),
    "Add or replace a cell comment.",
    since = "0.9.6",
    batchOp = Some("comment"),
    cliVerb = Some("comment")
  )

  private val removeComment = spec(
    "remove-comment",
    Vector(req("ref", Ref, "The cell")),
    "Remove a cell's comment.",
    since = "0.9.6",
    batchOp = Some("remove-comment"),
    cliVerb = Some("remove-comment")
  )

  private val hyperlink = spec(
    "hyperlink",
    Vector(
      req("ref", Ref, "The cell"),
      opt("target", Str, "URL or internal location; omit to clear the hyperlink", "url")
    ),
    "Set or clear a cell hyperlink.",
    since = "0.10.0",
    batchOp = Some("hyperlink")
  )

  private val addConditionalFormat = spec(
    "add-conditional-format",
    Vector(
      req("range", Range, "The range the rule applies to"),
      req("rule", Str, "Rule DSL as `cf add`: cellIs:greaterThan:100, between:1:10, expression:…"),
      opt("bold", Bool, "Differential style: bold"),
      opt("italic", Bool, "Differential style: italic"),
      opt("underline", Bool, "Differential style: underline"),
      opt("strike", Bool, "Differential style: strikethrough"),
      opt("bg", Color, "Differential style: fill color"),
      opt("fg", Color, "Differential style: font color")
    ),
    "Add a conditional-formatting block; priorities are assigned in order.",
    since = "0.13.0",
    batchOp = Some("cf"),
    cliVerb = Some("cf add"),
    idempotent = false
  )

  private val addChart = spec(
    "add-chart",
    Vector(
      req("type", Enum(Vector("column", "bar", "line", "pie")), "Chart type"),
      opt("grouping", Enum(Vector("clustered", "stacked", "percent-stacked")), "Bar grouping"),
      req("data", Range, "Values range; may be sheet-qualified"),
      opt("categories", Range, "Categories range; may be sheet-qualified"),
      opt("seriesNames", Str, "Comma-separated series names, positional"),
      opt("seriesColors", Str, "Comma-separated series colors, positional"),
      opt("title", Str, "Chart title"),
      opt("legend", Enum(Vector("right", "left", "top", "bottom", "top-right", "none")), "Legend"),
      req("at", Range, "Placement: an anchor cell or the range the chart covers")
    ),
    "Add a chart, mirroring `chart add`.",
    since = "0.15.0",
    batchOp = Some("chart"),
    cliVerb = Some("chart add"),
    idempotent = false
  )

  private val addImage = spec(
    "add-image",
    Vector(
      req("path", Str, "Image file (png, jpeg, gif, bmp)"),
      req("at", Range, "Anchor cell (natural size) or the range the image covers"),
      opt("size", Str, "WxH in pixels for a cell anchor, e.g. 320x240")
    ),
    "Embed a picture at a cell or over a range.",
    since = "0.11.0",
    cliVerb = Some("add-image"),
    idempotent = false
  )

  private val freeze = spec(
    "freeze",
    Vector(req("ref", Ref, "The cell: rows above and columns left are frozen")),
    "Freeze panes above and left of the cell.",
    since = "0.10.0",
    batchOp = Some("freeze"),
    cliVerb = Some("freeze")
  )

  private val unfreeze = spec(
    "unfreeze",
    Vector.empty,
    "Remove the sheet's freeze panes.",
    since = "0.10.0",
    batchOp = Some("unfreeze"),
    cliVerb = Some("unfreeze")
  )

  private val setSheetView = spec(
    "set-sheet-view",
    Vector(
      opt("gridlines", Bool, "Show gridlines"),
      opt("zoom", Int, "Zoom percentage (10-400)"),
      opt("tabSelected", Bool, "Select this sheet's tab")
    ),
    "Sheet view settings: gridlines, zoom, selected tab.",
    since = "0.13.0",
    batchOp = Some("sheet-view"),
    cliVerb = Some("sheet-view")
  )

  private val setTabColor = spec(
    "set-tab-color",
    Vector(
      opt("color", Color, "Tab color: named, #hex, rgb(r,g,b) or theme:accent1[:tint]"),
      opt("clear", Bool, "Remove the tab color")
    ),
    "Set or clear the sheet tab color.",
    since = "0.13.0",
    batchOp = Some("tab-color"),
    cliVerb = Some("tab-color")
  )

  private val setAutoFilter = spec(
    "set-auto-filter",
    Vector(
      opt("range", Range, "The filter range; may be sheet-qualified"),
      opt("clear", Bool, "Strip the sheet's autoFilter, even one preserved from the source file")
    ),
    "Set or clear the sheet-level autoFilter.",
    since = "0.18.0",
    batchOp = Some("autofilter"),
    cliVerb = Some("autofilter")
  )

  private val setPageSetup = spec(
    "set-page-setup",
    Vector(
      opt("orientation", Enum(orientationValues), "Page orientation"),
      opt("scale", Int, "Print scale, 10-400"),
      opt("fitToWidth", Int, "Pages wide; 0 = automatic"),
      opt("fitToHeight", Int, "Pages tall; 0 = automatic (GH-463)"),
      opt("fitToPage", Bool, "Force the fitToPage flag on or off; omit to derive it")
    ),
    "Print setup: orientation, scale, fit-to pages.",
    since = "0.13.0",
    batchOp = Some("page-setup"),
    cliVerb = Some("page-setup")
  )

  private val setHeaderFooter = spec(
    "set-header-footer",
    Vector(
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
    "Print headers and footers.",
    since = "0.13.0",
    batchOp = Some("header-footer"),
    cliVerb = Some("header-footer")
  )

  private def structuralRows(name: String, doc: String): EditSpec =
    spec(
      name,
      Vector(req("at", RowKind, "1-based row"), opt("count", Int, "How many (default 1)")),
      doc,
      since = "0.19.0",
      cliVerb = Some(name),
      cellMutating = true,
      structural = true,
      needsFormula = true,
      idempotent = false
    )

  private def structuralCols(name: String, doc: String): EditSpec =
    spec(
      name,
      Vector(req("at", ColumnKind, "Column letter"), opt("count", Int, "How many (default 1)")),
      doc,
      since = "0.19.0",
      cliVerb = Some(name),
      cellMutating = true,
      structural = true,
      needsFormula = true,
      idempotent = false
    )

  private val insertRows =
    structuralRows("insert-rows", "Insert rows before `at`, rewriting references on every sheet.")
  private val deleteRows =
    structuralRows("delete-rows", "Delete rows from `at`, rewriting references on every sheet.")
  private val insertCols =
    structuralCols("insert-cols", "Insert columns before `at`, rewriting references everywhere.")
  private val deleteCols =
    structuralCols("delete-cols", "Delete columns from `at`, rewriting references everywhere.")

  private val addSheet = spec(
    "add-sheet",
    Vector(
      req("name", SheetKind, "Name of the new sheet"),
      opt("after", SheetKind, "Insert after this sheet (default: at the end)"),
      opt("before", SheetKind, "Insert before this sheet")
    ),
    "Add an empty sheet.",
    since = "0.9.6",
    batchOp = Some("add-sheet"),
    cliVerb = Some("add-sheet"),
    sheetScoped = false,
    structural = true,
    idempotent = false
  )

  private val removeSheet = spec(
    "remove-sheet",
    Vector(req("name", SheetKind, "The sheet to remove")),
    "Remove a sheet (a workbook keeps at least one).",
    since = "0.9.6",
    cliVerb = Some("remove-sheet"),
    sheetScoped = false,
    structural = true,
    idempotent = false
  )

  private val renameSheet = spec(
    "rename-sheet",
    Vector(req("from", SheetKind, "Current sheet name"), req("to", SheetKind, "New sheet name")),
    "Rename a sheet, rewriting every formula, defined name and rule that refers to it; a rename " +
      "of the scope's default sheet retargets the edits after it.",
    since = "0.9.6",
    batchOp = Some("rename-sheet"),
    cliVerb = Some("rename-sheet"),
    sheetScoped = false,
    structural = true,
    needsFormula = true,
    idempotent = false
  )

  private val moveSheet = spec(
    "move-sheet",
    Vector(
      req("name", SheetKind, "The sheet to move"),
      opt("to", Int, "Target position (0-based)"),
      opt("after", SheetKind, "Move after this sheet"),
      opt("before", SheetKind, "Move before this sheet")
    ),
    "Move a sheet to a position, or after/before another sheet (exactly one).",
    since = "0.9.6",
    cliVerb = Some("move-sheet"),
    oneOf = Vector(Set("to"), Set("after"), Set("before")),
    sheetScoped = false,
    structural = true
  )

  private val copySheet = spec(
    "copy-sheet",
    Vector(
      req("source", SheetKind, "The sheet to copy"),
      req("target", SheetKind, "Name of the copy")
    ),
    "Copy a sheet under a new name.",
    since = "0.9.6",
    cliVerb = Some("copy-sheet"),
    sheetScoped = false,
    structural = true,
    idempotent = false
  )

  private val hideSheet = spec(
    "hide-sheet",
    Vector(
      req("name", SheetKind, "The sheet to hide"),
      opt("veryHidden", Bool, "Very hidden: not reachable from Excel's UI, only VBA")
    ),
    "Hide a sheet from the tab bar (a workbook keeps at least one visible sheet).",
    since = "0.9.6",
    cliVerb = Some("sheets hide"),
    sheetScoped = false
  )

  private val showSheet = spec(
    "show-sheet",
    Vector(req("name", SheetKind, "The sheet to show")),
    "Make a hidden sheet visible.",
    since = "0.9.6",
    cliVerb = Some("sheets show"),
    sheetScoped = false
  )

  private val defineName = spec(
    "define-name",
    Vector(
      req("name", Str, "The defined name"),
      req("refersTo", Str, "The reference or formula it points to, e.g. Sheet1!$A$1:$A$10"),
      opt("scope", SheetKind, "Scope the name to this sheet (default: workbook)")
    ),
    "Add or replace a defined name (named range).",
    since = "0.10.0",
    cliVerb = Some("name add"),
    sheetScoped = false,
    structural = true
  )

  private val removeName = spec(
    "remove-name",
    Vector(
      req("name", Str, "The defined name"),
      opt("scope", SheetKind, "The sheet the name is scoped to (default: workbook)")
    ),
    "Remove a defined name.",
    since = "0.10.0",
    cliVerb = Some("name remove"),
    sheetScoped = false,
    structural = true,
    idempotent = false
  )

  /** One row per [[Edit]] case, in declaration order. */
  val all: Vector[EditSpec] = Vector(
    put,
    putValues,
    putFormula,
    putFormulas,
    dragFormula,
    fill,
    copy,
    sort,
    clear,
    style,
    merge,
    unmerge,
    colWidth,
    rowHeight,
    hideCols,
    showCols,
    hideRows,
    showRows,
    groupRows,
    groupCols,
    ungroupRows,
    ungroupCols,
    autoFit,
    setComment,
    removeComment,
    hyperlink,
    addConditionalFormat,
    addChart,
    addImage,
    freeze,
    unfreeze,
    setSheetView,
    setTabColor,
    setAutoFilter,
    setPageSetup,
    setHeaderFooter,
    insertRows,
    deleteRows,
    insertCols,
    deleteCols,
    addSheet,
    removeSheet,
    renameSheet,
    moveSheet,
    copySheet,
    hideSheet,
    showSheet,
    defineName,
    removeName
  )

  private val byKey: Map[String, EditSpec] =
    // Case names first so `find("put")` is the Put row; a batch op name then resolves to the FIRST
    // case that is a form of it (`putf` → put-formula), aliases last.
    val names = all.map(s => normalize(s.name) -> s)
    val batch = all.flatMap(s => s.batchOp.map(normalize(_) -> s))
    val aliases = all.flatMap(s => s.aliases.map(normalize(_) -> s))
    (names ++ batch ++ aliases).foldLeft(Map.empty[String, EditSpec]) { case (m, (k, v)) =>
      if m.contains(k) then m else m.updated(k, v)
    }

  /** Resolve a case name, a batch op name or an alias in any case/separator spelling. */
  def find(name: String): Option[EditSpec] = byKey.get(normalize(name))

  /** The row of an edit. Total and exhaustive: a new case must be placed in the table. */
  def specOf(edit: Edit): EditSpec = edit match
    case _: Edit.Put => put
    case _: Edit.PutValues => putValues
    case _: Edit.PutFormula => putFormula
    case _: Edit.PutFormulas => putFormulas
    case _: Edit.DragFormula => dragFormula
    case _: Edit.Fill => fill
    case _: Edit.Copy => copy
    case _: Edit.Sort => sort
    case _: Edit.Clear => clear
    case _: Edit.Style => style
    case _: Edit.Merge => merge
    case _: Edit.Unmerge => unmerge
    case _: Edit.ColWidth => colWidth
    case _: Edit.RowHeight => rowHeight
    case _: Edit.HideCols => hideCols
    case _: Edit.ShowCols => showCols
    case _: Edit.HideRows => hideRows
    case _: Edit.ShowRows => showRows
    case _: Edit.GroupRows => groupRows
    case _: Edit.GroupCols => groupCols
    case _: Edit.UngroupRows => ungroupRows
    case _: Edit.UngroupCols => ungroupCols
    case _: Edit.AutoFit => autoFit
    case _: Edit.SetComment => setComment
    case _: Edit.RemoveComment => removeComment
    case _: Edit.Hyperlink => hyperlink
    case _: Edit.AddConditionalFormat => addConditionalFormat
    case _: Edit.AddChart => addChart
    case _: Edit.AddImage => addImage
    case _: Edit.Freeze => freeze
    case _: Edit.Unfreeze => unfreeze
    case _: Edit.SetSheetView => setSheetView
    case _: Edit.SetTabColor => setTabColor
    case _: Edit.SetAutoFilter => setAutoFilter
    case _: Edit.SetPageSetup => setPageSetup
    case _: Edit.SetHeaderFooter => setHeaderFooter
    case _: Edit.InsertRows => insertRows
    case _: Edit.DeleteRows => deleteRows
    case _: Edit.InsertCols => insertCols
    case _: Edit.DeleteCols => deleteCols
    case _: Edit.AddSheet => addSheet
    case _: Edit.RemoveSheet => removeSheet
    case _: Edit.RenameSheet => renameSheet
    case _: Edit.MoveSheet => moveSheet
    case _: Edit.CopySheet => copySheet
    case _: Edit.HideSheet => hideSheet
    case _: Edit.ShowSheet => showSheet
    case _: Edit.DefineName => defineName
    case _: Edit.RemoveName => removeName

  /** The kebab name of an edit's case (`DragFormula` → `drag-formula`). */
  def nameOf(edit: Edit): String = specOf(edit).name
