<!-- GENERATED from `xl batch --schema` by DocsGenSpec; do not edit by hand.
     A diff here is a contract change. Regenerate: XL_UPDATE_DOCS=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.DocsGenSpec -->

# xl batch operations

A JSON array of operations `xl batch` applies in order. Every sheet-scoped op accepts an optional "sheet" key; property names are accepted in camelCase or kebab-case.

`xl batch --schema` prints this table as a JSON Schema (draft 2020-12): an array whose
items are exactly one of the per-op object schemas below, discriminated by `op`. Every
property is accepted in camelCase or kebab-case; the aliases column lists the other names
a property answers to. An op the streaming writer cannot apply is refused by index under
`--stream` before any byte is written (`UNSUPPORTED_IN_STREAM`, exit 2). `put`/`putf`
`format` is an explicit hint and replaces the cell's number format; a detected format
(currency, percent, ISO date) only applies to a General cell.

## Operations

| op | aliases | stream | mutates cells | sheet key | CLI twin | since | summary |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `put` | — | yes | yes | yes | `put` | 0.1.0 | Write a value (or a row-major array of values) with optional number format. |
| `putf` | — | yes | yes | yes | `putf` | 0.1.0 | Write a formula: one cell, dragged across a range from an anchor, or explicit per cell. |
| `style` | — | yes | no | yes | `style` | 0.1.0 | Style a range: font, fill, alignment, number format, borders (merged unless `replace`). |
| `merge` | — | yes | no | yes | `merge` | 0.1.0 | Merge a range into one cell. |
| `unmerge` | — | yes | no | yes | `unmerge` | 0.1.0 | Unmerge a merged range. |
| `colwidth` | — | yes | no | yes | `col` | 0.1.0 | Set a column's width. |
| `rowheight` | — | yes | no | yes | `row` | 0.1.0 | Set a row's height. |
| `comment` | — | no | no | yes | `comment` | 0.9.6 | Add or replace a cell comment. |
| `remove-comment` | — | no | no | yes | `remove-comment` | 0.9.6 | Remove a cell's comment. |
| `hyperlink` | — | no | no | yes | — | 0.10.0 | Set or clear a cell hyperlink. |
| `clear` | — | no | yes | yes | `clear` | 0.9.6 | Clear a range: contents by default (unmerging overlaps), or styles/comments/all. |
| `col-hide` | — | yes | no | yes | `col` | 0.9.6 | Hide a column. |
| `col-show` | — | yes | no | yes | `col` | 0.9.6 | Show a hidden column. |
| `row-hide` | — | yes | no | yes | `row` | 0.9.6 | Hide a row. |
| `row-show` | — | yes | no | yes | `row` | 0.9.6 | Show a hidden row. |
| `group-rows` | — | no | no | yes | `group-rows` | 0.18.0 | Group rows into an outline. |
| `group-cols` | — | no | no | yes | `group-cols` | 0.18.0 | Group columns into an outline. |
| `ungroup-rows` | — | no | no | yes | `ungroup-rows` | 0.18.0 | Clear the outline level and collapse markers of the rows. |
| `ungroup-cols` | — | no | no | yes | `ungroup-cols` | 0.18.0 | Clear the outline level and collapse markers of the cols. |
| `autofit` | — | no | no | yes | `autofit` | 0.9.6 | Auto-fit column widths to their content. |
| `add-sheet` | — | no | no | no | `add-sheet` | 0.9.6 | Add an empty sheet. |
| `rename-sheet` | — | no | no | no | `rename-sheet` | 0.9.6 | Rename a sheet, rewriting every formula, defined name and rule that refers to it; a rename of the batch's default sheet retargets the ops after it. |
| `freeze` | — | no | no | yes | `freeze` | 0.10.0 | Freeze panes above and left of the cell. |
| `unfreeze` | — | no | no | yes | `unfreeze` | 0.10.0 | Remove the sheet's freeze panes. |
| `copy` | — | no | yes | yes | `copy` | 0.10.0 | Copy a range, shifting relative references like Excel; each side may name its own sheet. |
| `chart` | — | no | no | yes | `chart add` | 0.15.0 | Add a chart, mirroring `chart add`. |
| `sheet-view` | — | no | no | yes | `sheet-view` | 0.13.0 | Sheet view settings: gridlines, zoom, selected tab. |
| `tab-color` | — | no | no | yes | `tab-color` | 0.13.0 | Set or clear the sheet tab color. |
| `autofilter` | — | no | no | yes | `autofilter` | 0.18.0 | Set or clear the sheet-level autoFilter. |
| `page-setup` | — | no | no | yes | `page-setup` | 0.13.0 | Print setup: orientation, scale, fit-to pages. |
| `header-footer` | — | no | no | yes | `header-footer` | 0.13.0 | Print headers and footers. |
| `cf` | — | no | no | yes | `cf add` | 0.13.0 | Add a conditional-formatting rule; priorities are assigned in order. |

## Properties

### `put`

Write a value (or a row-major array of values) with optional number format.

```json
{"op":"put","ref":"A1","value":1234.5,"format":"currency"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "put" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `ref` | string | yes | — | Target cell, or a range when `values` is given |
| `value` | string \| number \| boolean \| null | no | — | The value: JSON number/boolean/null as-is; strings are smart-detected (currency, percent, ISO date) unless `detect` is false |
| `values` | array of string \| number \| boolean \| null | no | — | Row-major values for the `ref` range, one per cell |
| `format` | string | no | `numFormat`, `num-format` | Number format: a name (general, integer, decimal, currency, percent, date, datetime, time, text) or an Excel format code. Explicit: it REPLACES the cell's number format (GH-560). |
| `detect` | boolean | no | — | Smart string detection (default true); false stores strings as text |

Exactly one of: {value} \| {values}.

### `putf`

Write a formula: one cell, dragged across a range from an anchor, or explicit per cell.

```json
{"op":"putf","ref":"B2:B10","value":"=A2*2","from":"B2","format":"#,##0.0"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "putf" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `ref` | string | yes | — | Target cell, or a range when dragging (`from`) or listing `values` |
| `value` | string | no | `formula` | The formula (leading = optional) |
| `values` | array of string | no | — | Row-major formulas for the `ref` range, one per cell, as-is |
| `from` | string | no | `anchor` | Anchor cell: the formula is dragged across `ref` from here, shifting relative refs like Excel fill-down |
| `format` | string | no | `numFormat`, `num-format` | Number format: a name (general, integer, decimal, currency, percent, date, datetime, time, text) or an Excel format code. Explicit: it REPLACES the cell's number format (GH-560). |

Exactly one of: {value} \| {values}.

### `style`

Style a range: font, fill, alignment, number format, borders (merged unless `replace`).

```json
{"op":"style","range":"A1:D1","bold":true,"bg":"#4472C4","fg":"#FFFFFF","align":"center"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "style" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `range` | string | yes | — | Cells to style (a single cell is a 1x1 range) |
| `bold` | boolean | no | — | Bold font |
| `italic` | boolean | no | — | Italic font |
| `underline` | boolean | no | — | Underlined font |
| `bg` | string | no | — | Fill color: named, #hex, rgb(r,g,b) or theme:accent1[:tint] |
| `fg` | string | no | — | Font color |
| `fontSize` | number | no | `font-size` | Font size in points |
| `fontName` | string | no | `font-name` | Font name |
| `align` | string one of left, center, right | no | `halign` | Horizontal alignment |
| `valign` | string one of top, middle, center, bottom | no | — | Vertical alignment |
| `wrap` | boolean | no | — | Wrap text |
| `numFormat` | string | no | `num-format`, `format` | Number format name or Excel format code (applied as given, a suspect string warns) |
| `border` | string one of none, thin, medium, thick, dashed, dotted, double | no | — | All-sides border style |
| `borderTop` | string one of none, thin, medium, thick, dashed, dotted, double | no | `border-top` | Top border style |
| `borderRight` | string one of none, thin, medium, thick, dashed, dotted, double | no | `border-right` | Right border style |
| `borderBottom` | string one of none, thin, medium, thick, dashed, dotted, double | no | `border-bottom` | Bottom border style |
| `borderLeft` | string one of none, thin, medium, thick, dashed, dotted, double | no | `border-left` | Left border style |
| `borderColor` | string | no | `border-color` | Color of every border named in the op |
| `replace` | boolean | no | — | Replace the whole cell style instead of merging into it |

### `merge`

Merge a range into one cell.

```json
{"op":"merge","range":"A1:D1"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "merge" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `range` | string | yes | — | The range |

### `unmerge`

Unmerge a merged range.

```json
{"op":"unmerge","range":"A1:D1"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "unmerge" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `range` | string | yes | — | The range |

### `colwidth`

Set a column's width.

```json
{"op":"colwidth","col":"A","width":15.5}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "colwidth" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `col` | string | yes | — | Column letter |
| `width` | number | yes | — | Width in character units |

### `rowheight`

Set a row's height.

```json
{"op":"rowheight","row":1,"height":30}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "rowheight" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `row` | integer | yes | — | 1-based row number |
| `height` | number | yes | — | Height in points |

### `comment`

Add or replace a cell comment.

```json
{"op":"comment","ref":"A1","text":"Note","author":"User"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "comment" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `ref` | string | yes | — | The cell |
| `text` | string | yes | — | Comment text |
| `author` | string | no | — | Comment author |

### `remove-comment`

Remove a cell's comment.

```json
{"op":"remove-comment","ref":"A1"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "remove-comment" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `ref` | string | yes | — | The cell |

### `hyperlink`

Set or clear a cell hyperlink.

```json
{"op":"hyperlink","ref":"A1","target":"https://example.com"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "hyperlink" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `ref` | string | yes | — | The cell |
| `target` | string | no | `url` | URL or internal location; omit to clear the hyperlink |

### `clear`

Clear a range: contents by default (unmerging overlaps), or styles/comments/all.

```json
{"op":"clear","range":"A1:B10","all":true}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "clear" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `range` | string | yes | — | The range |
| `all` | boolean | no | — | Clear contents, styles and comments |
| `styles` | boolean | no | — | Clear styles only |
| `comments` | boolean | no | — | Clear comments only |

### `col-hide`

Hide a column.

```json
{"op":"col-hide","col":"C"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "col-hide" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `col` | string | yes | — | Column letter |

### `col-show`

Show a hidden column.

```json
{"op":"col-show","col":"C"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "col-show" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `col` | string | yes | — | Column letter |

### `row-hide`

Hide a row.

```json
{"op":"row-hide","row":5}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "row-hide" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `row` | integer | yes | — | 1-based row number |

### `row-show`

Show a hidden row.

```json
{"op":"row-show","row":5}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "row-show" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `row` | integer | yes | — | 1-based row number |

### `group-rows`

Group rows into an outline.

```json
{"op":"group-rows","rows":"10:20","level":1,"collapsed":false}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "group-rows" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `rows` | string | yes | — | Row span, e.g. "10:20" |
| `level` | integer | no | — | Outline level (default 1) |
| `collapsed` | boolean | no | — | Hide the members and mark the summary row/column after the group |

### `group-cols`

Group columns into an outline.

```json
{"op":"group-cols","cols":"E:H","level":1,"collapsed":false}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "group-cols" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `cols` | string | yes | — | Column span, e.g. "E:H" |
| `level` | integer | no | — | Outline level (default 1) |
| `collapsed` | boolean | no | — | Hide the members and mark the summary row/column after the group |

### `ungroup-rows`

Clear the outline level and collapse markers of the rows.

```json
{"op":"ungroup-rows","rows":"10:20"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "ungroup-rows" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `rows` | string | yes | — | Row span, e.g. "10:20" |

### `ungroup-cols`

Clear the outline level and collapse markers of the cols.

```json
{"op":"ungroup-cols","cols":"E:H"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "ungroup-cols" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `cols` | string | yes | — | Column span, e.g. "E:H" |

### `autofit`

Auto-fit column widths to their content.

```json
{"op":"autofit","columns":"A:F"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "autofit" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `columns` | string | no | — | A column ("A") or span ("A:F"); omit for every used column |

### `add-sheet`

Add an empty sheet.

```json
{"op":"add-sheet","name":"Summary","after":"Sheet1"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "add-sheet" | yes | — | — |
| `name` | string | yes | — | Name of the new sheet |
| `after` | string | no | — | Insert after this sheet (default: at the end) |

### `rename-sheet`

Rename a sheet, rewriting every formula, defined name and rule that refers to it; a rename of the batch's default sheet retargets the ops after it.

```json
{"op":"rename-sheet","from":"Old","to":"New"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "rename-sheet" | yes | — | — |
| `from` | string | yes | — | Current sheet name |
| `to` | string | yes | — | New sheet name |

### `freeze`

Freeze panes above and left of the cell.

```json
{"op":"freeze","ref":"B2"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "freeze" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `ref` | string | yes | — | The cell |

### `unfreeze`

Remove the sheet's freeze panes.

```json
{"op":"unfreeze"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "unfreeze" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |

### `copy`

Copy a range, shifting relative references like Excel; each side may name its own sheet.

```json
{"op":"copy","source":"A1:B2","target":"D1","valuesOnly":false}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "copy" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `source` | string | yes | — | Source cell or range; may be sheet-qualified |
| `target` | string | yes | — | Target cell (expanded to the source's size) or range; may be sheet-qualified |
| `valuesOnly` | boolean | no | `values-only` | Paste cached values instead of formulas |

### `chart`

Add a chart, mirroring `chart add`.

```json
{"op":"chart","type":"column","data":"B2:C4","categories":"A2:A4","seriesNames":"North,South","title":"Sales","legend":"right","at":"E2:K15"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "chart" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `type` | string one of column, bar, line, pie | yes | — | Chart type |
| `grouping` | string one of clustered, stacked, percent-stacked | no | — | Bar/column grouping |
| `data` | string | yes | — | Values range; may be sheet-qualified |
| `categories` | string | no | — | Categories range; may be sheet-qualified |
| `seriesNames` | string | no | `series-names` | Comma-separated series names, positional |
| `seriesColors` | string | no | `series-colors` | Comma-separated series colors, positional |
| `title` | string | no | — | Chart title |
| `legend` | string one of right, left, top, bottom, top-right, none | no | — | Legend position |
| `at` | string | yes | — | Placement: an anchor cell or the range the chart covers |

### `sheet-view`

Sheet view settings: gridlines, zoom, selected tab.

```json
{"op":"sheet-view","gridlines":false,"zoom":85}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "sheet-view" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `gridlines` | boolean | no | — | Show gridlines |
| `zoom` | integer | no | — | Zoom percentage |
| `tabSelected` | boolean | no | `tab-selected` | Select this sheet's tab |

### `tab-color`

Set or clear the sheet tab color.

```json
{"op":"tab-color","color":"#1F4E79"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "tab-color" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `color` | string | no | — | Tab color: named, #hex, rgb(r,g,b) or theme:accent1[:tint] |
| `clear` | boolean | no | — | Remove the tab color |

### `autofilter`

Set or clear the sheet-level autoFilter.

```json
{"op":"autofilter","range":"A1:M29"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "autofilter" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `range` | string | no | — | The filter range; may be sheet-qualified |
| `clear` | boolean | no | — | Strip the sheet's autoFilter, even one preserved from the source file |

### `page-setup`

Print setup: orientation, scale, fit-to pages.

```json
{"op":"page-setup","orientation":"landscape","fitToWidth":1,"fitToHeight":0}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "page-setup" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `orientation` | string one of portrait, landscape | no | — | Page orientation |
| `scale` | integer | no | — | Print scale, 10-400 |
| `fitToWidth` | integer | no | `fit-to-width` | Pages wide; 0 = automatic |
| `fitToHeight` | integer | no | `fit-to-height` | Pages tall; 0 = automatic (GH-463) |
| `fitToPage` | boolean | no | `fit-to-page` | Force the fitToPage flag on or off; omit to derive it |

### `header-footer`

Print headers and footers.

```json
{"op":"header-footer","oddFooter":"&LConfidential&RPage &P of &N"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "header-footer" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `oddHeader` | string | no | `odd-header` | Header (&L/&C/&R sections; &P page, &N total, &D date, &F file, &A sheet) |
| `oddFooter` | string | no | `odd-footer` | Footer |
| `evenHeader` | string | no | `even-header` | Even-page header (sets differentOddEven) |
| `evenFooter` | string | no | `even-footer` | Even-page footer (sets differentOddEven) |
| `firstHeader` | string | no | `first-header` | First-page header (sets differentFirst) |
| `firstFooter` | string | no | `first-footer` | First-page footer (sets differentFirst) |
| `differentOddEven` | boolean | no | `different-odd-even` | Force the odd/even flag on |
| `differentFirst` | boolean | no | `different-first` | Force the first-page flag on |

### `cf`

Add a conditional-formatting rule; priorities are assigned in order.

```json
{"op":"cf","range":"A1:A10","rule":"cellIs:greaterThan:100","bold":true,"bg":"#FFC7CE","fg":"#9C0006"}
```

| property | type | required | aliases | description |
| --- | --- | --- | --- | --- |
| `op` | const "cf" | yes | — | — |
| `sheet` | string | no | — | Sheet for the op's unqualified refs: a sheet-qualified ref wins, then this key, then -s/--sheet |
| `range` | string | yes | — | The range the rule applies to |
| `rule` | string | yes | — | Rule DSL as `cf add`: cellIs:greaterThan:100, between:1:10, expression:…, colorScale:…, dataBar:…, top10:…, text:… |
| `bold` | boolean | no | — | Differential style: bold |
| `italic` | boolean | no | — | Differential style: italic |
| `underline` | boolean | no | — | Differential style: underline |
| `strike` | boolean | no | — | Differential style: strikethrough |
| `bg` | string | no | — | Differential style: fill color |
| `fg` | string | no | — | Differential style: font color |
