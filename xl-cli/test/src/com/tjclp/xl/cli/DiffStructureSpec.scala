package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import scala.util.Random

import cats.effect.{ExitCode, IO}
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.addressing.{CellRange, Column, Row, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cf.{CfOperator, CfRule, ConditionalFormat}
import com.tjclp.xl.cli.commands.DiffCommands
import com.tjclp.xl.cli.commands.DiffCommands.{
  AxisChange,
  NameChange,
  NameEntry,
  NameSnapshot,
  PropValue,
  Property,
  PropertyChange,
  SheetOrder,
  WorkbookDiff
}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.{ColumnProperties, DataValidation, FreezePane, RowProperties}
import com.tjclp.xl.styles.{CellStyle, Dxf}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.workbooks.DefinedName

/**
 * Sheet-structure comparison in `diff` (Weaver: "a row-height-only change reports Files are
 * identical"): rows and columns, sheet properties, conditional formats, data validations, sheet
 * order and defined names, their equivalence rules, the `-s` / `--formulas-only` / `--cells-only`
 * interplay, the JSON and markdown shapes, and the exit code through a real write and read.
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class DiffStructureSpec extends CatsEffectSuite:

  private def wb(sheets: Sheet*): Workbook = Workbook(sheets.toVector)

  private def diffOf(
    a: Workbook,
    b: Workbook,
    filter: Option[String] = None,
    formulasOnly: Boolean = false,
    cellsOnly: Boolean = false
  ): WorkbookDiff =
    DiffCommands
      .computeDiff(a, b, filter, formulasOnly = formulasOnly, cellsOnly = cellsOnly)
      .fold(err => fail(err), identity)

  private def md(diff: WorkbookDiff): String = DiffCommands.renderMarkdown(diff, "a.xlsx", "b.xlsx")

  private def json(diff: WorkbookDiff): ujson.Value = ujson.read(DiffCommands.renderJson(diff))

  private def size(d: String): PropValue = PropValue.Size(BigDecimal(d))

  private def rng(a1: String): CellRange = CellRange.parse(a1).toOption.get

  private def row(n: Int): Row = Row.from1(n)

  private def col(letter: String): Column = Column.fromLetter(letter).toOption.get

  private def withRows(sheet: Sheet, rows: (Int, RowProperties)*): Sheet =
    rows.foldLeft(sheet) { case (s, (n, p)) => s.setRowProperties(row(n), p) }

  private def height(d: Double): RowProperties = RowProperties(height = Some(d))

  private def rowsOf(diff: WorkbookDiff): Vector[AxisChange] = diff.sheets.flatMap(_.rows)

  private def columnsOf(diff: WorkbookDiff): Vector[AxisChange] = diff.sheets.flatMap(_.columns)

  private def propertiesOf(diff: WorkbookDiff): Vector[PropertyChange] =
    diff.sheets.flatMap(_.properties)

  private val base = Sheet("S").put(ref"A1", CellValue.Number(1))

  // ========== Weaver's case ==========

  test("a row-height-only change is not identical: 5:5 height 15 -> 30") {
    val b = withRows(base, 5 -> height(30))
    val diff = diffOf(wb(base), wb(b))
    assert(!diff.identical)
    assertEquals(
      rowsOf(diff),
      Vector(
        AxisChange("5:5", Vector(PropertyChange(Property.Height, size("15"), size("30"))), false)
      )
    )
    val text = md(diff)
    assert(text.contains("\nRows (1):\n  5:5: height 15 -> 30\n"), text)
    assert(!text.contains("Files are identical"), text)
    val r = json(diff)("sheets")(0)("rows")(0)
    assertEquals(r("ref").str, "5:5")
    assertEquals(r("changes")(0)("property").str, "height")
    assertEquals(r("changes")(0)("before").num, 15.0)
    assertEquals(r("changes")(0)("after").num, 30.0)
    assertEquals(r("styleChanged").bool, false)
  }

  // ========== Sizes ==========

  test("sizes compare effective values at 2 decimals; a non-finite size counts as unspecified") {
    // explicit 15 vs absent, no defaults: the stock 15pt row
    assert(diffOf(wb(withRows(base, 5 -> height(15))), wb(base)).identical)
    // explicit 20 vs absent under a shared default of 20
    val d20 = base.copy(defaultRowHeight = Some(20.0))
    assert(diffOf(wb(withRows(d20, 5 -> height(20))), wb(d20)).identical)
    assert(
      diffOf(wb(withRows(base, 5 -> height(15))), wb(withRows(base, 5 -> height(15.004)))).identical
    )
    val differs =
      diffOf(wb(withRows(base, 5 -> height(15))), wb(withRows(base, 5 -> height(15.01))))
    assertEquals(
      rowsOf(differs),
      Vector(
        AxisChange("5:5", Vector(PropertyChange(Property.Height, size("15"), size("15.01"))), false)
      )
    )
    assert(diffOf(wb(withRows(base, 5 -> height(Double.NaN))), wb(base)).identical)
    assert(diffOf(wb(withRows(base, 5 -> height(Double.PositiveInfinity))), wb(base)).identical)
  }

  test("a sheet-default change is reported once, as a sheet property, never per row") {
    val d15 = base.copy(defaultRowHeight = Some(15.0))
    val d20 = base.copy(defaultRowHeight = Some(20.0))
    val plain = diffOf(wb(d15), wb(d20))
    assertEquals(rowsOf(plain), Vector.empty)
    assertEquals(
      propertiesOf(plain),
      Vector(PropertyChange(Property.DefaultRowHeight, size("15"), size("20")))
    )
    // row 5 explicit 20 against a default of 20 with no row 5 entry: same rendered height
    val same = diffOf(wb(withRows(d15, 5 -> height(20))), wb(d20))
    assertEquals(rowsOf(same), Vector.empty)
    assertEquals(propertiesOf(same).map(_.property), Vector(Property.DefaultRowHeight))
    // row 5 explicit 15 against a default of 20: it really renders taller
    val taller = diffOf(wb(withRows(base, 5 -> height(15))), wb(d20))
    assertEquals(
      rowsOf(taller),
      Vector(
        AxisChange("5:5", Vector(PropertyChange(Property.Height, size("15"), size("20"))), false)
      )
    )
    val text = md(taller)
    assert(text.contains("  5:5: height 15 -> 20\n"), text)
    assert(text.contains("\nSheet properties (1):\n  defaultRowHeight 15 -> 20\n"), text)
    // None vs an explicit stock default: an xl-written book against its Excel resave
    assert(diffOf(wb(base), wb(d15)).identical)
  }

  test("hidden, outlineLevel (None = 0) and collapsed, in property order inside one entry") {
    val zero = withRows(base, 5 -> RowProperties(outlineLevel = Some(0)))
    assert(diffOf(wb(base), wb(zero)).identical)
    val grouped = withRows(
      base,
      5 -> RowProperties(
        height = Some(20),
        hidden = true,
        outlineLevel = Some(2),
        collapsed = true
      )
    )
    val diff = diffOf(wb(base), wb(grouped))
    assertEquals(
      rowsOf(diff),
      Vector(
        AxisChange(
          "5:5",
          Vector(
            PropertyChange(Property.Height, size("15"), size("20")),
            PropertyChange(Property.Hidden, PropValue.Flag(false), PropValue.Flag(true)),
            PropertyChange(Property.OutlineLevel, PropValue.Level(0), PropValue.Level(2)),
            PropertyChange(Property.Collapsed, PropValue.Flag(false), PropValue.Flag(true))
          ),
          false
        )
      )
    )
    assert(
      md(diff).contains(
        "  5:5: height 15 -> 20, hidden false -> true, outlineLevel 0 -> 2, collapsed false -> true\n"
      ),
      md(diff)
    )
  }

  test("a row default style compares RESOLVED formatting and reports styleChanged") {
    val bold = CellStyle.default.withFont(Font.default.withBold(true))
    val italic = CellStyle.default.withFont(Font.default.withItalic(true))
    def styled(s: Sheet, pre: Vector[CellStyle], style: CellStyle): Sheet =
      val registry = pre.foldLeft(s.styleRegistry)((r, st) => r.register(st)._1)
      val (withStyle, id) = registry.register(style)
      s.copy(styleRegistry = withStyle).setRowProperties(row(3), RowProperties(styleId = Some(id)))
    // the same formatting under different ids
    val a = styled(base, Vector.empty, bold)
    val b = styled(base, Vector(italic), bold)
    assertNotEquals(a.rowProperties(row(3)).styleId, b.rowProperties(row(3)).styleId)
    assert(diffOf(wb(a), wb(b)).identical)
    // an id resolving to the default style is no style
    val defaultId = base.setRowProperties(
      row(3),
      RowProperties(styleId = Some(com.tjclp.xl.styles.units.StyleId(0)))
    )
    assert(diffOf(wb(base), wb(defaultId)).identical)
    // different formatting
    val c = styled(base, Vector.empty, italic)
    val diff = diffOf(wb(a), wb(c))
    assertEquals(rowsOf(diff), Vector(AxisChange("3:3", Vector.empty, true)))
    assert(md(diff).contains("\nRows (1):\n  3:3: [style]\n"), md(diff))
  }

  test("runs collapse consecutive rows with equal changes, with no gap bridging") {
    val tall = (5 to 200).map(_ -> height(30))
    val run = diffOf(wb(base), wb(withRows(base, tall*)))
    assertEquals(rowsOf(run).map(_.ref), Vector("5:200"))
    val gap = diffOf(wb(base), wb(withRows(base, 5 -> height(30), 7 -> height(30))))
    assertEquals(rowsOf(gap).map(_.ref), Vector("5:5", "7:7"))
    val adjacent = diffOf(wb(base), wb(withRows(base, 5 -> height(30), 6 -> height(31))))
    assertEquals(rowsOf(adjacent).map(_.ref), Vector("5:5", "6:6"))
  }

  test("columns: stock width, a width change, a whole K:XFD span, a default width change") {
    val stock = base.setColumnProperties(col("C"), ColumnProperties(width = Some(9.140625)))
    assert(diffOf(wb(stock), wb(base)).identical)
    val wide = base.setColumnProperties(col("C"), ColumnProperties(width = Some(14)))
    val widened = diffOf(wb(stock), wb(wide))
    assertEquals(
      columnsOf(widened),
      Vector(
        AxisChange("C:C", Vector(PropertyChange(Property.Width, size("9.14"), size("14"))), false)
      )
    )
    assert(md(widened).contains("\nColumns (1):\n  C:C: width 9.14 -> 14\n"), md(widened))
    val hiddenTail = (10 to Column.MaxIndex0).foldLeft(base) { (s, i) =>
      s.setColumnProperties(Column.from0(i), ColumnProperties(hidden = true))
    }
    assertEquals(hiddenTail.columnProperties.size, 16374)
    val tail = diffOf(wb(base), wb(hiddenTail))
    assertEquals(
      columnsOf(tail),
      Vector(
        AxisChange(
          "K:XFD",
          Vector(PropertyChange(Property.Hidden, PropValue.Flag(false), PropValue.Flag(true))),
          false
        )
      )
    )
    val defaultWidth = diffOf(wb(base), wb(base.copy(defaultColumnWidth = Some(12))))
    assertEquals(
      propertiesOf(defaultWidth),
      Vector(PropertyChange(Property.DefaultColumnWidth, size("9.14"), size("12")))
    )
    assert(md(defaultWidth).contains("  defaultColumnWidth 9.14 -> 12\n"), md(defaultWidth))
  }

  // ========== Sheet properties ==========

  test("freezePanes compares the anchor only: scroll state and Remove are not changes") {
    val frozen = base.freezeAt(ref"B2")
    val added = diffOf(wb(base), wb(frozen))
    assertEquals(
      propertiesOf(added),
      Vector(
        PropertyChange(
          Property.FreezePanes,
          PropValue.Anchor(None),
          PropValue.Anchor(Some(ref"B2"))
        )
      )
    )
    assert(md(added).contains("  freezePanes (none) -> B2\n"), md(added))
    val p = json(added)("sheets")(0)("properties")(0)
    assertEquals(p("property").str, "freezePanes")
    assertEquals(p("before"), ujson.Null)
    assertEquals(p("after").str, "B2")
    assert(diffOf(wb(base.freezeAt(ref"B2", ref"A40")), wb(frozen)).identical)
    assert(diffOf(wb(base.copy(freezePane = Some(FreezePane.Remove))), wb(base)).identical)
    val moved = diffOf(wb(frozen), wb(base.freezeAt(ref"C3")))
    assertEquals(
      propertiesOf(moved),
      Vector(
        PropertyChange(
          Property.FreezePanes,
          PropValue.Anchor(Some(ref"B2")),
          PropValue.Anchor(Some(ref"C3"))
        )
      )
    )
  }

  test("visibility: a sheet hidden between versions is reported; veryHidden by name") {
    val book = wb(base, Sheet("T"))
    def state(s: Option[String]): Workbook =
      book.setSheetState(SheetName.unsafe("T"), s).fold(e => fail(e.message), identity)
    val hidden = diffOf(book, state(Some("hidden")))
    assertEquals(hidden.sheets.map(_.name), Vector("T"))
    assertEquals(
      propertiesOf(hidden),
      Vector(
        PropertyChange(Property.Visibility, PropValue.State("visible"), PropValue.State("hidden"))
      )
    )
    assert(md(hidden).contains("  visibility visible -> hidden\n"), md(hidden))
    val very = diffOf(state(Some("hidden")), state(Some("veryHidden")))
    assertEquals(
      propertiesOf(very),
      Vector(
        PropertyChange(
          Property.Visibility,
          PropValue.State("hidden"),
          PropValue.State("veryHidden")
        )
      )
    )
  }

  // ========== Conditional formats ==========

  private val red = Dxf.fill(Color.Rgb(0xffff0000))
  private val green = Dxf.fill(Color.Rgb(0xff00ff00))

  private def cellIs(threshold: String, priority: Int, dxf: Dxf = red): CfRule =
    CfRule.CellIs(CfOperator.GreaterThan, threshold, None, Some(dxf), priority)

  private def cf(sheet: Sheet, blocks: ConditionalFormat*): Sheet =
    sheet.copy(conditionalFormats = blocks.toVector)

  private def block(ranges: String, rules: CfRule*): ConditionalFormat =
    ConditionalFormat.Rules(ranges.split(" ").toVector.map(rng), rules.toVector)

  test("conditional formats: added, removed and changed by sqref") {
    val a = cf(base, block("H2:H5", cellIs("100", 1)), block("A1:A9", cellIs("1", 2)))
    val b = cf(base, block("H2:H5", cellIs("200", 1)), block("C1", cellIs("1", 2)))
    val diff = diffOf(wb(a), wb(b))
    val sd = diff.sheets.head
    assertEquals(sd.conditionalFormatsAdded, Vector("C1"))
    assertEquals(sd.conditionalFormatsRemoved, Vector("A1:A9"))
    assertEquals(sd.conditionalFormatsChanged, Vector("H2:H5"))
    val text = md(diff)
    assert(text.contains("\nConditional formats added: C1\n"), text)
    assert(text.contains("\nConditional formats removed: A1:A9\n"), text)
    assert(text.contains("\nConditional formats changed: H2:H5\n"), text)
  }

  test("conditional formats: priorities, sqref token order and block order are not changes") {
    val a = cf(
      base,
      block("H2:H5 J2:J5", cellIs("100", 1), cellIs("50", 2, green)),
      block("A1:A9", cellIs("1", 3))
    )
    // renumbered priorities (same order within the block), reversed tokens, reversed blocks
    val b = cf(
      base,
      block("A1:A9", cellIs("1", 7)),
      block("J2:J5 H2:H5", cellIs("100", 4), cellIs("50", 5, green))
    )
    assert(diffOf(wb(a), wb(b)).identical)
    // the order of rules WITHIN a block is kept: swapping precedence is a change
    val swapped = cf(
      base,
      block("H2:H5 J2:J5", cellIs("50", 1, green), cellIs("100", 2)),
      block("A1:A9", cellIs("1", 3))
    )
    assertEquals(
      diffOf(wb(a), wb(swapped)).sheets.head.conditionalFormatsChanged,
      Vector("H2:H5 J2:J5")
    )
  }

  test("conditional formats: a Preserved rule compares its dxf payload, not priority or dxfId") {
    def preserved(priority: Int, dxfId: Int, payload: String): CfRule =
      CfRule.Preserved(
        s"""<cfRule type="iconSet" priority="$priority" dxfId="$dxfId" xr:uid="{$priority}"/>""",
        Some(priority),
        Some(payload)
      )
    val a = cf(base, block("B2:B9", preserved(1, 0, "<dxf><font><b/></font></dxf>")))
    val renumbered = cf(base, block("B2:B9", preserved(9, 4, "<dxf><font><b/></font></dxf>")))
    assert(diffOf(wb(a), wb(renumbered)).identical)
    val restyled = cf(base, block("B2:B9", preserved(1, 0, "<dxf><font><i/></font></dxf>")))
    assertEquals(diffOf(wb(a), wb(restyled)).sheets.head.conditionalFormatsChanged, Vector("B2:B9"))
  }

  // ========== Data validations ==========

  test("data validations: added, removed and changed by sqref; xr:uid is not content") {
    val a = base
      .withDataValidation(rng("B2:B9"), DataValidation.list("$Z$1:$Z$3"))
      .withDataValidation(rng("D2"), DataValidation.listOf("Yes", "No"))
    val b = base
      .withDataValidation(rng("B2:B9"), DataValidation.list("$Z$1:$Z$4"))
      .withDataValidation(rng("E2"), DataValidation.listOf("Yes", "No"))
    val sd = diffOf(wb(a), wb(b)).sheets.head
    assertEquals(sd.dataValidationsAdded, Vector("E2"))
    assertEquals(sd.dataValidationsRemoved, Vector("D2"))
    assertEquals(sd.dataValidationsChanged, Vector("B2:B9"))
    val text = md(diffOf(wb(a), wb(b)))
    assert(text.contains("\nData validations changed: B2:B9\n"), text)
    val decl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
    def preserved(uid: String): Sheet =
      base.copy(dataValidations =
        Vector(
          DataValidation.Preserved(
            decl + s"""<dataValidation type="whole" imeMode="on" sqref="Z9" xr:uid="{$uid}"/>"""
          )
        )
      )
    assert(diffOf(wb(preserved("A-1")), wb(preserved("B-2"))).identical)
  }

  test("data validations: a Preserved block's sqref token order is not content") {
    val decl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
    def preserved(sqref: String, uid: String, list: String): Sheet =
      base.copy(dataValidations =
        Vector(
          DataValidation.Preserved(
            decl + s"""<dataValidation type="list" allowBlank="1" showInputMessage="1" """ +
              s"""showErrorMessage="1" sqref="$sqref" xr:uid="{$uid}">""" +
              s"""<formula1>&quot;$list&quot;</formula1></dataValidation>"""
          )
        )
      )
    val a = preserved("B2:B9 D2", "A", "Low,High")
    assert(diffOf(wb(a), wb(preserved("D2 B2:B9", "C", "Low,High"))).identical)
    assertEquals(
      diffOf(
        wb(a),
        wb(preserved("D2 B2:B9", "C", "Low,Mid,High"))
      ).sheets.head.dataValidationsChanged,
      Vector("B2:B9 D2")
    )
    // an unparseable sqref keys by its raw text, so its token order still counts
    assertEquals(
      diffOf(
        wb(preserved("B2:B9 ??", "A", "Low")),
        wb(preserved("?? B2:B9", "A", "Low"))
      ).sheets.head.dataValidationsAdded,
      Vector("?? B2:B9")
    )
  }

  test("conditional formats: a Preserved block's sqref token order is not content") {
    val decl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
    def preserved(sqref: String, uid: String, dxfId: Int): Sheet =
      cf(
        base,
        ConditionalFormat.Preserved(
          decl + s"""<conditionalFormatting sqref="$sqref" x:unknown="1" xr:uid="{$uid}">""" +
            s"""<cfRule type="expression" dxfId="$dxfId" priority="$dxfId">""" +
            """<formula>H2&gt;0</formula></cfRule></conditionalFormatting>"""
        )
      )
    val a = preserved("H2:H5 J2:J5", "A", 0)
    assert(diffOf(wb(a), wb(preserved("J2:J5 H2:H5", "C", 0))).identical)
    // a block-level Preserved has no dxf payload field, so it keeps its dxfId
    assertEquals(
      diffOf(wb(a), wb(preserved("J2:J5 H2:H5", "C", 3))).sheets.head.conditionalFormatsChanged,
      Vector("H2:H5 J2:J5")
    )
  }

  // ========== Sheet order ==========

  test("sheet order: a reorder of shared sheets is reported; an insertion or -s is not") {
    val (sa, sb, sc) = (Sheet("A"), Sheet("B"), Sheet("C"))
    val reordered = diffOf(wb(sa, sb, sc), wb(sb, sa, sc))
    assert(!reordered.identical)
    assertEquals(
      reordered.sheetOrder,
      Some(SheetOrder(Vector("A", "B", "C"), Vector("B", "A", "C")))
    )
    assert(md(reordered).contains("\nSheet order: A, B, C -> B, A, C\n"), md(reordered))
    val inserted = diffOf(wb(sa, sb, sc), wb(sa, Sheet("X"), sb, sc))
    assertEquals(inserted.sheetOrder, None)
    assertEquals(inserted.sheetsAdded, Vector("X"))
    assert(diffOf(wb(sa, sb, sc), wb(sb, sa, sc), filter = Some("A")).identical)
  }

  // ========== Defined names ==========

  private def named(book: Workbook, names: DefinedName*): Workbook =
    book.copy(metadata = book.metadata.copy(definedNames = names.toVector))

  test("defined names: added, removed and changed (retarget, hidden) with case-insensitive keys") {
    val book = wb(Sheet("Model"), Sheet("Data"))
    val a = named(
      book,
      DefinedName("Rate", "Model!$B$1"),
      DefinedName("Old", "Model!$C$1"),
      DefinedName("Flag", "Model!$D$1")
    )
    val b = named(
      book,
      DefinedName("Rate", "=Model!$B$2"),
      DefinedName("New", "Model!$C$2"),
      DefinedName("Flag", "Model!$D$1", hidden = true)
    )
    val diff = diffOf(a, b)
    assertEquals(diff.namesAdded, Vector(NameEntry("New", None, NameSnapshot("Model!$C$2", false))))
    assertEquals(
      diff.namesRemoved,
      Vector(NameEntry("Old", None, NameSnapshot("Model!$C$1", false)))
    )
    assertEquals(
      diff.namesChanged,
      Vector(
        NameChange(
          "Flag",
          None,
          NameSnapshot("Model!$D$1", false),
          NameSnapshot("Model!$D$1", true)
        ),
        NameChange(
          "Rate",
          None,
          NameSnapshot("Model!$B$1", false),
          NameSnapshot("Model!$B$2", false)
        )
      )
    )
    val text = md(diff)
    assert(text.contains("\nNames added (1):\n  New = Model!$C$2\n"), text)
    assert(text.contains("\nNames removed (1):\n  Old = Model!$C$1\n"), text)
    assert(
      text.contains(
        "\nNames changed (2):\n  Flag: Model!$D$1 [hidden false -> true]\n" +
          "  Rate: Model!$B$1 -> Model!$B$2\n"
      ),
      text
    )
    // a case-only spelling difference is the same name
    assert(
      diffOf(named(book, DefinedName("rate", "1")), named(book, DefinedName("RATE", "1"))).identical
    )
  }

  test("defined names: _xlnm.* is skipped; scoped names follow their sheet; odd ids are labelled") {
    val book = wb(Sheet("Model"), Sheet("Data"))
    val builtIns = named(
      book,
      DefinedName("_xlnm._FilterDatabase", "Model!$A$1:$B$9", Some(0), hidden = true),
      DefinedName("_xlnm.Print_Titles", "Model!$1:$1", Some(0))
    )
    assert(diffOf(book, builtIns).identical)
    // the scoped name's localSheetId moves 0 -> 1 because its sheet moved: only the reorder shows
    val before = named(book, DefinedName("LocalRate", "Model!$B$1", Some(0)))
    val after =
      named(wb(Sheet("Data"), Sheet("Model")), DefinedName("LocalRate", "Model!$B$1", Some(1)))
    val moved = diffOf(before, after)
    assert(moved.sheetOrder.isDefined)
    assertEquals(moved.namesAdded ++ moved.namesRemoved, Vector.empty)
    assertEquals(moved.namesChanged, Vector.empty)
    val orphan = diffOf(book, named(book, DefinedName("Lost", "1", Some(7))))
    assertEquals(orphan.namesAdded.map(_.scope), Vector(Some("[localSheetId 7]")))
  }

  test("defined names: -s keeps only the filtered sheet's own names; labels quote the scope") {
    val book = wb(Sheet("Model"), Sheet("My Sheet"))
    val b = named(
      book,
      DefinedName("Global", "1"),
      DefinedName("LocalRate", "Model!$B$1", Some(0)),
      DefinedName("Rate", "2", Some(1))
    )
    val all = diffOf(book, b)
    assertEquals(
      all.namesAdded.map(e => (e.scope, e.name)),
      Vector((None, "Global"), (Some("Model"), "LocalRate"), (Some("My Sheet"), "Rate"))
    )
    val text = md(all)
    assert(text.contains("  Model!LocalRate = Model!$B$1\n"), text)
    assert(text.contains("  'My Sheet'!Rate = 2\n"), text)
    val filtered = diffOf(book, b, filter = Some("Model"))
    assertEquals(filtered.namesAdded.map(_.name), Vector("LocalRate"))
    val n = json(all)("namesAdded")
    assertEquals(n(0)("scope"), ujson.Null)
    assertEquals(n(1)("scope").str, "Model")
    assertEquals(n(1)("formula").str, "Model!$B$1")
    assertEquals(n(1)("hidden").bool, false)
  }

  // ========== Scope flags ==========

  /** Two books whose cells agree and whose structure differs in every compared category. */
  private def everyCategory: (Workbook, Workbook) =
    val a = Sheet("S")
      .put(ref"A1", CellValue.Number(1))
      .withDataValidation(rng("B2"), DataValidation.listOf("a", "b"))
    val b = withRows(Sheet("S").put(ref"A1", CellValue.Number(1)), 2 -> height(30))
      .setColumnProperties(col("B"), ColumnProperties(width = Some(20)))
      .freezeAt(ref"A2")
      .conditionalFormat(rng("A1:A9"), cellIs("1", 1))
    val bookA = named(wb(a, Sheet("T")), DefinedName("N", "1"))
    val bookB = named(wb(Sheet("T"), b), DefinedName("N", "2"))
    (bookA, bookB)

  test("--cells-only restores the cell-only scope; the JSON still carries every key") {
    val (a, b) = everyCategory
    val full = diffOf(a, b)
    assert(!full.identical)
    val sd = full.sheets.head
    assertEquals(sd.rows.length, 1)
    assertEquals(sd.columns.length, 1)
    assertEquals(sd.properties.map(_.property), Vector(Property.FreezePanes))
    assertEquals(sd.conditionalFormatsAdded, Vector("A1:A9"))
    assertEquals(sd.dataValidationsRemoved, Vector("B2"))
    assert(full.sheetOrder.isDefined)
    assertEquals(full.namesChanged.length, 1)
    assertEquals(full.structureChanges, 7)
    val cellsOnly = diffOf(a, b, cellsOnly = true)
    assert(cellsOnly.identical, s"$cellsOnly")
    val j = json(cellsOnly)
    assertEquals(j("sheetOrder"), ujson.Null)
    Vector("namesAdded", "namesRemoved", "namesChanged").foreach(k =>
      assertEquals(j(k).arr.length, 0)
    )
    assertEquals(md(cellsOnly).contains("Files are identical."), true)
  }

  test("--formulas-only still reports a row-height change") {
    val diff = diffOf(wb(base), wb(withRows(base, 1 -> height(30))), formulasOnly = true)
    assertEquals(rowsOf(diff).map(_.ref), Vector("1:1"))
  }

  // ========== Determinism and shapes ==========

  test("determinism: construction order never changes the rendered report") {
    val rnd = new Random(677)
    val heights = (1 to 200).map(i => (i * 2) -> height(16 + rnd.nextInt(40))).toVector
    val names = (1 to 20).map(i => DefinedName(s"N$i", s"$i")).toVector
    val blocks = Vector(block("A1:A9", cellIs("1", 1)), block("C1:C9", cellIs("2", 2)))
    def build(
      order: Vector[(Int, RowProperties)],
      ns: Vector[DefinedName],
      bs: Vector[ConditionalFormat]
    ) =
      named(wb(cf(withRows(base, order*), bs*)), ns*)
    val one = diffOf(wb(base), build(heights, names, blocks))
    val two = diffOf(
      wb(base),
      build(new Random(1).shuffle(heights), new Random(2).shuffle(names), blocks.reverse)
    )
    assertEquals(DiffCommands.renderJson(one), DiffCommands.renderJson(two))
    assertEquals(md(one), md(two))
    val refs = rowsOf(one).map(_.ref.takeWhile(_ != ':').toInt)
    assertEquals(refs, refs.sorted)
    assertEquals(refs.length, 200)
  }

  test("JSON: every new key is present and empty on a cell-only diff; values are typed") {
    val cellOnly = diffOf(wb(base), wb(Sheet("S").put(ref"A1", CellValue.Number(2))))
    val j = json(cellOnly)
    assertEquals(
      j.obj.keys.toVector,
      Vector(
        "identical",
        "sheetsAdded",
        "sheetsRemoved",
        "sheetOrder",
        "namesAdded",
        "namesRemoved",
        "namesChanged",
        "sheets"
      )
    )
    assertEquals(j("sheetOrder"), ujson.Null)
    val s = j("sheets")(0)
    val newKeys = Vector(
      "rows",
      "columns",
      "properties",
      "conditionalFormatsAdded",
      "conditionalFormatsRemoved",
      "conditionalFormatsChanged",
      "dataValidationsAdded",
      "dataValidationsRemoved",
      "dataValidationsChanged"
    )
    assertEquals(s.obj.keys.toVector.takeRight(9), newKeys)
    assertEquals(s.obj.keys.toVector.dropRight(9).last, "hyperlinksChanged")
    newKeys.foreach(k => assertEquals(s(k).arr.length, 0, k))

    val typed = withRows(
      base.copy(defaultColumnWidth = Some(12)).freezeAt(ref"B2"),
      5 -> RowProperties(height = Some(20), hidden = true, outlineLevel = Some(1))
    )
    val book = wb(typed, Sheet("T"))
    val hiddenT =
      book.setSheetState(SheetName.unsafe("T"), Some("hidden")).fold(e => fail(e.message), identity)
    val t = json(diffOf(wb(base, Sheet("T")), hiddenT))
    val changes = t("sheets")(0)("rows")(0)("changes").arr
    assert(changes(0)("after").numOpt.isDefined) // height
    assert(changes(1)("after").boolOpt.isDefined) // hidden
    assertEquals(changes(2)("after"), ujson.Num(1)) // outlineLevel
    val props = t("sheets")(0)("properties").arr
    assertEquals(props(0)("property").str, "defaultColumnWidth")
    assertEquals(props(0)("after").num, 12.0)
    assertEquals(props(1)("after").str, "B2")
    assertEquals(t("sheets")(1)("properties")(0)("after").str, "hidden")
  }

  test("markdown: the Summary gains '; N structure change(s)' only when N > 0") {
    val structural = diffOf(
      wb(base),
      wb(
        withRows(base, 1 -> height(30))
          .setColumnProperties(col("A"), ColumnProperties(width = Some(20)))
      )
    )
    val text = md(structural)
    assert(
      text.endsWith(
        "\nSummary: 0 changed, 0 added, 0 removed cell(s) across 1 sheet(s); 2 structure change(s)\n"
      ),
      text
    )
    val cellOnly = md(diffOf(wb(base), wb(Sheet("S").put(ref"A1", CellValue.Number(2)))))
    assert(
      cellOnly.endsWith("\nSummary: 1 changed, 0 added, 0 removed cell(s) across 1 sheet(s)\n"),
      cellOnly
    )
  }

  // ========== End to end ==========

  private def writeTemp(workbook: Workbook): IO[Path] =
    IO.blocking {
      val p = Files.createTempFile("xl-diff-structure", ".xlsx")
      p.toFile.deleteOnExit()
      p
    }.flatTap(p => ExcelIO.instance[IO].write(workbook, p))

  private val quiet = CliIO(_ => IO.unit, _ => IO.unit, IO.pure(""), _ => IO.unit)

  test("end to end: a row-height-only pair exits 1, and 0 under --cells-only") {
    for
      fa <- writeTemp(wb(base))
      fb <- writeTemp(wb(withRows(base, 1 -> height(30))))
      differs <- Main.runDiff(fa, fb, None, None, DiffFormat.Markdown, io = quiet)
      cellsOnly <- Main.runDiff(
        fa,
        fb,
        None,
        None,
        DiffFormat.Markdown,
        cellsOnly = true,
        io = quiet
      )
    yield
      assertEquals(differs, ExitCode(1))
      assertEquals(cellsOnly, ExitCode.Success)
  }
