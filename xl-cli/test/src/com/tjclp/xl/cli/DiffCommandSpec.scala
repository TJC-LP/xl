package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.{ExitCode, IO}
import munit.CatsEffectSuite

import com.tjclp.xl.{Sheet, Workbook, style, given}
import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.cli.commands.DiffCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.font.Font

/**
 * Tests for the diff command (GH-137).
 *
 * Covers: cell value/formula/style comparison, added/removed cells, sheet add/remove, merge,
 * comment and hyperlink deltas, the stable JSON schema, markdown rendering, sheet filtering, and
 * the diff-tool exit-code convention (0 identical, 1 differs).
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial", "org.wartremover.warts.IterableOps"))
class DiffCommandSpec extends CatsEffectSuite:

  private def wb(sheets: Sheet*): Workbook = Workbook(sheets.toVector)

  // ========== Identity ==========

  test("diff: identical workbooks report identical") {
    val s = Sheet("S").put(ref"A1", CellValue.Text("x")).put(ref"B2", CellValue.Number(5))
    val result = DiffCommands.computeDiff(wb(s), wb(s), None)
    assert(result.isRight, s"Got $result")
    assert(result.toOption.get.identical, "Identical workbooks should produce an empty diff")
  }

  test("diff: explicit Empty cell vs missing cell is not a difference") {
    val a = Sheet("S").put(ref"A1", CellValue.Text("x")).put(ref"B1", CellValue.Empty)
    val b = Sheet("S").put(ref"A1", CellValue.Text("x"))
    val result = DiffCommands.computeDiff(wb(a), wb(b), None)
    assert(result.toOption.get.identical, "Trivial empty cells should not count as differences")
  }

  // ========== Cell changes ==========

  test("diff: changed value is reported with before/after") {
    val a = Sheet("S").put(ref"A5", CellValue.Text("Revenue"))
    val b = Sheet("S").put(ref"A5", CellValue.Text("Total Revenue"))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    assert(!diff.identical)
    val sheetDiff = diff.sheets.head
    assertEquals(sheetDiff.changed.length, 1)
    val change = sheetDiff.changed.head
    assertEquals(change.ref.toA1, "A5")
    assertEquals(change.before.value, "Revenue")
    assertEquals(change.after.value, "Total Revenue")
    assert(!change.styleChanged)
  }

  test("diff: changed formula text is reported") {
    val a = Sheet("S").put(ref"C5", CellValue.Formula("=SUM(C2:C4)", None))
    val b = Sheet("S").put(ref"C5", CellValue.Formula("=SUM(C2:C4)*1.1", None))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    val change = diff.sheets.head.changed.head
    assertEquals(change.before.formula, Some("=SUM(C2:C4)"))
    assertEquals(change.after.formula, Some("=SUM(C2:C4)*1.1"))
  }

  test("diff: identical formula with different cached value is not a change") {
    val a = Sheet("S").put(ref"C1", CellValue.Formula("=A1", Some(CellValue.Number(1))))
    val b = Sheet("S").put(ref"C1", CellValue.Formula("=A1", Some(CellValue.Number(2))))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    assert(diff.identical, "Cached values are derived; formula text is the comparison key")
  }

  test("diff: style-only change sets styleChanged with equal values") {
    val bold = CellStyle.default.withFont(Font.default.withBold(true))
    val a = Sheet("S").put(ref"A1", CellValue.Text("x"))
    val b = Sheet("S").put(ref"A1", CellValue.Text("x")).style(ref"A1", bold)
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    val change = diff.sheets.head.changed.head
    assert(change.styleChanged, "Resolved style differs")
    assertEquals(change.before.value, change.after.value)
  }

  test("diff: added and removed cells are reported") {
    val a = Sheet("S").put(ref"A1", CellValue.Text("keep")).put(ref"E5", CellValue.Text("old"))
    val b = Sheet("S").put(ref"A1", CellValue.Text("keep")).put(ref"D5", CellValue.Text("new"))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    val sheetDiff = diff.sheets.head
    assertEquals(sheetDiff.added.map(_._1.toA1), Vector("D5"))
    assertEquals(sheetDiff.removed.map(_._1.toA1), Vector("E5"))
  }

  test("diff: changed refs are sorted in row-major A1 order") {
    val a = Sheet("S")
      .put(ref"B2", CellValue.Number(1))
      .put(ref"A2", CellValue.Number(1))
      .put(ref"A1", CellValue.Number(1))
    val b = Sheet("S")
      .put(ref"B2", CellValue.Number(2))
      .put(ref"A2", CellValue.Number(2))
      .put(ref"A1", CellValue.Number(2))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    assertEquals(diff.sheets.head.changed.map(_.ref.toA1), Vector("A1", "A2", "B2"))
  }

  // ========== Sheet-level differences ==========

  test("diff: added and removed sheets are reported") {
    val a = wb(Sheet("Common"), Sheet("OnlyA"))
    val b = wb(Sheet("Common"), Sheet("OnlyB"))
    val diff = DiffCommands.computeDiff(a, b, None).toOption.get
    assertEquals(diff.sheetsAdded, Vector("OnlyB"))
    assertEquals(diff.sheetsRemoved, Vector("OnlyA"))
    assert(!diff.identical)
  }

  test("diff: --sheet filter compares only the named sheet") {
    val a = wb(
      Sheet("S1").put(ref"A1", CellValue.Number(1)),
      Sheet("S2").put(ref"A1", CellValue.Number(1))
    )
    val b = wb(
      Sheet("S1").put(ref"A1", CellValue.Number(2)),
      Sheet("S2").put(ref"A1", CellValue.Number(2))
    )
    val diff = DiffCommands.computeDiff(a, b, Some("S2")).toOption.get
    assertEquals(diff.sheets.map(_.name), Vector("S2"))
  }

  test("diff: --sheet filter missing in both workbooks errors") {
    val result = DiffCommands.computeDiff(wb(Sheet("S")), wb(Sheet("S")), Some("Nope"))
    assert(result.isLeft, s"Expected Left, got $result")
  }

  // ========== Merges, comments, hyperlinks ==========

  test("diff: merge deltas are reported") {
    val range = CellRange.parse("A1:B2").toOption.get
    val a = Sheet("S").put(ref"A1", CellValue.Text("t"))
    val b = Sheet("S").put(ref"A1", CellValue.Text("t")).merge(range)
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    assertEquals(diff.sheets.head.mergesAdded, Vector("A1:B2"))
    assertEquals(diff.sheets.head.mergesRemoved, Vector.empty)
  }

  test("diff: comment deltas are reported (added/removed/changed)") {
    val a = Sheet("S")
      .put(ref"A1", CellValue.Text("x"))
      .comment(ref"A1", Comment.plainText("old", Some("Bob")))
      .put(ref"B1", CellValue.Text("y"))
      .comment(ref"B1", Comment.plainText("gone", None))
    val b = Sheet("S")
      .put(ref"A1", CellValue.Text("x"))
      .comment(ref"A1", Comment.plainText("new", Some("Bob")))
      .put(ref"B1", CellValue.Text("y"))
      .put(ref"C1", CellValue.Text("z"))
      .comment(ref"C1", Comment.plainText("fresh", None))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    val sd = diff.sheets.head
    assertEquals(sd.commentsAdded, Vector("C1"))
    assertEquals(sd.commentsRemoved, Vector("B1"))
    assertEquals(sd.commentsChanged, Vector("A1"))
  }

  test("diff: hyperlink deltas are reported") {
    val base = Sheet("S").put(ref"A1", CellValue.Text("link")).put(ref"B1", CellValue.Text("x"))
    val a = base.put(base(ref"A1").withHyperlink("https://old.example"))
    val b = base
      .put(base(ref"A1").withHyperlink("https://new.example"))
      .put(base(ref"B1").withHyperlink("https://added.example"))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    val sd = diff.sheets.head
    assertEquals(sd.hyperlinksChanged, Vector("A1"))
    assertEquals(sd.hyperlinksAdded, Vector("B1"))
  }

  // ========== Rendering ==========

  test("diff: markdown output groups by sheet with A1 refs and summary") {
    val a = Sheet("S1").put(ref"A5", CellValue.Text("Revenue")).put(ref"E5", CellValue.Text("Old"))
    val b = Sheet("S1").put(ref"A5", CellValue.Text("Total")).put(ref"D5", CellValue.Text("New"))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    val md = DiffCommands.renderMarkdown(diff, "a.xlsx", "b.xlsx")
    assert(md.contains("a.xlsx"), md)
    assert(md.contains("b.xlsx"), md)
    assert(md.contains("S1"), md)
    assert(md.contains("A5"), md)
    assert(md.contains("D5"), md)
    assert(md.contains("E5"), md)
    assert(md.toLowerCase.contains("summary"), md)
  }

  test("diff: markdown output for identical files says identical") {
    val s = Sheet("S").put(ref"A1", CellValue.Number(1))
    val diff = DiffCommands.computeDiff(wb(s), wb(s), None).toOption.get
    val md = DiffCommands.renderMarkdown(diff, "a.xlsx", "b.xlsx")
    assert(md.toLowerCase.contains("identical"), md)
  }

  test("diff: JSON output follows the stable schema") {
    val bold = CellStyle.default.withFont(Font.default.withBold(true))
    val a = Sheet("S1")
      .put(ref"A5", CellValue.Text("Revenue"))
      .put(ref"E5", CellValue.Text("Old"))
    val b = Sheet("S1")
      .put(ref"A5", CellValue.Text("Total"))
      .style(ref"A5", bold)
      .put(ref"D5", CellValue.Text("New"))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    val json = ujson.read(DiffCommands.renderJson(diff))

    assertEquals(json("identical").bool, false)
    assertEquals(json("sheetsAdded").arr.toVector, Vector.empty[ujson.Value])
    assertEquals(json("sheetsRemoved").arr.toVector, Vector.empty[ujson.Value])
    val sheet = json("sheets").arr.head
    assertEquals(sheet("name").str, "S1")
    val change = sheet("changed").arr.head
    assertEquals(change("ref").str, "A5")
    assertEquals(change("before")("value").str, "Revenue")
    assertEquals(change("after")("value").str, "Total")
    assertEquals(change("styleChanged").bool, true)
    val added = sheet("added").arr.head
    assertEquals(added("ref").str, "D5")
    assertEquals(added("value").str, "New")
    val removed = sheet("removed").arr.head
    assertEquals(removed("ref").str, "E5")
  }

  test("diff: JSON formula change carries formula text, null for plain values") {
    val a = Sheet("S").put(ref"C1", CellValue.Formula("=SUM(A1:A2)", None))
    val b = Sheet("S").put(ref"C1", CellValue.Formula("=SUM(A1:A3)", None))
    val diff = DiffCommands.computeDiff(wb(a), wb(b), None).toOption.get
    val json = ujson.read(DiffCommands.renderJson(diff))
    val change = json("sheets").arr.head("changed").arr.head
    assertEquals(change("before")("formula").str, "=SUM(A1:A2)")
    assertEquals(change("after")("formula").str, "=SUM(A1:A3)")
  }

  // ========== Exit codes (end-to-end through Main.runDiff) ==========

  private def writeTemp(workbook: Workbook): IO[Path] =
    IO.blocking {
      val p = Files.createTempFile("xl-diff", ".xlsx")
      p.toFile.deleteOnExit()
      p
    }.flatTap(p => ExcelIO.instance[IO].write(workbook, p))

  test("diff: exit code 0 for identical files") {
    val s = Sheet("S").put(ref"A1", CellValue.Number(42))
    for
      fa <- writeTemp(wb(s))
      fb <- writeTemp(wb(s))
      code <- Main.runDiff(fa, fb, None, None, DiffFormat.Markdown)
    yield assertEquals(code, ExitCode.Success)
  }

  test("diff: exit code 1 when files differ") {
    for
      fa <- writeTemp(wb(Sheet("S").put(ref"A1", CellValue.Number(1))))
      fb <- writeTemp(wb(Sheet("S").put(ref"A1", CellValue.Number(2))))
      code <- Main.runDiff(fa, fb, None, None, DiffFormat.Json)
    yield assertEquals(code, ExitCode(1))
  }

  test("diff: exit code 2 when a file cannot be read") {
    for
      fa <- writeTemp(wb(Sheet("S")))
      code <- Main.runDiff(
        fa,
        java.nio.file.Paths.get("/nonexistent/nope.xlsx"),
        None,
        None,
        DiffFormat.Markdown
      )
    yield assertEquals(code, ExitCode(2))
  }
