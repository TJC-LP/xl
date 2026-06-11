package com.tjclp.xl.cli.helpers

import munit.FunSuite

import com.tjclp.xl.cli.helpers.MarkdownTableParser.ColumnAlignment

/**
 * Tests for the GFM (GitHub Flavored Markdown) pipe-table parser (GH-159).
 *
 * Pins: header/delimiter/body structure, alignment markers, escaped pipes, whitespace trimming,
 * optional outer pipes, ragged-row normalization, and table boundary detection.
 */
class MarkdownTableParserSpec extends FunSuite:

  private def lines(ls: String*): String = ls.mkString("\n")

  test("parse: basic table with header, delimiter, and body rows") {
    val input = lines(
      "| Name | Qty |",
      "|------|-----|",
      "| Apple | 3 |",
      "| Pear | 5 |"
    )
    val result = MarkdownTableParser.parse(input)
    assert(result.isRight, s"Expected Right, got $result")
    val table = result.toOption.get
    assertEquals(table.rows.length, 3)
    assertEquals(table.rows(0), Vector("Name", "Qty"))
    assertEquals(table.rows(1), Vector("Apple", "3"))
    assertEquals(table.rows(2), Vector("Pear", "5"))
  }

  test("parse: alignment markers map to left/center/right/none") {
    val input = lines(
      "| A | B | C | D |",
      "|:---|:---:|---:|---|",
      "| 1 | 2 | 3 | 4 |"
    )
    val result = MarkdownTableParser.parse(input)
    assert(result.isRight, s"Expected Right, got $result")
    val table = result.toOption.get
    assertEquals(
      table.alignments,
      Vector(
        ColumnAlignment.Left,
        ColumnAlignment.Center,
        ColumnAlignment.Right,
        ColumnAlignment.None
      )
    )
  }

  test("parse: escaped pipes inside cells become literal pipes") {
    val input = lines(
      "| Expr | Result |",
      "|------|--------|",
      "| a \\| b | 10 |"
    )
    val result = MarkdownTableParser.parse(input)
    assert(result.isRight, s"Expected Right, got $result")
    val table = result.toOption.get
    assertEquals(table.rows(1), Vector("a | b", "10"))
  }

  test("parse: surrounding whitespace in cells is trimmed") {
    val input = lines(
      "|   Name    |   Qty   |",
      "|-----------|---------|",
      "|   Apple   |   3     |"
    )
    val result = MarkdownTableParser.parse(input)
    assert(result.isRight)
    val table = result.toOption.get
    assertEquals(table.rows(0), Vector("Name", "Qty"))
    assertEquals(table.rows(1), Vector("Apple", "3"))
  }

  test("parse: leading/trailing outer pipes are optional") {
    val input = lines(
      "Name | Qty",
      "--- | ---",
      "Apple | 3"
    )
    val result = MarkdownTableParser.parse(input)
    assert(result.isRight, s"Expected Right, got $result")
    val table = result.toOption.get
    assertEquals(table.rows(0), Vector("Name", "Qty"))
    assertEquals(table.rows(1), Vector("Apple", "3"))
  }

  test("parse: ragged body rows are padded/truncated to delimiter width") {
    val input = lines(
      "| A | B | C |",
      "|---|---|---|",
      "| 1 | 2 |",
      "| 1 | 2 | 3 | 4 |"
    )
    val result = MarkdownTableParser.parse(input)
    assert(result.isRight)
    val table = result.toOption.get
    assertEquals(table.rows(1), Vector("1", "2", ""))
    assertEquals(table.rows(2), Vector("1", "2", "3"))
  }

  test("parse: preamble text before the table is skipped") {
    val input = lines(
      "Some intro prose.",
      "",
      "| X | Y |",
      "|---|---|",
      "| 1 | 2 |"
    )
    val result = MarkdownTableParser.parse(input)
    assert(result.isRight, s"Expected Right, got $result")
    assertEquals(result.toOption.get.rows(0), Vector("X", "Y"))
  }

  test("parse: table ends at first blank line") {
    val input = lines(
      "| X | Y |",
      "|---|---|",
      "| 1 | 2 |",
      "",
      "| not | parsed |"
    )
    val result = MarkdownTableParser.parse(input)
    assert(result.isRight)
    assertEquals(result.toOption.get.rows.length, 2)
  }

  test("parse: input with no table returns Left") {
    val result = MarkdownTableParser.parse("just some text\nwith no table at all")
    assert(result.isLeft, s"Expected Left, got $result")
  }

  test("parse: pipe line without valid delimiter row returns Left") {
    val input = lines(
      "| A | B |",
      "| 1 | 2 |"
    )
    val result = MarkdownTableParser.parse(input)
    assert(result.isLeft, s"Expected Left (no delimiter row), got $result")
  }

  test("parse: empty input returns Left") {
    assert(MarkdownTableParser.parse("").isLeft)
  }
