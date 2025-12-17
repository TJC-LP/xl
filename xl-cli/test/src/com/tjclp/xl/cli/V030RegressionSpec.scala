package com.tjclp.xl.cli

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.helpers.StyleBuilder
import com.tjclp.xl.cli.output.Format
import com.tjclp.xl.macros.ref
import com.tjclp.xl.styles.{CellStyle, StyleId}
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Regression tests for v0.3.0 bug fixes.
 *
 * Tests cover:
 *   - #85: Style merging (default behavior)
 *   - #87: Border display in cell command
 */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class V030RegressionSpec extends CatsEffectSuite:

  // ========== #85: Style Merging Tests ==========

  test("mergeStyles: preserves existing bold when adding italic") {
    val existing = CellStyle.default.copy(
      font = Font.default.withBold(true)
    )
    val newStyle = CellStyle.default.copy(
      font = Font.default.withItalic(true)
    )

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assert(merged.font.bold, "Bold should be preserved")
    assert(merged.font.italic, "Italic should be added")
  }

  test("mergeStyles: preserves existing fill when adding font color") {
    val existing = CellStyle.default.copy(
      fill = Fill.Solid(Color.fromRgb(0, 0, 128)) // Navy
    )
    val newStyle = CellStyle.default.copy(
      font = Font.default.withColor(Color.fromRgb(255, 0, 0)) // Red text
    )

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    merged.fill match
      case Fill.Solid(color) =>
        color match
          case Color.Rgb(argb) =>
            assertEquals(argb & 0xffffff, 0x000080, "Fill should be preserved as navy")
          case _ => fail("Fill color should be RGB")
      case _ => fail("Fill should be solid")

    merged.font.color match
      case Some(Color.Rgb(argb)) =>
        assertEquals(argb & 0xffffff, 0xff0000, "Font color should be red")
      case _ => fail("Font color should be set to red")
  }

  test("mergeStyles: new numFmt overrides existing") {
    val existing = CellStyle.default.copy(numFmt = NumFmt.Decimal)
    val newStyle = CellStyle.default.copy(numFmt = NumFmt.Percent)

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assertEquals(merged.numFmt, NumFmt.Percent, "NumFmt should be overridden")
  }

  test("mergeStyles: General numFmt does not override existing") {
    val existing = CellStyle.default.copy(numFmt = NumFmt.Currency)
    val newStyle = CellStyle.default.copy(numFmt = NumFmt.General)

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assertEquals(merged.numFmt, NumFmt.Currency, "General should not override existing")
  }

  test("mergeStyles: border overrides when new is non-empty") {
    val existing = CellStyle.default
    val newBorder = Border.all(BorderStyle.Thin, None)
    val newStyle = CellStyle.default.copy(border = newBorder)

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assertEquals(merged.border.top.style, BorderStyle.Thin, "Border should be set")
  }

  test("mergeStyles: alignment horizontal overrides when non-default") {
    val existing = CellStyle.default.copy(
      align = Align.default.withVAlign(VAlign.Middle)
    )
    val newStyle = CellStyle.default.copy(
      align = Align.default.withHAlign(HAlign.Center)
    )

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assertEquals(merged.align.horizontal, HAlign.Center, "HAlign should be Center")
    assertEquals(merged.align.vertical, VAlign.Middle, "VAlign should be preserved")
  }

  test("mergeStyles: font size overrides when different from default") {
    val existing = CellStyle.default.copy(
      font = Font.default.withBold(true)
    )
    val newStyle = CellStyle.default.copy(
      font = Font.default.withSize(14.0)
    )

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assert(merged.font.bold, "Bold should be preserved")
    assertEquals(merged.font.sizePt, 14.0, "Font size should be 14")
  }

  // ========== #87: Border Display Tests ==========

  test("formatStyle: displays thin border on all sides") {
    IO {
      val style = CellStyle.default.copy(
        border = Border.all(BorderStyle.Thin, None)
      )

      val formatted = formatStyleHelper(Some(style))

      assert(formatted.isDefined, "Style should be formatted")
      assert(formatted.get.contains("Border:"), "Should contain Border label")
      assert(formatted.get.contains("Thin"), "Should mention Thin style")
      assert(formatted.get.contains("all sides"), "Should indicate all sides")
    }
  }

  test("formatStyle: displays individual border sides") {
    IO {
      val style = CellStyle.default.copy(
        border = Border(
          top = BorderSide(BorderStyle.Thick, None),
          bottom = BorderSide(BorderStyle.None, None),
          left = BorderSide(BorderStyle.Thin, None),
          right = BorderSide(BorderStyle.None, None)
        )
      )

      val formatted = formatStyleHelper(Some(style))

      assert(formatted.isDefined, "Style should be formatted")
      assert(formatted.get.contains("top: Thick"), "Should mention top Thick")
      assert(formatted.get.contains("left: Thin"), "Should mention left Thin")
      assert(!formatted.get.contains("bottom:"), "Should not mention bottom None")
      assert(!formatted.get.contains("right:"), "Should not mention right None")
    }
  }

  test("formatStyle: no border section for empty borders") {
    IO {
      val style = CellStyle.default // No borders

      val formatted = formatStyleHelper(Some(style))

      // Should be None because default style has no non-default properties
      assertEquals(formatted, None, "Default style should format to None")
    }
  }

  test("formatStyle: displays font with border combined") {
    IO {
      val style = CellStyle.default.copy(
        font = Font.default.withBold(true),
        border = Border.all(BorderStyle.Medium, None)
      )

      val formatted = formatStyleHelper(Some(style))

      assert(formatted.isDefined, "Style should be formatted")
      assert(formatted.get.contains("Font:"), "Should contain Font section")
      assert(formatted.get.contains("bold"), "Should mention bold")
      assert(formatted.get.contains("Border:"), "Should contain Border section")
      assert(formatted.get.contains("Medium"), "Should mention Medium border")
    }
  }

  // ========== Per-Side Border Control Tests (PR #91) ==========

  test("mergeStyles: per-side border merge preserves existing top when adding bottom") {
    val existing = CellStyle.default.copy(
      border = Border(
        top = BorderSide(BorderStyle.Thick, None),
        bottom = BorderSide.none,
        left = BorderSide.none,
        right = BorderSide.none
      )
    )
    val newStyle = CellStyle.default.copy(
      border = Border(
        top = BorderSide.none,
        bottom = BorderSide(BorderStyle.Thin, None),
        left = BorderSide.none,
        right = BorderSide.none
      )
    )

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assertEquals(merged.border.top.style, BorderStyle.Thick, "Top should be preserved")
    assertEquals(merged.border.bottom.style, BorderStyle.Thin, "Bottom should be added")
    assertEquals(merged.border.left.style, BorderStyle.None, "Left should remain none")
    assertEquals(merged.border.right.style, BorderStyle.None, "Right should remain none")
  }

  test("mergeStyles: per-side border merge preserves all existing sides when new has none") {
    val existing = CellStyle.default.copy(
      border = Border.all(BorderStyle.Medium, None)
    )
    val newStyle = CellStyle.default.copy(
      border = Border.none
    )

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assertEquals(merged.border.top.style, BorderStyle.Medium, "Top should be preserved")
    assertEquals(merged.border.bottom.style, BorderStyle.Medium, "Bottom should be preserved")
    assertEquals(merged.border.left.style, BorderStyle.Medium, "Left should be preserved")
    assertEquals(merged.border.right.style, BorderStyle.Medium, "Right should be preserved")
  }

  test("mergeStyles: per-side border new side overrides existing same side") {
    val existing = CellStyle.default.copy(
      border = Border(
        top = BorderSide(BorderStyle.Thin, None),
        bottom = BorderSide(BorderStyle.Thin, None),
        left = BorderSide(BorderStyle.Thin, None),
        right = BorderSide(BorderStyle.Thin, None)
      )
    )
    val newStyle = CellStyle.default.copy(
      border = Border(
        top = BorderSide(BorderStyle.Thick, None), // Override top
        bottom = BorderSide.none, // Keep existing
        left = BorderSide(BorderStyle.Double, None), // Override left
        right = BorderSide.none // Keep existing
      )
    )

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assertEquals(merged.border.top.style, BorderStyle.Thick, "Top should be overridden to Thick")
    assertEquals(merged.border.bottom.style, BorderStyle.Thin, "Bottom should be preserved as Thin")
    assertEquals(merged.border.left.style, BorderStyle.Double, "Left should be overridden to Double")
    assertEquals(merged.border.right.style, BorderStyle.Thin, "Right should be preserved as Thin")
  }

  test("mergeStyles: per-side border preserves color from existing when new has no color") {
    val blueColor = Color.fromRgb(0, 0, 255)
    val existing = CellStyle.default.copy(
      border = Border(
        top = BorderSide(BorderStyle.Thin, Some(blueColor)),
        bottom = BorderSide.none,
        left = BorderSide.none,
        right = BorderSide.none
      )
    )
    val newStyle = CellStyle.default.copy(
      border = Border(
        top = BorderSide.none, // Don't touch top
        bottom = BorderSide(BorderStyle.Thin, None), // Add bottom without color
        left = BorderSide.none,
        right = BorderSide.none
      )
    )

    val merged = StyleBuilder.mergeStyles(existing, newStyle)

    assertEquals(merged.border.top.style, BorderStyle.Thin, "Top style preserved")
    assertEquals(merged.border.top.color, Some(blueColor), "Top color preserved")
    assertEquals(merged.border.bottom.style, BorderStyle.Thin, "Bottom added")
    assertEquals(merged.border.bottom.color, None, "Bottom has no color")
  }

  test("buildCellStyle: --border sets all sides") {
    StyleBuilder
      .buildCellStyle(
        bold = false,
        italic = false,
        underline = false,
        bg = None,
        fg = None,
        fontSize = None,
        fontName = None,
        align = None,
        valign = None,
        wrap = false,
        numFormat = None,
        border = Some("thin"),
        borderTop = None,
        borderRight = None,
        borderBottom = None,
        borderLeft = None,
        borderColor = None
      )
      .map { style =>
        assertEquals(style.border.top.style, BorderStyle.Thin)
        assertEquals(style.border.right.style, BorderStyle.Thin)
        assertEquals(style.border.bottom.style, BorderStyle.Thin)
        assertEquals(style.border.left.style, BorderStyle.Thin)
      }
  }

  test("buildCellStyle: per-side options override --border") {
    StyleBuilder
      .buildCellStyle(
        bold = false,
        italic = false,
        underline = false,
        bg = None,
        fg = None,
        fontSize = None,
        fontName = None,
        align = None,
        valign = None,
        wrap = false,
        numFormat = None,
        border = Some("thin"), // Base: all thin
        borderTop = Some("thick"), // Override: top thick
        borderRight = None, // Keep thin
        borderBottom = Some("medium"), // Override: bottom medium
        borderLeft = None, // Keep thin
        borderColor = None
      )
      .map { style =>
        assertEquals(style.border.top.style, BorderStyle.Thick, "Top overridden to Thick")
        assertEquals(style.border.right.style, BorderStyle.Thin, "Right stays Thin")
        assertEquals(style.border.bottom.style, BorderStyle.Medium, "Bottom overridden to Medium")
        assertEquals(style.border.left.style, BorderStyle.Thin, "Left stays Thin")
      }
  }

  test("buildCellStyle: per-side options only (no --border base)") {
    StyleBuilder
      .buildCellStyle(
        bold = false,
        italic = false,
        underline = false,
        bg = None,
        fg = None,
        fontSize = None,
        fontName = None,
        align = None,
        valign = None,
        wrap = false,
        numFormat = None,
        border = None, // No base
        borderTop = Some("thick"), // Only top
        borderRight = None,
        borderBottom = None,
        borderLeft = Some("dashed"), // Only left
        borderColor = None
      )
      .map { style =>
        assertEquals(style.border.top.style, BorderStyle.Thick, "Top is Thick")
        assertEquals(style.border.right.style, BorderStyle.None, "Right is None")
        assertEquals(style.border.bottom.style, BorderStyle.None, "Bottom is None")
        assertEquals(style.border.left.style, BorderStyle.Dashed, "Left is Dashed")
      }
  }

  test("buildCellStyle: border color applies to all specified sides") {
    StyleBuilder
      .buildCellStyle(
        bold = false,
        italic = false,
        underline = false,
        bg = None,
        fg = None,
        fontSize = None,
        fontName = None,
        align = None,
        valign = None,
        wrap = false,
        numFormat = None,
        border = None,
        borderTop = Some("thin"),
        borderRight = None,
        borderBottom = Some("thin"),
        borderLeft = None,
        borderColor = Some("#FF0000") // Red
      )
      .map { style =>
        // Top and bottom should have red color
        assert(style.border.top.color.isDefined, "Top should have color")
        assert(style.border.bottom.color.isDefined, "Bottom should have color")
        // Right and left are None style, so color doesn't matter
        assertEquals(style.border.right.style, BorderStyle.None)
        assertEquals(style.border.left.style, BorderStyle.None)
      }
  }

  test("buildCellStyle: no borders when all options are None") {
    StyleBuilder
      .buildCellStyle(
        bold = false,
        italic = false,
        underline = false,
        bg = None,
        fg = None,
        fontSize = None,
        fontName = None,
        align = None,
        valign = None,
        wrap = false,
        numFormat = None,
        border = None,
        borderTop = None,
        borderRight = None,
        borderBottom = None,
        borderLeft = None,
        borderColor = None
      )
      .map { style =>
        assertEquals(style.border, Border.none, "Border should be none")
      }
  }

  // Helper to access Format.formatStyle (which is private)
  // We test via the public cellInfo method
  private def formatStyleHelper(style: Option[CellStyle]): Option[String] =
    // Use reflection or test via the full cellInfo output
    val output = Format.cellInfo(
      ref = ref"A1",
      value = CellValue.Text("test"),
      formatted = "test",
      style = style,
      comment = None,
      hyperlink = None,
      dependencies = Vector.empty,
      dependents = Vector.empty
    )

    // Extract the Style section if present
    val lines = output.split("\n")
    val styleStart = lines.indexWhere(_.startsWith("Style:"))
    if styleStart >= 0 then
      val styleLines = lines.drop(styleStart).takeWhile(l => l.startsWith("Style:") || l.startsWith("  "))
      Some(styleLines.mkString("\n"))
    else None
