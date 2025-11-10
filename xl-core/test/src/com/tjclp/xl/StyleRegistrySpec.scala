package com.tjclp.xl

import munit.FunSuite
import com.tjclp.xl.style.{StyleRegistry, CellStyle, StyleId, Font, Fill, Color, Border, BorderStyle, NumFmt, Align, HAlign, VAlign}

/** Tests for StyleRegistry */
class StyleRegistrySpec extends FunSuite:

  test("default registry contains CellStyle.default at index 0") {
    val registry = StyleRegistry.default
    assertEquals(registry.size, 1)
    assertEquals(registry.get(StyleId(0)), Some(CellStyle.default))
    assertEquals(registry.indexOf(CellStyle.default), Some(StyleId(0)))
  }

  test("register new style appends and returns index") {
    val registry = StyleRegistry.default
    val boldStyle = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))

    val (updated, idx) = registry.register(boldStyle)

    assertEquals(idx, StyleId(1), "New style should get index 1")
    assertEquals(updated.size, 2, "Should have 2 styles now")
    assertEquals(updated.get(StyleId(1)), Some(boldStyle))
  }

  test("register duplicate style returns existing index") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))

    val (registry1, idx1) = StyleRegistry.default.register(boldStyle)
    val (registry2, idx2) = registry1.register(boldStyle)

    assertEquals(idx1, idx2, "Same style should get same index")
    assertEquals(registry2.size, 2, "Should not duplicate")
  }

  test("register multiple distinct styles") {
    val bold = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))
    val red = CellStyle.default.withFill(Fill.Solid(Color.fromRgb(255, 0, 0)))
    val bordered = CellStyle.default.withBorder(Border.all(BorderStyle.Thin))

    val registry = StyleRegistry.default
    val (r1, idx1) = registry.register(bold)
    val (r2, idx2) = r1.register(red)
    val (r3, idx3) = r2.register(bordered)

    assertEquals(idx1, StyleId(1))
    assertEquals(idx2, StyleId(2))
    assertEquals(idx3, StyleId(3))
    assertEquals(r3.size, 4) // default + 3 custom
  }

  test("indexOf returns None for unregistered style") {
    val registry = StyleRegistry.default
    val customStyle = CellStyle.default.withFont(Font("Arial", 16.0))

    assertEquals(registry.indexOf(customStyle), None)
  }

  test("get returns None for invalid index") {
    val registry = StyleRegistry.default

    assertEquals(registry.get(StyleId(-1)), None)
    assertEquals(registry.get(StyleId(1)), None) // Only has index 0
    assertEquals(registry.get(StyleId(100)), None)
  }

  test("canonical key deduplication works") {
    // Two different instances with same properties
    val style1 = CellStyle.default
      .withFont(Font("Arial", 12.0, bold = true))
      .withFill(Fill.Solid(Color.Rgb(0xFF000000)))
    val style2 = CellStyle.default
      .withFont(Font("Arial", 12.0, bold = true))
      .withFill(Fill.Solid(Color.Rgb(0xFF000000)))

    val (r1, idx1) = StyleRegistry.default.register(style1)
    val (r2, idx2) = r1.register(style2)

    assertEquals(idx1, idx2, "Structurally equal styles should get same index")
    assertEquals(r2.size, 2, "Should not duplicate")
  }

  test("isEmpty returns true for default-only registry") {
    val registry = StyleRegistry.default
    assert(registry.isEmpty)

    val (updated, _) = registry.register(CellStyle.default.withFont(Font("Arial", 14.0, bold = true)))
    assert(!updated.isEmpty, "Should not be empty after adding custom style")
  }

  test("register preserves order") {
    val style1 = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))
    val style2 = CellStyle.default.withFill(Fill.Solid(Color.Rgb(0xFFFF0000)))
    val style3 = CellStyle.default.withBorder(Border.all(BorderStyle.Thick))

    val (r1, _) = StyleRegistry.default.register(style1)
    val (r2, _) = r1.register(style2)
    val (r3, _) = r2.register(style3)

    assertEquals(r3.get(StyleId(0)), Some(CellStyle.default))
    assertEquals(r3.get(StyleId(1)), Some(style1))
    assertEquals(r3.get(StyleId(2)), Some(style2))
    assertEquals(r3.get(StyleId(3)), Some(style3))
  }

  test("register complex style with all components") {
    val complexStyle = CellStyle(
      font = Font("Calibri", 12.0, bold = true, italic = true),
      fill = Fill.Solid(Color.Rgb(0xFFFFCC00)),
      border = Border.all(BorderStyle.Medium),
      numFmt = NumFmt.Currency,
      align = Align(HAlign.Center, VAlign.Middle, wrapText = true)
    )

    val (registry, idx) = StyleRegistry.default.register(complexStyle)

    assertEquals(idx, StyleId(1))
    assertEquals(registry.get(StyleId(1)), Some(complexStyle))
    assertEquals(registry.indexOf(complexStyle), Some(StyleId(1)))
  }

  test("multiple registrations of same style return same index") {
    val boldStyle = CellStyle.default.withFont(Font("Arial", 12.0, bold = true))

    var registry = StyleRegistry.default
    val indices = (1 to 10).map { _ =>
      val (updated, idx) = registry.register(boldStyle)
      registry = updated
      idx
    }

    assert(indices.forall(_ == StyleId(1)), "All registrations should return index 1")
    assertEquals(registry.size, 2, "Should have exactly 2 styles")
  }
