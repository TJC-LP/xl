package com.tjclp.xl

import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen, Prop}
import org.scalacheck.Prop.*
import cats.syntax.all.*
import com.tjclp.xl.styles.{*, given}

/** Property tests for StylePatch monoid laws */
class StylePatchSpec extends ScalaCheckSuite:

  // Reuse generators from StyleSpec
  val styleSpec = new StyleSpec
  import styleSpec.{genFont, genFill, genBorder, genNumFmt, genAlign, genCellStyle}
  import styleSpec.{given Arbitrary[Font], given Arbitrary[Fill], given Arbitrary[Border], given Arbitrary[CellStyle]}

  // Generator for StylePatch
  val genStylePatch: Gen[StylePatch] = Gen.oneOf(
    genFont.map(StylePatch.SetFont.apply),
    genFill.map(StylePatch.SetFill.apply),
    genBorder.map(StylePatch.SetBorder.apply),
    genNumFmt.map(StylePatch.SetNumFmt.apply),
    genAlign.map(StylePatch.SetAlign.apply)
  )

  given Arbitrary[StylePatch] = Arbitrary(genStylePatch)

  // Import the Monoid given
  import StylePatch.{given, *}

  // ========== Monoid Law Tests ==========

  property("StylePatch Monoid: left identity (empty |+| p == p)") {
    forAll { (patch: StylePatch) =>
      val combined = (StylePatch.empty: StylePatch) |+| (patch: StylePatch)

      // Apply both to default style and verify they produce same result
      val style = CellStyle.default
      val result1 = applyPatch(style, combined)
      val result2 = applyPatch(style, patch)

      assertEquals(result1, result2)
      true
    }
  }

  property("StylePatch Monoid: right identity (p |+| empty == p)") {
    forAll { (patch: StylePatch) =>
      val combined = (patch: StylePatch) |+| (StylePatch.empty: StylePatch)

      // Apply both to default style and verify they produce same result
      val style = CellStyle.default
      val result1 = applyPatch(style, combined)
      val result2 = applyPatch(style, patch)

      assertEquals(result1, result2)
      true
    }
  }

  property("StylePatch Monoid: associativity ((p1 |+| p2) |+| p3 == p1 |+| (p2 |+| p3))") {
    forAll { (p1: StylePatch, p2: StylePatch, p3: StylePatch) =>
      val left = ((p1: StylePatch) |+| (p2: StylePatch)) |+| (p3: StylePatch)
      val right = (p1: StylePatch) |+| ((p2: StylePatch) |+| (p3: StylePatch))

      // Apply both to same style and verify results are identical
      val style = CellStyle.default
      val leftResult = applyPatch(style, left)
      val rightResult = applyPatch(style, right)

      assertEquals(leftResult, rightResult)
      true
    }
  }

  // ========== Patch Application Tests ==========

  test("SetFont patch updates font") {
    val style = CellStyle.default
    val newFont = Font("Arial", 14.0, bold = true)
    val patch = StylePatch.SetFont(newFont)

    val result = applyPatch(style, patch)
    assertEquals(result.font, newFont)
    assertEquals(result.fill, style.fill) // Other properties unchanged
  }

  test("SetFill patch updates fill") {
    val style = CellStyle.default
    val newFill = Fill.Solid(Color.Rgb(0xFF0000FF))
    val patch = StylePatch.SetFill(newFill)

    val result = applyPatch(style, patch)
    assertEquals(result.fill, newFill)
    assertEquals(result.font, style.font) // Other properties unchanged
  }

  test("SetBorder patch updates border") {
    val style = CellStyle.default
    val newBorder = Border.all(BorderStyle.Thick)
    val patch = StylePatch.SetBorder(newBorder)

    val result = applyPatch(style, patch)
    assertEquals(result.border, newBorder)
    assertEquals(result.font, style.font) // Other properties unchanged
  }

  test("SetNumFmt patch updates number format") {
    val style = CellStyle.default
    val newNumFmt = NumFmt.Percent
    val patch = StylePatch.SetNumFmt(newNumFmt)

    val result = applyPatch(style, patch)
    assertEquals(result.numFmt, newNumFmt)
    assertEquals(result.font, style.font) // Other properties unchanged
  }

  test("SetAlign patch updates alignment") {
    val style = CellStyle.default
    val newAlign = Align(HAlign.Center, VAlign.Middle, wrapText = true)
    val patch = StylePatch.SetAlign(newAlign)

    val result = applyPatch(style, patch)
    assertEquals(result.align, newAlign)
    assertEquals(result.font, style.font) // Other properties unchanged
  }

  test("Batch patch applies multiple patches in order") {
    val style = CellStyle.default
    val font = Font("Arial", 14.0, bold = true)
    val fill = Fill.Solid(Color.Rgb(0xFFFF0000))
    val border = Border.all(BorderStyle.Thin)

    val batch = StylePatch.Batch(Vector(
      StylePatch.SetFont(font),
      StylePatch.SetFill(fill),
      StylePatch.SetBorder(border)
    ))

    val result = applyPatch(style, batch)
    assertEquals(result.font, font)
    assertEquals(result.fill, fill)
    assertEquals(result.border, border)
  }

  // ========== Idempotence Tests ==========

  property("SetFont patch is idempotent") {
    forAll { (font: Font) =>
      val style = CellStyle.default
      val patch = StylePatch.SetFont(font)

      val once = applyPatch(style, patch)
      val twice = applyPatch(once, patch)

      assertEquals(once, twice)
      true
    }
  }

  property("SetFill patch is idempotent") {
    forAll { (fill: Fill) =>
      val style = CellStyle.default
      val patch = StylePatch.SetFill(fill)

      val once = applyPatch(style, patch)
      val twice = applyPatch(once, patch)

      assertEquals(once, twice)
      true
    }
  }

  property("SetBorder patch is idempotent") {
    forAll { (border: Border) =>
      val style = CellStyle.default
      val patch = StylePatch.SetBorder(border)

      val once = applyPatch(style, patch)
      val twice = applyPatch(once, patch)

      assertEquals(once, twice)
      true
    }
  }

  // ========== Override Semantics Tests ==========

  test("Later SetFont overrides earlier SetFont") {
    val style = CellStyle.default
    val font1 = Font("Arial", 12.0)
    val font2 = Font("Calibri", 14.0)

    val patch = (StylePatch.SetFont(font1): StylePatch) |+| (StylePatch.SetFont(font2): StylePatch)
    val result = applyPatch(style, patch)

    assertEquals(result.font, font2)
  }

  test("Later SetFill overrides earlier SetFill") {
    val style = CellStyle.default
    val fill1 = Fill.Solid(Color.Rgb(0xFF0000FF))
    val fill2 = Fill.Solid(Color.Rgb(0xFF00FF00))

    val patch = (StylePatch.SetFill(fill1): StylePatch) |+| (StylePatch.SetFill(fill2): StylePatch)
    val result = applyPatch(style, patch)

    assertEquals(result.fill, fill2)
  }

  test("Different patch types compose without interference") {
    val style = CellStyle.default
    val font = Font("Arial", 14.0, bold = true)
    val fill = Fill.Solid(Color.Rgb(0xFFFF0000))

    val patch = (StylePatch.SetFont(font): StylePatch) |+| (StylePatch.SetFill(fill): StylePatch)
    val result = applyPatch(style, patch)

    assertEquals(result.font, font)
    assertEquals(result.fill, fill)
  }

  // ========== Extension Method Tests ==========

  test("CellStyle.applyPatch extension method works") {
    val style = CellStyle.default
    val font = Font("Arial", 14.0, bold = true)
    val patch = StylePatch.SetFont(font)

    val result = style.applyPatch(patch)
    assertEquals(result.font, font)
  }

  test("CellStyle.applyPatches extension method works with varargs") {
    val style = CellStyle.default
    val font = Font("Arial", 14.0, bold = true)
    val fill = Fill.Solid(Color.Rgb(0xFFFF0000))

    val result = style.applyPatches(
      StylePatch.SetFont(font),
      StylePatch.SetFill(fill)
    )

    assertEquals(result.font, font)
    assertEquals(result.fill, fill)
  }

  // ========== Composition Tests ==========

  test("Complex patch composition maintains all changes") {
    val style = CellStyle.default

    val font = Font("Arial", 16.0, bold = true, italic = true)
    val fill = Fill.Solid(Color.Rgb(0xFFFFFFFF))
    val border = Border.all(BorderStyle.Medium, Some(Color.Rgb(0xFF000000)))
    val numFmt = NumFmt.Currency
    val align = Align(HAlign.Right, VAlign.Middle, wrapText = true, indent = 2)

    val patch = (StylePatch.SetFont(font): StylePatch) |+|
                (StylePatch.SetFill(fill): StylePatch) |+|
                (StylePatch.SetBorder(border): StylePatch) |+|
                (StylePatch.SetNumFmt(numFmt): StylePatch) |+|
                (StylePatch.SetAlign(align): StylePatch)

    val result = applyPatch(style, patch)

    assertEquals(result.font, font)
    assertEquals(result.fill, fill)
    assertEquals(result.border, border)
    assertEquals(result.numFmt, numFmt)
    assertEquals(result.align, align)
  }

  property("Patch application preserves unchanged properties") {
    forAll { (style: CellStyle, font: Font) =>
      val patch = StylePatch.SetFont(font)
      val result = applyPatch(style, patch)

      // Font should change
      assertEquals(result.font, font)

      // Everything else should be preserved
      assertEquals(result.fill, style.fill)
      assertEquals(result.border, style.border)
      assertEquals(result.numFmt, style.numFmt)
      assertEquals(result.align, style.align)
      true
    }
  }
