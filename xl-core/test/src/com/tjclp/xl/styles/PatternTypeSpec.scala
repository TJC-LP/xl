package com.tjclp.xl.styles

import munit.ScalaCheckSuite
import org.scalacheck.Gen
import org.scalacheck.Prop.*

import com.tjclp.xl.styles.color.Color
import com.tjclp.xl.styles.fill.{Fill, PatternType}

/**
 * GH-566: the ST_PatternType token table lives on the enum so every writer (DOM, SAX, the streaming
 * style patcher) spells it identically, and `Fill.Pattern` carries CT_PatternFill's two OPTIONAL
 * colours so a texture with an automatic foreground or background is representable.
 */
class PatternTypeSpec extends ScalaCheckSuite:

  /** ECMA-376 Part 1, §18.18.55 — pinned independently of the implementation. */
  private val ecmaTokens: Map[PatternType, String] = Map(
    PatternType.None -> "none",
    PatternType.Solid -> "solid",
    PatternType.Gray125 -> "gray125",
    PatternType.Gray0625 -> "gray0625",
    PatternType.DarkGray -> "darkGray",
    PatternType.MediumGray -> "mediumGray",
    PatternType.LightGray -> "lightGray",
    PatternType.DarkHorizontal -> "darkHorizontal",
    PatternType.DarkVertical -> "darkVertical",
    PatternType.DarkDown -> "darkDown",
    PatternType.DarkUp -> "darkUp",
    PatternType.DarkGrid -> "darkGrid",
    PatternType.DarkTrellis -> "darkTrellis",
    PatternType.LightHorizontal -> "lightHorizontal",
    PatternType.LightVertical -> "lightVertical",
    PatternType.LightDown -> "lightDown",
    PatternType.LightUp -> "lightUp",
    PatternType.LightGrid -> "lightGrid",
    PatternType.LightTrellis -> "lightTrellis"
  )

  test("all 19 ST_PatternType values are present and token is the ECMA table") {
    assertEquals(PatternType.values.toSet, ecmaTokens.keySet)
    assertEquals(PatternType.values.length, 19)
    ecmaTokens.foreach { case (p, token) => assertEquals(PatternType.token(p), token) }
  }

  property("fromToken . token = Some (every value)") {
    forAll(Gen.oneOf(PatternType.values.toIndexedSeq)) { p =>
      PatternType.fromToken(PatternType.token(p)) == Some(p)
    }
  }

  property("fromToken is case-insensitive (pre-GH-287 files carry lowercased tokens)") {
    val genCasing: Gen[String => String] = Gen.oneOf(
      (s: String) => s.toLowerCase,
      (s: String) => s.toUpperCase,
      (s: String) =>
        s.zipWithIndex.map { case (c, i) => if i % 2 == 0 then c.toUpper else c }.mkString
    )
    forAll(Gen.oneOf(PatternType.values.toIndexedSeq), genCasing) { (p, casing) =>
      PatternType.fromToken(casing(PatternType.token(p))) == Some(p)
    }
  }

  test("fromToken rejects unknown tokens") {
    assertEquals(PatternType.fromToken(""), None)
    assertEquals(PatternType.fromToken("hatched"), None)
    assertEquals(PatternType.fromToken("dark Gray"), None)
  }

  test("Fill.pattern is the both-colours convenience over the optional-colour case") {
    val fg = Color.Rgb(0xff000000)
    val bg = Color.Rgb(0xffffffff)
    assertEquals(
      Fill.pattern(fg, bg, PatternType.LightUp),
      Fill.Pattern(Some(fg), Some(bg), PatternType.LightUp)
    )
  }

  test("automatic colours are distinct fills (canonicalKey / equality tell them apart)") {
    val fg = Color.Rgb(0xff808080)
    val fgOnly = Fill.Pattern(Some(fg), None, PatternType.MediumGray)
    val both = Fill.Pattern(Some(fg), Some(Color.Rgb(0xffffffff)), PatternType.MediumGray)
    val bare = Fill.Pattern(None, None, PatternType.MediumGray)
    assertNotEquals(fgOnly, both)
    assertNotEquals(fgOnly, bare)
    assertNotEquals(bare, Fill.Pattern(None, None, PatternType.DarkGray))
    val keys = List(fgOnly, both, bare).map(f => CellStyle.default.withFill(f).canonicalKey)
    assertEquals(keys.distinct.size, 3)
  }
