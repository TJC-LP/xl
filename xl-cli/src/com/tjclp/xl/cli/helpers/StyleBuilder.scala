package com.tjclp.xl.cli.helpers

import scala.util.chaining.*

import cats.effect.IO
import cats.syntax.traverse.*
import com.tjclp.xl.cli.ColorParser
import com.tjclp.xl.styles.alignment.{Align, HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderStyle}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/**
 * Style building utilities for CLI commands.
 *
 * Provides helpers for building CellStyle from CLI options and parsing style-related strings.
 */
object StyleBuilder:

  /**
   * Build a CellStyle from CLI options.
   *
   * @return
   *   IO containing CellStyle, or error if parsing fails
   */
  def buildCellStyle(
    bold: Boolean,
    italic: Boolean,
    underline: Boolean,
    bg: Option[String],
    fg: Option[String],
    fontSize: Option[Double],
    fontName: Option[String],
    align: Option[String],
    valign: Option[String],
    wrap: Boolean,
    numFormat: Option[String],
    border: Option[String],
    borderColor: Option[String]
  ): IO[CellStyle] =
    for
      bgColor <- bg.traverse(s => IO.fromEither(ColorParser.parse(s).left.map(new Exception(_))))
      fgColor <- fg.traverse(s => IO.fromEither(ColorParser.parse(s).left.map(new Exception(_))))
      bdrColor <- borderColor.traverse(s =>
        IO.fromEither(ColorParser.parse(s).left.map(new Exception(_)))
      )
      hAlign <- align.traverse(s => IO.fromEither(parseHAlign(s).left.map(new Exception(_))))
      vAlign <- valign.traverse(s => IO.fromEither(parseVAlign(s).left.map(new Exception(_))))
      bdrStyle <- border.traverse(s =>
        IO.fromEither(parseBorderStyle(s).left.map(new Exception(_)))
      )
      nFmt <- numFormat.traverse(s => IO.fromEither(parseNumFmt(s).left.map(new Exception(_))))
    yield
      val font = Font.default
        .withBold(bold)
        .withItalic(italic)
        .withUnderline(underline)
        .pipe(f => fgColor.fold(f)(c => f.withColor(c)))
        .pipe(f => fontSize.fold(f)(s => f.withSize(s)))
        .pipe(f => fontName.fold(f)(n => f.withName(n)))

      val fill = bgColor.map(Fill.Solid.apply).getOrElse(Fill.None)

      val cellBorder = bdrStyle
        .map(style => Border.all(style, bdrColor))
        .getOrElse(Border.none)

      val alignment = Align.default
        .pipe(a => hAlign.fold(a)(h => a.withHAlign(h)))
        .pipe(a => vAlign.fold(a)(v => a.withVAlign(v)))
        .pipe(a => if wrap then a.withWrap() else a)

      CellStyle(
        font = font,
        fill = fill,
        border = cellBorder,
        numFmt = nFmt.getOrElse(NumFmt.General),
        align = alignment
      )

  /**
   * Build a description list of style options for output.
   */
  def buildStyleDescription(
    bold: Boolean,
    italic: Boolean,
    underline: Boolean,
    bg: Option[String],
    fg: Option[String],
    fontSize: Option[Double],
    fontName: Option[String],
    align: Option[String],
    valign: Option[String],
    wrap: Boolean,
    numFormat: Option[String],
    border: Option[String]
  ): List[String] =
    List(
      if bold then Some("bold") else None,
      if italic then Some("italic") else None,
      if underline then Some("underline") else None,
      bg.map(c => s"bg=$c"),
      fg.map(c => s"fg=$c"),
      fontSize.map(s => s"font-size=$s"),
      fontName.map(n => s"font-name=$n"),
      align.map(a => s"align=$a"),
      valign.map(v => s"valign=$v"),
      if wrap then Some("wrap") else None,
      numFormat.map(f => s"format=$f"),
      border.map(b => s"border=$b")
    ).flatten

  /**
   * Parse horizontal alignment string.
   */
  def parseHAlign(s: String): Either[String, HAlign] =
    s.toLowerCase match
      case "left" => Right(HAlign.Left)
      case "center" => Right(HAlign.Center)
      case "right" => Right(HAlign.Right)
      case "justify" => Right(HAlign.Justify)
      case "general" => Right(HAlign.General)
      case other => Left(s"Unknown horizontal alignment: $other. Use left, center, right, justify")

  /**
   * Parse vertical alignment string.
   */
  def parseVAlign(s: String): Either[String, VAlign] =
    s.toLowerCase match
      case "top" => Right(VAlign.Top)
      case "middle" | "center" => Right(VAlign.Middle)
      case "bottom" => Right(VAlign.Bottom)
      case other => Left(s"Unknown vertical alignment: $other. Use top, middle, bottom")

  /**
   * Parse border style string.
   */
  def parseBorderStyle(s: String): Either[String, BorderStyle] =
    s.toLowerCase match
      case "none" => Right(BorderStyle.None)
      case "thin" => Right(BorderStyle.Thin)
      case "medium" => Right(BorderStyle.Medium)
      case "thick" => Right(BorderStyle.Thick)
      case "dashed" => Right(BorderStyle.Dashed)
      case "dotted" => Right(BorderStyle.Dotted)
      case "double" => Right(BorderStyle.Double)
      case other => Left(s"Unknown border style: $other. Use none, thin, medium, thick")

  /**
   * Parse number format string.
   */
  def parseNumFmt(s: String): Either[String, NumFmt] =
    s.toLowerCase match
      case "general" => Right(NumFmt.General)
      case "number" => Right(NumFmt.Decimal)
      case "currency" => Right(NumFmt.Currency)
      case "percent" => Right(NumFmt.Percent)
      case "date" => Right(NumFmt.Date)
      case "text" => Right(NumFmt.Text)
      case other =>
        Left(s"Unknown number format: $other. Use general, number, currency, percent, date, text")

  /**
   * Merge two CellStyles, applying non-default values from newStyle to existingStyle.
   *
   * The merge logic applies each component from newStyle only if it differs from the default,
   * otherwise preserves the value from existingStyle.
   */
  def mergeStyles(existingStyle: CellStyle, newStyle: CellStyle): CellStyle =
    val defaultStyle = CellStyle.default

    // Merge fonts: apply non-default properties from newStyle
    val mergedFont = {
      val existing = existingStyle.font
      val newer = newStyle.font
      val default = Font.default
      existing
        .withBold(if newer.bold != default.bold then newer.bold else existing.bold)
        .withItalic(if newer.italic != default.italic then newer.italic else existing.italic)
        .withUnderline(
          if newer.underline != default.underline then newer.underline else existing.underline
        )
        .pipe(f =>
          if newer.color != default.color then newer.color.map(f.withColor).getOrElse(f) else f
        )
        .pipe(f => if newer.sizePt != default.sizePt then f.withSize(newer.sizePt) else f)
        .pipe(f => if newer.name != default.name then f.withName(newer.name) else f)
    }

    // Merge fill: use newStyle if not None
    val mergedFill =
      if newStyle.fill != Fill.None then newStyle.fill else existingStyle.fill

    // Merge border: use newStyle if not none
    val mergedBorder =
      if newStyle.border != Border.none then newStyle.border else existingStyle.border

    // Merge numFmt: use newStyle if not General
    val mergedNumFmt =
      if newStyle.numFmt != NumFmt.General then newStyle.numFmt else existingStyle.numFmt

    // Merge alignment: apply non-default properties
    val mergedAlign = {
      val existing = existingStyle.align
      val newer = newStyle.align
      val default = Align.default
      existing
        .pipe(a =>
          if newer.horizontal != default.horizontal then a.withHAlign(newer.horizontal) else a
        )
        .pipe(a => if newer.vertical != default.vertical then a.withVAlign(newer.vertical) else a)
        .pipe(a => if newer.wrapText != default.wrapText then a.withWrap(newer.wrapText) else a)
    }

    CellStyle(
      font = mergedFont,
      fill = mergedFill,
      border = mergedBorder,
      numFmt = mergedNumFmt,
      align = mergedAlign
    )
