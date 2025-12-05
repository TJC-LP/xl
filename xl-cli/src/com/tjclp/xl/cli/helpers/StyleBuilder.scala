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
