package com.tjclp.xl.cli.raster

import java.nio.file.{Files, Path}
import javax.imageio.ImageIO

import cats.effect.IO

import com.tjclp.xl.{*, given}

import munit.CatsEffectSuite

/** Compare the actual raster with an unclipped oracle, including both ends of the number. */
class TightTextRasterSpec extends CatsEffectSuite:
  private def check(width: Double, styled: Boolean = false): IO[Unit] =
    // Raw numeric cells have no style and draw at 15px. The Double codec attaches a style,
    // which draws at 14px; using that codec alone would miss the original failure.
    val base =
      if styled then Sheet("GH-505").put(ref"B1" -> 1234567.9)
      else Sheet("GH-505").put(ref"B1", CellValue.Number(BigDecimal("1234567.9")))
    val sheet = base
      .setColumnProperties(Column.from0(1), ColumnProperties(width = Some(width)))
    // Empty neighbours leave room for the oracle's glyphs outside B1's clipping rectangle.
    val clipped = sheet.toSvg(ref"A1:C1")
    val unclipped = clipped.replace(" clip-path=\"url(#clip-B1)\"", "")
    val directory = sys.env.get("XL_505_EVIDENCE_DIR") match
      case Some(path) => Files.createDirectories(Path.of(path))
      case None => Files.createTempDirectory("xl-505-")
    val prefix = s"${if styled then "styled-" else ""}width-$width"
    val clippedPath = directory.resolve(s"$prefix-clipped.png")
    val oraclePath = directory.resolve(s"$prefix-unclipped.png")
    val svgPath = directory.resolve(s"$prefix.svg")
    val run = for
      _ <- IO(Files.writeString(svgPath, clipped))
      _ <- BatikRasterizer.convertSvgToRaster(clipped, clippedPath, RasterFormat.Png, dpi = 96)
      _ <- BatikRasterizer.convertSvgToRaster(unclipped, oraclePath, RasterFormat.Png, dpi = 96)
      _ <- IO {
        val actual = Option(ImageIO.read(clippedPath.toFile)).getOrElse(fail("No clipped PNG"))
        val oracle = Option(ImageIO.read(oraclePath.toFile)).getOrElse(fail("No oracle PNG"))
        val differences = for
          y <- 0 until actual.getHeight
          x <- 0 until actual.getWidth
          if actual.getRGB(x, y) != oracle.getRGB(x, y)
        yield (x, y)
        assertEquals(differences.size, 0, s"GH-505: width $width clips glyph pixels: $differences")
        assert(clipped.contains("1234567.9"), "The number must fit without a hash marker")
      }
    yield ()
    if sys.env.contains("XL_505_EVIDENCE_DIR") then run
    else
      run.guarantee(IO {
        List(clippedPath, oraclePath, svgPath).foreach(Files.deleteIfExists)
        Files.deleteIfExists(directory)
        ()
      })

  test("GH-505: Batik preserves both ends at width 9.0 (one pixel slack)") { check(9.0) }
  test("GH-505: width 9.25 control preserves both ends") { check(9.25) }
  test("GH-505: wide column control preserves both ends") { check(12.0) }
  test("GH-505: codec-styled width 9.0 control preserves both ends") { check(9.0, styled = true) }
