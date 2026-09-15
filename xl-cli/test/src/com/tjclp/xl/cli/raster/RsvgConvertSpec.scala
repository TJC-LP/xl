package com.tjclp.xl.cli.raster

import java.nio.file.{Files, Path}

import cats.effect.unsafe.implicits.global
import munit.FunSuite

/**
 * GH-664: rsvg-convert reads the SVG from stdin when it is given no input file — every version, the
 * C-era 2.40 and the Rust rewrite alike. A bare `-` positional is only an alias for stdin in recent
 * librsvg; 2.5x (Debian's 2.54.7) opens it as a file and fails with `Error opening file …/-`, so
 * the default PNG export exited `RASTERIZER_UNAVAILABLE` on a box whose `xl rasterizers` listed
 * rsvg-convert as available. The adapter therefore passes no positional at all.
 */
class RsvgConvertSpec extends FunSuite:

  private val out = Path.of("out", "sheet.png")

  test("the argv names no input file: the SVG travels on stdin, never as a `-` positional") {
    val args = RsvgConvert.args("png", out, 144)
    assert(!args.contains("-"), s"a bare `-` is a filename to librsvg 2.5x: $args")
    // every token is either a flag or the value of the flag before it — no positional remains
    val flags = Set("--format", "--dpi-x", "--dpi-y", "-o")
    val positionals = args.zipWithIndex.collect {
      case (tok, i) if !flags.contains(tok) && !(i > 0 && flags.contains(args(i - 1))) => tok
    }
    assertEquals(positionals, Nil, s"unexpected positional in $args")
  }

  test("the argv carries the format, both DPI axes and the absolute output path") {
    val args = RsvgConvert.args("pdf", out, 300)
    assertEquals(args.take(2), List("--format", "pdf"))
    assertEquals(args.slice(2, 6), List("--dpi-x", "300", "--dpi-y", "300"))
    assertEquals(args.slice(6, 8), List("-o", out.toAbsolutePath.toString))
    assertEquals(args.length, 8)
  }

  test("end-to-end: an SVG piped on stdin renders to a PNG (skipped without rsvg-convert)") {
    val available = RsvgConvert.isAvailable.unsafeRunSync()
    assume(available, "rsvg-convert is not on PATH")
    val dir = Files.createTempDirectory("xl-rsvg-")
    val png = dir.resolve("cell.png")
    val svg =
      """<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4"><rect width="4" height="4" fill="red"/></svg>"""
    try
      RsvgConvert.convertSvgToRaster(svg, png, RasterFormat.Png, 144).unsafeRunSync()
      val bytes = Files.readAllBytes(png)
      val magic = bytes.take(8).map(_ & 0xff).toList
      assertEquals(magic, List(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a), "PNG signature")
    finally
      Files.deleteIfExists(png)
      Files.deleteIfExists(dir)
  }
