package com.tjclp.xl.cli

import munit.FunSuite

/**
 * Tests for the pure formatting core of `xl rasterizers` (GH-359).
 *
 * `Main.renderRasterizerTable` receives pre-computed availability, so these tests exercise every
 * environment (JVM, native binary, nothing installed) without subprocesses or network.
 */
class RasterizerListSpec extends FunSuite:

  private def render(
    batik: Boolean = false,
    cairo: Boolean = false,
    rsvg: Boolean = false,
    resvg: Boolean = false,
    imageMagick: Boolean = false,
    imageMagickDiag: String = "ImageMagick not found",
    nativeImage: Boolean = false
  ): (String, Boolean) =
    Main.renderRasterizerTable(
      batikAvail = batik,
      cairoAvail = cairo,
      rsvgAvail = rsvg,
      resvgAvail = resvg,
      imageMagickAvail = imageMagick,
      imageMagickDiag = imageMagickDiag,
      nativeImage = nativeImage
    )

  private def lineFor(output: String, backend: String): String =
    output.linesIterator.find(_.startsWith(backend)).getOrElse("")

  test("table lists backends in probe order, ImageMagick last (opt-in only)") {
    val (out, _) = render(batik = true)
    val positions =
      List("batik", "cairosvg", "rsvg-convert", "resvg", "imagemagick").map(out.indexOf)
    assert(positions.forall(_ >= 0), s"every backend must appear: $positions\n$out")
    assertEquals(positions, positions.sorted, s"rows must follow the probe order:\n$out")
  }

  test("hasWorking is true when any backend is available") {
    assert(render(batik = true)._2)
    assert(render(cairo = true)._2)
    assert(render(resvg = true)._2)
    assert(
      render(
        imageMagick = true,
        imageMagickDiag = "ImageMagick 7 (magick) available, SVG delegate: rsvg"
      )._2
    )
  }

  test("available Batik row shows the built-in default") {
    val (out, _) = render(batik = true)
    val row = lineFor(out, "batik")
    assert(row.contains("available"), s"batik row: $row")
    assert(row.contains("Built-in default"), s"batik row: $row")
  }

  test("Batik row explains the native binary when unavailable on native (GH-359)") {
    val (out, _) = render(nativeImage = true)
    val row = lineFor(out, "batik")
    assert(row.contains("unavailable"), s"batik row: $row")
    assert(row.toLowerCase.contains("native"), s"batik row must name the native binary: $row")
  }

  test("Batik row does not claim native binary when AWT is merely absent on a JVM") {
    val (out, _) = render(nativeImage = false)
    val row = lineFor(out, "batik")
    assert(row.contains("unavailable"), s"batik row: $row")
    assert(!row.toLowerCase.contains("native"), s"batik row must stay generic on the JVM: $row")
  }

  test("nothing installed: hasWorking=false, warning and install one-liners printed") {
    val (out, hasWorking) = render()
    assert(!hasWorking)
    assert(out.contains("WARNING"), s"warning expected:\n$out")
    assert(out.contains("pip install cairosvg"), "cairosvg one-liner")
    assert(out.contains("apt install librsvg2-bin"), "rsvg-convert one-liner")
    assert(out.contains("linebender/resvg/releases"), "resvg prebuilt release pointer")
    assert(out.contains("--rasterizer imagemagick"), "ImageMagick opt-in reminder")
  }

  test("nothing installed on native binary: warning names Batik as by-design dead (GH-359)") {
    val (out, hasWorking) = render(nativeImage = true)
    assert(!hasWorking)
    assert(out.contains("native binary"), s"native wording expected:\n$out")
    assert(out.contains("never work"), s"by-design wording expected:\n$out")
  }

  test("something installed: success footer instead of warning") {
    val (out, hasWorking) = render(batik = true)
    assert(hasWorking)
    assert(out.contains("At least one rasterizer is available"))
    assert(!out.contains("WARNING"))
  }

  test("ImageMagick with missing SVG delegate reports broken, not missing") {
    val (out, _) = render(
      imageMagick = false,
      imageMagickDiag = "ImageMagick 7 found but SVG delegate 'rsvg' is missing"
    )
    val row = lineFor(out, "imagemagick")
    assert(row.contains("broken"), s"imagemagick row: $row")
  }

  test("ImageMagick not installed reports missing") {
    val (out, _) = render()
    val row = lineFor(out, "imagemagick")
    assert(row.contains("missing"), s"imagemagick row: $row")
  }
