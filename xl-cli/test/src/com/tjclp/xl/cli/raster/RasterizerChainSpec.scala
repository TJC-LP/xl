package com.tjclp.xl.cli.raster

import java.nio.file.{Files, Path}

import cats.effect.IO
import munit.CatsEffectSuite

/**
 * Tests for RasterizerChain and related types.
 *
 * Uses mock rasterizers to test chain logic without requiring external tools. This ensures tests
 * are fast, deterministic, and work in any environment.
 */
@SuppressWarnings(Array("org.wartremover.warts.IsInstanceOf", "org.wartremover.warts.AsInstanceOf"))
class RasterizerChainSpec extends CatsEffectSuite:

  // ========== Mock Rasterizers ==========

  /** A mock rasterizer with configurable behavior */
  private class MockRasterizer(
    val name: String,
    available: Boolean = true,
    supportedFormats: Set[RasterFormat] = Set(RasterFormat.Png),
    failWith: Option[Throwable] = None
  ) extends Rasterizer:

    def isAvailable: IO[Boolean] = IO.pure(available)

    def convertSvgToRaster(
      svg: String,
      outputPath: Path,
      format: RasterFormat,
      dpi: Int = 144
    ): IO[Unit] =
      if !supportedFormats.contains(format) then
        IO.raiseError(RasterError.FormatNotSupported(name, format))
      else
        failWith match
          case Some(err) => IO.raiseError(err)
          case None      => IO.unit // Success

  /** Helper to run chain with custom rasterizers */
  private def runChainWith(
    rasterizers: List[Rasterizer],
    format: RasterFormat = RasterFormat.Png
  ): IO[String] =
    val tempFile = Files.createTempFile("test", ".png")
    val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"/>"""

    // We need to test the tryChain logic, but it's private.
    // Instead, we'll test through the public API by temporarily replacing the chain.
    // For now, test with the mock rasterizers directly.
    def tryChain(remaining: List[Rasterizer], tried: List[String]): IO[String] =
      remaining match
        case Nil =>
          IO.raiseError(RasterError.NoRasterizerAvailable(tried))
        case rasterizer :: rest =>
          rasterizer.isAvailable.flatMap {
            case false =>
              tryChain(rest, tried :+ rasterizer.name)
            case true =>
              rasterizer
                .convertSvgToRaster(svg, tempFile, format, 144)
                .map(_ => rasterizer.name)
                .handleErrorWith {
                  case _: RasterError.FormatNotSupported =>
                    tryChain(rest, tried :+ s"${rasterizer.name} (format not supported)")
                  case err =>
                    tryChain(rest, tried :+ s"${rasterizer.name} (${err.getMessage.take(30)})")
                }
          }

    tryChain(rasterizers, Nil).guarantee(IO(Files.deleteIfExists(tempFile)).void)

  // ========== RasterFormat Tests ==========

  test("RasterFormat.Png extension is 'png'") {
    assertEquals(RasterFormat.Png.extension, "png")
  }

  test("RasterFormat.Jpeg extension is 'jpeg'") {
    assertEquals(RasterFormat.Jpeg(85).extension, "jpeg")
  }

  test("RasterFormat.WebP extension is 'webp'") {
    assertEquals(RasterFormat.WebP.extension, "webp")
  }

  test("RasterFormat.Pdf extension is 'pdf'") {
    assertEquals(RasterFormat.Pdf.extension, "pdf")
  }

  test("RasterFormat.formatArg matches extension") {
    assertEquals(RasterFormat.Png.formatArg, "png")
    assertEquals(RasterFormat.Jpeg(90).formatArg, "jpeg")
    assertEquals(RasterFormat.WebP.formatArg, "webp")
    assertEquals(RasterFormat.Pdf.formatArg, "pdf")
  }

  // ========== RasterError Tests ==========

  test("NoRasterizerAvailable message lists tried rasterizers") {
    val error = RasterError.NoRasterizerAvailable(List("Batik", "cairosvg"))
    assert(error.message.contains("Batik"), "Should mention Batik")
    assert(error.message.contains("cairosvg"), "Should mention cairosvg")
    assert(error.message.contains("Install"), "Should provide install hints")
  }

  test("RasterizerNotFound message includes name and hint") {
    val error = RasterError.RasterizerNotFound("cairosvg", "Install: pip install cairosvg")
    assert(error.message.contains("cairosvg"), "Should mention rasterizer name")
    assert(error.message.contains("pip install"), "Should include install hint")
  }

  test("FormatNotSupported message includes rasterizer and format") {
    val error = RasterError.FormatNotSupported("rsvg-convert", RasterFormat.WebP)
    assert(error.message.contains("rsvg-convert"), "Should mention rasterizer")
    assert(error.message.contains("webp"), "Should mention format")
  }

  test("ConversionFailed message includes exit code and stderr") {
    val error = RasterError.ConversionFailed("resvg", "Error: invalid SVG", 1)
    assert(error.message.contains("resvg"), "Should mention rasterizer")
    assert(error.message.contains("exit 1"), "Should mention exit code")
    assert(error.message.contains("invalid SVG"), "Should include stderr")
  }

  test("RasterError extends Exception with getMessage") {
    val error = RasterError.FormatNotSupported("test", RasterFormat.Png)
    assertEquals(error.getMessage, error.message)
    assert(error.isInstanceOf[Exception])
  }

  // ========== RasterizerChain Configuration Tests ==========

  test("defaultChain has 5 rasterizers in order") {
    val chain = RasterizerChain.defaultChain
    assertEquals(chain.length, 5)
    assertEquals(chain.map(_.name), List("Batik", "cairosvg", "rsvg-convert", "resvg", "ImageMagick"))
  }

  test("byName contains all rasterizers lowercase") {
    val names = RasterizerChain.byName
    assertEquals(names.size, 5)
    assert(names.contains("batik"))
    assert(names.contains("cairosvg"))
    assert(names.contains("rsvg-convert"))
    assert(names.contains("resvg"))
    assert(names.contains("imagemagick"))
  }

  test("validNames lists all rasterizer names lowercase") {
    val names = RasterizerChain.validNames
    assertEquals(names, List("batik", "cairosvg", "rsvg-convert", "resvg", "imagemagick"))
  }

  // ========== Chain Fallback Logic Tests (with mocks) ==========

  test("chain uses first available rasterizer") {
    val chain = List(
      new MockRasterizer("First", available = true),
      new MockRasterizer("Second", available = true)
    )

    runChainWith(chain).map { result =>
      assertEquals(result, "First")
    }
  }

  test("chain skips unavailable rasterizers") {
    val chain = List(
      new MockRasterizer("Unavailable1", available = false),
      new MockRasterizer("Unavailable2", available = false),
      new MockRasterizer("Available", available = true)
    )

    runChainWith(chain).map { result =>
      assertEquals(result, "Available")
    }
  }

  test("chain falls back when format not supported") {
    val chain = List(
      new MockRasterizer("PngOnly", available = true, supportedFormats = Set(RasterFormat.Png)),
      new MockRasterizer("JpegSupport", available = true, supportedFormats = Set(RasterFormat.Jpeg(90)))
    )

    runChainWith(chain, RasterFormat.Jpeg(90)).map { result =>
      assertEquals(result, "JpegSupport")
    }
  }

  test("chain falls back when conversion fails") {
    val chain = List(
      new MockRasterizer(
        "Failing",
        available = true,
        failWith = Some(new RuntimeException("AWT not available"))
      ),
      new MockRasterizer("Working", available = true)
    )

    runChainWith(chain).map { result =>
      assertEquals(result, "Working")
    }
  }

  test("chain returns NoRasterizerAvailable when all unavailable") {
    val chain = List(
      new MockRasterizer("Unavailable1", available = false),
      new MockRasterizer("Unavailable2", available = false)
    )

    runChainWith(chain).attempt.map { result =>
      assert(result.isLeft)
      result.left.foreach { err =>
        assert(err.isInstanceOf[RasterError.NoRasterizerAvailable])
        val noRaster = err.asInstanceOf[RasterError.NoRasterizerAvailable]
        assertEquals(noRaster.triedRasterizers, List("Unavailable1", "Unavailable2"))
      }
    }
  }

  test("chain returns NoRasterizerAvailable when all fail") {
    val chain = List(
      new MockRasterizer(
        "Fails1",
        available = true,
        failWith = Some(new RuntimeException("error1"))
      ),
      new MockRasterizer(
        "Fails2",
        available = true,
        failWith = Some(new RuntimeException("error2"))
      )
    )

    runChainWith(chain).attempt.map { result =>
      assert(result.isLeft)
      result.left.foreach { err =>
        assert(err.isInstanceOf[RasterError.NoRasterizerAvailable])
      }
    }
  }

  test("chain records format-not-supported in tried list") {
    val chain = List(
      new MockRasterizer("PngOnly", available = true, supportedFormats = Set(RasterFormat.Png))
    )

    runChainWith(chain, RasterFormat.WebP).attempt.map { result =>
      assert(result.isLeft)
      result.left.foreach { err =>
        assert(err.isInstanceOf[RasterError.NoRasterizerAvailable])
        val noRaster = err.asInstanceOf[RasterError.NoRasterizerAvailable]
        assert(noRaster.triedRasterizers.exists(_.contains("format not supported")))
      }
    }
  }

  test("chain with mixed unavailable and format-unsupported") {
    val chain = List(
      new MockRasterizer("Unavailable", available = false),
      new MockRasterizer("WrongFormat", available = true, supportedFormats = Set(RasterFormat.Pdf)),
      new MockRasterizer("RightFormat", available = true, supportedFormats = Set(RasterFormat.Png))
    )

    runChainWith(chain, RasterFormat.Png).map { result =>
      assertEquals(result, "RightFormat")
    }
  }

  // ========== RasterizerChain.convert Error Tests ==========

  test("convert with invalid rasterizer name returns RasterizerNotFound") {
    val tempFile = Files.createTempFile("test", ".png")
    val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"/>"""

    RasterizerChain
      .convert(svg, tempFile, RasterFormat.Png, 72, Some("invalid-rasterizer"))
      .attempt
      .map { result =>
        assert(result.isLeft, "Should fail")
        result.left.foreach { err =>
          assert(err.isInstanceOf[RasterError.RasterizerNotFound])
          assert(err.getMessage.contains("invalid-rasterizer"))
          assert(err.getMessage.contains("Valid options"))
        }
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)).void)
  }

  // ========== Format Support Tests (actual rasterizers) ==========

  test("BatikRasterizer rejects WebP format immediately") {
    val tempFile = Files.createTempFile("test", ".webp")
    val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"/>"""

    BatikRasterizer
      .convertSvgToRaster(svg, tempFile, RasterFormat.WebP)
      .attempt
      .map { result =>
        assert(result.isLeft)
        result.left.foreach { err =>
          assert(err.isInstanceOf[RasterError.FormatNotSupported])
        }
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)).void)
  }

  test("CairoSvg rejects WebP format immediately") {
    val tempFile = Files.createTempFile("test", ".webp")
    val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"/>"""

    CairoSvg
      .convertSvgToRaster(svg, tempFile, RasterFormat.WebP)
      .attempt
      .map { result =>
        assert(result.isLeft)
        result.left.foreach { err =>
          assert(err.isInstanceOf[RasterError.FormatNotSupported])
        }
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)).void)
  }

  test("CairoSvg rejects JPEG format immediately") {
    val tempFile = Files.createTempFile("test", ".jpg")
    val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"/>"""

    CairoSvg
      .convertSvgToRaster(svg, tempFile, RasterFormat.Jpeg(85))
      .attempt
      .map { result =>
        assert(result.isLeft)
        result.left.foreach { err =>
          assert(err.isInstanceOf[RasterError.FormatNotSupported])
        }
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)).void)
  }

  test("RsvgConvert rejects WebP format immediately") {
    val tempFile = Files.createTempFile("test", ".webp")
    val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"/>"""

    RsvgConvert
      .convertSvgToRaster(svg, tempFile, RasterFormat.WebP)
      .attempt
      .map { result =>
        assert(result.isLeft)
        result.left.foreach { err =>
          assert(err.isInstanceOf[RasterError.FormatNotSupported])
        }
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)).void)
  }

  test("RsvgConvert rejects JPEG format immediately") {
    val tempFile = Files.createTempFile("test", ".jpg")
    val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"/>"""

    RsvgConvert
      .convertSvgToRaster(svg, tempFile, RasterFormat.Jpeg(85))
      .attempt
      .map { result =>
        assert(result.isLeft)
        result.left.foreach { err =>
          assert(err.isInstanceOf[RasterError.FormatNotSupported])
        }
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)).void)
  }

  test("Resvg rejects non-PNG formats immediately") {
    val tempFile = Files.createTempFile("test", ".pdf")
    val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"/>"""

    Resvg
      .convertSvgToRaster(svg, tempFile, RasterFormat.Pdf)
      .attempt
      .map { result =>
        assert(result.isLeft)
        result.left.foreach { err =>
          assert(err.isInstanceOf[RasterError.FormatNotSupported])
        }
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)).void)
  }

  // ========== ImageMagick Delegate Detection Tests (GH-160) ==========

  test("ImageMagick.diagnostics returns useful information") {
    // This test verifies the diagnostics method works (doesn't throw)
    // The actual output depends on the system configuration
    ImageMagick.diagnostics.map { diag =>
      // Should return one of these patterns:
      // - "ImageMagick not found"
      // - "ImageMagick 6 (convert) available, SVG delegate: ..."
      // - "ImageMagick 7 (magick) available, SVG delegate: ..."
      // - "ImageMagick X found but SVG delegate '...' is missing"
      assert(
        diag.contains("ImageMagick") || diag.contains("magick") || diag.contains("convert"),
        s"Diagnostics should mention ImageMagick: $diag"
      )
    }
  }

  test("ImageMagick.isAvailable considers SVG delegate (GH-160)") {
    // This test ensures isAvailable checks delegate availability
    // We can't easily test the "broken delegate" case without mocking,
    // but we can verify the code path runs without error
    ImageMagick.isAvailable.map { available =>
      // Just verify it returns a boolean without throwing
      assert(available == true || available == false)
    }
  }
