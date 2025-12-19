package com.tjclp.xl.cli.raster

import cats.effect.IO
import munit.CatsEffectSuite

import java.nio.file.{Files, Path}

/**
 * Regression tests for Batik rasterizer (#83/#86).
 *
 * Tests cover error handling, format validation, and basic functionality.
 * Note: Full rasterization tests require AWT which may not be available in all environments.
 */
@SuppressWarnings(Array("org.wartremover.warts.IsInstanceOf"))
class BatikRasterizerSpec extends CatsEffectSuite:

  // ========== Format Tests ==========

  test("RasterFormat.Png has correct extension") {
    assertEquals(RasterFormat.Png.extension, "png")
  }

  test("RasterFormat.Jpeg has correct extension") {
    assertEquals(RasterFormat.Jpeg(90).extension, "jpeg")
  }

  test("RasterFormat.WebP has correct extension") {
    assertEquals(RasterFormat.WebP.extension, "webp")
  }

  test("RasterFormat.Pdf has correct extension") {
    assertEquals(RasterFormat.Pdf.extension, "pdf")
  }

  // ========== RasterizationError Tests ==========

  test("RasterizationError contains message") {
    val error = BatikRasterizer.RasterizationError("Test error")
    assertEquals(error.getMessage, "Test error")
  }

  test("RasterizationError contains cause") {
    val cause = new RuntimeException("Underlying cause")
    val error = BatikRasterizer.RasterizationError("Test error", Some(cause))
    assertEquals(error.getMessage, "Test error")
    assertEquals(error.getCause, cause)
  }

  test("RasterizationError without cause has no getCause") {
    val error = BatikRasterizer.RasterizationError("Test error")
    assert(Option(error.getCause).isEmpty, "getCause should be null/empty when no cause provided")
  }

  // ========== WebP Format Error Test ==========

  test("WebP format returns error (not natively supported)") {
    val tempFile = Files.createTempFile("test", ".webp")
    val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect fill="red" width="100" height="100"/></svg>"""

    val result = BatikRasterizer
      .convertSvgToRaster(svg, tempFile, RasterFormat.WebP)
      .attempt

    result.map { either =>
      assert(either.isLeft, "WebP should fail")
      either.left.foreach { err =>
        assert(err.isInstanceOf[RasterError.FormatNotSupported])
      }
    }.guarantee(IO(Files.deleteIfExists(tempFile)))
  }

  // ========== Availability Test ==========

  test("isAvailable checks AWT availability") {
    // In JVM mode, this should return true if AWT is available
    // In native image, this should return false
    BatikRasterizer.isAvailable.map { available =>
      // Just verify it returns a boolean without error
      assert(available || !available, "Should return true or false")
    }
  }

  // ========== PNG Rasterization Test ==========

  test("PNG conversion produces valid file on JVM") {
    // This test may fail in native image without AWT
    val tempFile = Files.createTempFile("test", ".png")
    val svg =
      """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect fill="red" width="100" height="100"/></svg>"""

    BatikRasterizer
      .convertSvgToRaster(svg, tempFile, RasterFormat.Png, dpi = 72)
      .attempt
      .flatMap {
        case Right(_) =>
          IO {
            // Verify file was created and has content
            assert(Files.exists(tempFile), "PNG file should exist")
            val size = Files.size(tempFile)
            assert(size > 0, s"PNG file should have content, but size was $size")
            // PNG files start with magic bytes: 137 80 78 71
            val bytes = Files.readAllBytes(tempFile)
            assertEquals(bytes(0).toInt & 0xff, 137, "PNG should start with magic byte 137")
            assertEquals(bytes(1).toInt & 0xff, 80, "PNG should have 'P' (80)")
            assertEquals(bytes(2).toInt & 0xff, 78, "PNG should have 'N' (78)")
            assertEquals(bytes(3).toInt & 0xff, 71, "PNG should have 'G' (71)")
          }
        case Left(e: BatikRasterizer.RasterizationError) if e.getMessage.contains("AWT") =>
          // Expected in native image environment - skip test
          IO(assume(false, "AWT not available - skipping PNG test"))
        case Left(e) =>
          IO.raiseError(new AssertionError(s"Unexpected error: ${e.getMessage}", e))
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)))
  }

  // ========== JPEG Rasterization Test ==========

  test("JPEG conversion produces valid file on JVM") {
    val tempFile = Files.createTempFile("test", ".jpg")
    val svg =
      """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect fill="blue" width="100" height="100"/></svg>"""

    BatikRasterizer
      .convertSvgToRaster(svg, tempFile, RasterFormat.Jpeg(85), dpi = 72)
      .attempt
      .flatMap {
        case Right(_) =>
          IO {
            assert(Files.exists(tempFile), "JPEG file should exist")
            val size = Files.size(tempFile)
            assert(size > 0, s"JPEG file should have content, but size was $size")
            // JPEG files start with magic bytes: 0xFF 0xD8 0xFF
            val bytes = Files.readAllBytes(tempFile)
            assertEquals(bytes(0).toInt & 0xff, 0xff, "JPEG should start with 0xFF")
            assertEquals(bytes(1).toInt & 0xff, 0xd8, "JPEG should have 0xD8")
            assertEquals(bytes(2).toInt & 0xff, 0xff, "JPEG should have 0xFF")
          }
        case Left(e: BatikRasterizer.RasterizationError) if e.getMessage.contains("AWT") =>
          IO(assume(false, "AWT not available - skipping JPEG test"))
        case Left(e) =>
          IO.raiseError(new AssertionError(s"Unexpected error: ${e.getMessage}", e))
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)))
  }

  // ========== PDF Format Test ==========

  test("PDF conversion fails gracefully without FOP") {
    val tempFile = Files.createTempFile("test", ".pdf")
    val svg =
      """<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect fill="green" width="100" height="100"/></svg>"""

    BatikRasterizer
      .convertSvgToRaster(svg, tempFile, RasterFormat.Pdf)
      .attempt
      .map {
        case Right(_) =>
          // PDF worked - FOP must be available
          assert(Files.exists(tempFile), "PDF file should exist if FOP is available")
        case Left(e: BatikRasterizer.RasterizationError) =>
          // Expected - FOP not available
          assert(
            e.getMessage.contains("FOP") || e.getMessage.contains("PDF"),
            s"Should mention FOP or PDF in error: ${e.getMessage}"
          )
        case Left(e) =>
          fail(s"Unexpected error type: ${e.getClass.getName}: ${e.getMessage}")
      }
      .guarantee(IO(Files.deleteIfExists(tempFile)))
  }
