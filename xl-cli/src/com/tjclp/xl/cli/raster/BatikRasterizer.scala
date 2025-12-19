package com.tjclp.xl.cli.raster

import java.io.StringReader
import java.nio.file.{Files, Path}

import scala.util.Using

import cats.effect.IO

import org.apache.batik.transcoder.{SVGAbstractTranscoder, TranscoderInput, TranscoderOutput}
import org.apache.batik.transcoder.image.{ImageTranscoder, JPEGTranscoder, PNGTranscoder}

/**
 * Pure-JVM SVG to raster image conversion using Apache Batik.
 *
 * This eliminates the external dependency on ImageMagick, providing:
 *   - Zero external dependencies for raster output
 *   - Consistent behavior across all platforms
 *   - Better error handling (no subprocess failures)
 *   - Works in headless environments
 *
 * Note: Batik requires AWT, which is not available in GraalVM native images. When running as a
 * native image, the rasterizer chain will fall back to external tools (cairosvg, rsvg-convert,
 * etc.).
 */
object BatikRasterizer extends Rasterizer:

  val name: String = "Batik"

  /** Error during rasterization */
  case class RasterizationError(message: String, cause: Option[Throwable] = None)
      extends Exception(message):
    cause.foreach(initCause)

  /**
   * Check if Batik rasterization is available.
   *
   * Batik is bundled in the classpath, but requires AWT which may not be available in native
   * images. We do a simple probe to check if AWT classes can be loaded.
   */
  def isAvailable: IO[Boolean] =
    IO.blocking {
      try
        // Try to load a basic AWT class to check availability
        Class.forName("java.awt.image.BufferedImage")
        true
      catch
        case _: UnsatisfiedLinkError => false
        case _: NoClassDefFoundError => false
        case _: ExceptionInInitializerError => false
    }.handleError(_ => false)

  /**
   * Convert SVG string to raster format and write to file.
   *
   * @param svg
   *   SVG content as string
   * @param outputPath
   *   Destination file path
   * @param format
   *   Target format (Png, Jpeg, WebP, Pdf)
   * @param dpi
   *   Resolution in DPI (default 144 for retina displays)
   */
  def convertSvgToRaster(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int = 144
  ): IO[Unit] =
    format match
      case RasterFormat.Png =>
        convertToPng(svg, outputPath, dpi)
      case RasterFormat.Jpeg(quality) =>
        convertToJpeg(svg, outputPath, dpi, quality)
      case RasterFormat.WebP =>
        // WebP not natively supported by Batik
        IO.raiseError(RasterError.FormatNotSupported(name, format))
      case RasterFormat.Pdf =>
        convertToPdf(svg, outputPath, dpi)

  /**
   * Convert SVG to PNG using Batik's PNGTranscoder.
   */
  private def convertToPng(svg: String, outputPath: Path, dpi: Int): IO[Unit] =
    IO.blocking {
      try
        val transcoder = new PNGTranscoder()

        // Set DPI (pixels per mm = dpi / 25.4)
        val pixelPerMm = dpi.toFloat / 25.4f
        transcoder.addTranscodingHint(
          SVGAbstractTranscoder.KEY_PIXEL_UNIT_TO_MILLIMETER,
          1.0f / pixelPerMm
        )

        transcode(transcoder, svg, outputPath)
      catch
        case _: UnsatisfiedLinkError =>
          throw RasterizationError(
            "PNG export requires AWT which is not available in GraalVM native image."
          )
        case _: NoClassDefFoundError =>
          throw RasterizationError(
            "PNG export requires AWT which is not available in GraalVM native image."
          )
        case _: ExceptionInInitializerError =>
          throw RasterizationError(
            "PNG export requires AWT which is not available in GraalVM native image."
          )
    }.handleErrorWith {
      case e: RasterizationError => IO.raiseError(e)
      case e =>
        IO.raiseError(RasterizationError(s"PNG conversion failed: ${e.getMessage}", Some(e)))
    }

  /**
   * Convert SVG to JPEG using Batik's JPEGTranscoder.
   */
  private def convertToJpeg(svg: String, outputPath: Path, dpi: Int, quality: Int): IO[Unit] =
    IO.blocking {
      try
        val transcoder = new JPEGTranscoder()

        // Set DPI
        val pixelPerMm = dpi.toFloat / 25.4f
        transcoder.addTranscodingHint(
          SVGAbstractTranscoder.KEY_PIXEL_UNIT_TO_MILLIMETER,
          1.0f / pixelPerMm
        )

        // Set JPEG quality (0.0 to 1.0)
        transcoder.addTranscodingHint(
          JPEGTranscoder.KEY_QUALITY,
          (quality / 100.0f).max(0.0f).min(1.0f)
        )

        transcode(transcoder, svg, outputPath)
      catch
        case _: UnsatisfiedLinkError =>
          throw RasterizationError(
            "JPEG export requires AWT which is not available in GraalVM native image."
          )
        case _: NoClassDefFoundError =>
          throw RasterizationError(
            "JPEG export requires AWT which is not available in GraalVM native image."
          )
        case _: ExceptionInInitializerError =>
          throw RasterizationError(
            "JPEG export requires AWT which is not available in GraalVM native image."
          )
    }.handleErrorWith {
      case e: RasterizationError => IO.raiseError(e)
      case e =>
        IO.raiseError(RasterizationError(s"JPEG conversion failed: ${e.getMessage}", Some(e)))
    }

  /**
   * Convert SVG to PDF using Batik's PDF transcoder.
   *
   * Note: Requires batik-extension or fop-transcoder for PDF support. Falls back to error if not
   * available.
   */
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf")) // Required for dynamic FOP loading
  private def convertToPdf(svg: String, outputPath: Path, dpi: Int): IO[Unit] =
    IO.blocking {
      // Try to load PDF transcoder dynamically
      // batik-transcoder doesn't include PDF by default
      try
        val pdfTranscoderClass = Class.forName("org.apache.fop.svg.PDFTranscoder")
        val transcoder =
          pdfTranscoderClass.getDeclaredConstructor().newInstance().asInstanceOf[ImageTranscoder]

        val pixelPerMm = dpi.toFloat / 25.4f
        transcoder.addTranscodingHint(
          SVGAbstractTranscoder.KEY_PIXEL_UNIT_TO_MILLIMETER,
          1.0f / pixelPerMm
        )

        transcode(transcoder, svg, outputPath)
      catch
        case _: ClassNotFoundException =>
          throw RasterizationError("PDF format requires Apache FOP library.")
    }.handleErrorWith {
      case e: RasterizationError => IO.raiseError(e)
      case e =>
        IO.raiseError(RasterizationError(s"PDF conversion failed: ${e.getMessage}", Some(e)))
    }

  /**
   * Common transcoding logic. Uses idiomatic Using for resource management and cleans up partial
   * files on failure.
   */
  private def transcode(transcoder: ImageTranscoder, svg: String, outputPath: Path): Unit =
    val input = new TranscoderInput(new StringReader(svg))
    Using(Files.newOutputStream(outputPath)) { outputStream =>
      val output = new TranscoderOutput(outputStream)
      transcoder.transcode(input, output)
    }.fold(
      { e =>
        Files.deleteIfExists(outputPath) // Clean up partial file on failure
        throw e
      },
      identity
    )
