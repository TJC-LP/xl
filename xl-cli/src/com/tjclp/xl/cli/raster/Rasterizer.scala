package com.tjclp.xl.cli.raster

import java.nio.file.Path

import cats.effect.IO
import cats.syntax.all.*

/**
 * Common interface for SVG-to-raster converters.
 *
 * The CLI supports multiple rasterization backends with automatic fallback:
 *   1. Batik (pure JVM, requires AWT - works in JVM mode, not native image)
 *   2. cairosvg (Python, very portable - pip install cairosvg)
 *   3. rsvg-convert (librsvg, fast - apt install librsvg2-bin)
 *   4. resvg (Rust, best quality - cargo install resvg)
 *   5. ImageMagick (last resort - may have SVG issues)
 */
trait Rasterizer:

  /** Human-readable name for this rasterizer (e.g., "Batik", "cairosvg") */
  def name: String

  /** Check if this rasterizer is available on the current system */
  def isAvailable: IO[Boolean]

  /**
   * Convert SVG content to a raster image.
   *
   * @param svg
   *   SVG content as a string
   * @param outputPath
   *   Destination file path
   * @param format
   *   Target format (PNG, JPEG, WebP, PDF)
   * @param dpi
   *   Resolution in DPI (default 144 for retina displays)
   * @return
   *   IO that completes when conversion is done, or fails with RasterError
   */
  def convertSvgToRaster(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int = 144
  ): IO[Unit]

/**
 * Output format for raster images.
 */
enum RasterFormat:
  case Png
  case Jpeg(quality: Int)
  case WebP
  case Pdf

  def extension: String = this match
    case Png => "png"
    case Jpeg(_) => "jpeg"
    case WebP => "webp"
    case Pdf => "pdf"

  def formatArg: String = this match
    case Png => "png"
    case Jpeg(_) => "jpeg"
    case WebP => "webp"
    case Pdf => "pdf"

/**
 * Errors that can occur during rasterization.
 */
sealed trait RasterError extends Exception:
  def message: String
  override def getMessage: String = message

object RasterError:
  /** No rasterizer available on the system */
  case class NoRasterizerAvailable(triedRasterizers: List[String]) extends RasterError:
    def message: String =
      s"""No SVG rasterizer available. Tried: ${triedRasterizers.mkString(", ")}
         |
         |Install one of:
         |  - cairosvg: pip install cairosvg
         |  - rsvg-convert: apt install librsvg2-bin (Debian/Ubuntu)
         |  - resvg: cargo install resvg
         |  - ImageMagick: apt install imagemagick
         |
         |Or use --format svg for vector output.""".stripMargin

  /** Specific rasterizer requested but not available */
  case class RasterizerNotFound(name: String, installHint: String) extends RasterError:
    def message: String = s"Rasterizer '$name' not found. $installHint"

  /** Format not supported by this rasterizer */
  case class FormatNotSupported(rasterizer: String, format: RasterFormat) extends RasterError:
    def message: String = s"$rasterizer does not support ${format.extension} format"

  /** Conversion failed */
  case class ConversionFailed(rasterizer: String, stderr: String, exitCode: Int)
      extends RasterError:
    def message: String =
      s"$rasterizer conversion failed (exit $exitCode): $stderr"

/**
 * Manages the fallback chain of rasterizers.
 */
object RasterizerChain:

  /**
   * Default fallback chain in priority order.
   *
   * Order rationale:
   *   - Batik: Pure JVM, no external deps, but doesn't work in native image
   *   - cairosvg: Very portable (Python), pre-installed in Claude.ai
   *   - rsvg-convert: Fast (C/Rust), common on Linux
   *   - resvg: Best SVG 2.0 support, but requires Rust/Cargo
   *   - ImageMagick: Widely installed but may have SVG rendering issues
   */
  def defaultChain: List[Rasterizer] = List(
    BatikRasterizer,
    CairoSvg,
    RsvgConvert,
    Resvg,
    ImageMagick
  )

  /** Map of rasterizer names for --rasterizer flag */
  def byName: Map[String, Rasterizer] = defaultChain.map(r => r.name.toLowerCase -> r).toMap

  /** Valid rasterizer names for CLI help */
  def validNames: List[String] = defaultChain.map(_.name.toLowerCase)

  /**
   * Convert SVG to raster using the fallback chain.
   *
   * @param svg
   *   SVG content
   * @param outputPath
   *   Destination file
   * @param format
   *   Output format
   * @param dpi
   *   Resolution
   * @param preferredRasterizer
   *   Optional name to force a specific rasterizer
   * @return
   *   IO containing the name of the rasterizer that succeeded
   */
  def convert(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int = 144,
    preferredRasterizer: Option[String] = None
  ): IO[String] =
    preferredRasterizer match
      case Some(name) =>
        // User requested specific rasterizer
        byName.get(name.toLowerCase) match
          case Some(rasterizer) =>
            rasterizer.isAvailable.flatMap {
              case true =>
                rasterizer
                  .convertSvgToRaster(svg, outputPath, format, dpi)
                  .as(rasterizer.name)
              case false =>
                IO.raiseError(
                  RasterError.RasterizerNotFound(
                    name,
                    installHintFor(name.toLowerCase)
                  )
                )
            }
          case None =>
            IO.raiseError(
              RasterError.RasterizerNotFound(
                name,
                s"Valid options: ${validNames.mkString(", ")}"
              )
            )

      case None =>
        // Try fallback chain
        tryChain(svg, outputPath, format, dpi, defaultChain, Nil)

  /**
   * Try each rasterizer in the chain until one succeeds.
   */
  private def tryChain(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int,
    remaining: List[Rasterizer],
    tried: List[String]
  ): IO[String] =
    remaining match
      case Nil =>
        IO.raiseError(RasterError.NoRasterizerAvailable(tried))

      case rasterizer :: rest =>
        rasterizer.isAvailable.flatMap {
          case false =>
            // Not available, try next
            tryChain(svg, outputPath, format, dpi, rest, tried :+ rasterizer.name)

          case true =>
            // Available, try to convert
            rasterizer
              .convertSvgToRaster(svg, outputPath, format, dpi)
              .map { _ =>
                // Log if we fell back from the default
                if tried.nonEmpty then
                  System.err.println(
                    s"Warning: ${tried.mkString(", ")} unavailable, using ${rasterizer.name}"
                  )
                rasterizer.name
              }
              .handleErrorWith { error =>
                // Conversion failed, try next (unless it's a format error)
                error match
                  case _: RasterError.FormatNotSupported =>
                    // Format not supported by this rasterizer, try next
                    tryChain(
                      svg,
                      outputPath,
                      format,
                      dpi,
                      rest,
                      tried :+ s"${rasterizer.name} (format not supported)"
                    )
                  case _ =>
                    // Other error (likely AWT/native image issue), try next
                    tryChain(
                      svg,
                      outputPath,
                      format,
                      dpi,
                      rest,
                      tried :+ s"${rasterizer.name} (${error.getMessage.take(50)})"
                    )
              }
        }

  private def installHintFor(name: String): String = name match
    case "batik" => "Batik requires JVM mode (not native image)"
    case "cairosvg" => "Install: pip install cairosvg"
    case "rsvg-convert" => "Install: apt install librsvg2-bin (Debian/Ubuntu)"
    case "resvg" => "Install: cargo install resvg"
    case "imagemagick" => "Install: apt install imagemagick"
    case _ => s"Unknown rasterizer: $name"
