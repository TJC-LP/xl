package com.tjclp.xl.cli.raster

import java.nio.charset.StandardCharsets
import java.nio.file.Path

import cats.effect.IO
import fs2.io.process.{ProcessBuilder, Processes}

/**
 * rsvg-convert integration for converting SVG to raster formats.
 *
 * rsvg-convert is part of librsvg, a fast SVG rendering library written in Rust (with C bindings).
 * It's commonly available on Linux systems.
 *
 * Install:
 *   - Debian/Ubuntu: apt install librsvg2-bin
 *   - macOS: brew install librsvg
 *   - Fedora: dnf install librsvg2-tools
 */
object RsvgConvert extends Rasterizer:

  val name: String = "rsvg-convert"

  // Get the Processes instance for IO
  private given Processes[IO] = Processes.forAsync[IO]

  /**
   * Check if rsvg-convert is available.
   */
  def isAvailable: IO[Boolean] =
    Processes[IO]
      .spawn(ProcessBuilder("rsvg-convert", List("--version")))
      .use { process =>
        for
          _ <- process.stdout.compile.drain
          _ <- process.stderr.compile.drain
          exitCode <- process.exitValue
        yield exitCode == 0
      }
      .handleError(_ => false)

  /**
   * Convert SVG to raster format using rsvg-convert.
   *
   * rsvg-convert supports PNG, PDF, PS, EPS, and SVG output.
   */
  def convertSvgToRaster(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int = 144
  ): IO[Unit] =
    // Check format support - rsvg-convert doesn't support JPEG or WebP
    format match
      case RasterFormat.Jpeg(_) | RasterFormat.WebP =>
        IO.raiseError(RasterError.FormatNotSupported(name, format))
      case _ =>
        val formatArg = format match
          case RasterFormat.Png => "png"
          case RasterFormat.Pdf => "pdf"
          case RasterFormat.Jpeg(_) => "png" // unreachable: rejected above
          case RasterFormat.WebP => "png" // unreachable: rejected above

        isAvailable.flatMap {
          case false =>
            IO.raiseError(
              RasterError.RasterizerNotFound(
                name,
                "Install: apt install librsvg2-bin (Debian/Ubuntu)"
              )
            )

          case true =>
            // Build command:
            // rsvg-convert --format png --dpi-x 144 --dpi-y 144 -o output.png -
            val args = List(
              "--format",
              formatArg,
              "--dpi-x",
              dpi.toString,
              "--dpi-y",
              dpi.toString,
              "-o",
              outputPath.toAbsolutePath.toString,
              "-" // Read from stdin
            )

            Processes[IO]
              .spawn(ProcessBuilder("rsvg-convert", args))
              .use { process =>
                val svgBytes = svg.getBytes(StandardCharsets.UTF_8)

                for
                  // Write SVG to stdin
                  _ <- fs2.Stream.emits(svgBytes).through(process.stdin).compile.drain
                  // Always drain stderr to prevent hanging
                  stderr <- process.stderr.through(fs2.text.utf8.decode).compile.string
                  exitCode <- process.exitValue
                  _ <-
                    if exitCode == 0 then IO.unit
                    else
                      IO.raiseError(
                        RasterError.ConversionFailed(name, stderr, exitCode)
                      )
                yield ()
              }
        }
