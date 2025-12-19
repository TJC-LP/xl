package com.tjclp.xl.cli.raster

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.IO
import fs2.io.process.{ProcessBuilder, Processes}

/**
 * resvg integration for converting SVG to raster formats.
 *
 * resvg is a high-quality SVG rendering library written in Rust with excellent SVG 2.0 support. It
 * produces the best quality output of all the rasterizers.
 *
 * Install: cargo install resvg
 *
 * Note: resvg requires file paths (doesn't support stdin), so we use a temp file.
 */
object Resvg extends Rasterizer:

  val name: String = "resvg"

  // Get the Processes instance for IO
  private given Processes[IO] = Processes.forAsync[IO]

  /**
   * Check if resvg is available.
   */
  def isAvailable: IO[Boolean] =
    Processes[IO]
      .spawn(ProcessBuilder("resvg", List("--help")))
      .use { process =>
        for
          _ <- process.stdout.compile.drain
          _ <- process.stderr.compile.drain
          exitCode <- process.exitValue
        yield exitCode == 0
      }
      .handleError(_ => false)

  /**
   * Convert SVG to raster format using resvg.
   *
   * resvg only supports PNG output natively.
   */
  def convertSvgToRaster(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int = 144
  ): IO[Unit] =
    // resvg only supports PNG natively
    format match
      case RasterFormat.Png =>
        isAvailable.flatMap {
          case false =>
            IO.raiseError(
              RasterError.RasterizerNotFound(
                name,
                "Install: cargo install resvg"
              )
            )

          case true =>
            // resvg requires file paths, so write SVG to temp file
            IO.blocking {
              val tempSvg = Files.createTempFile("xl-resvg-", ".svg")
              Files.write(tempSvg, svg.getBytes(StandardCharsets.UTF_8))
              tempSvg
            }.flatMap { tempSvg =>
              // Build command:
              // resvg --dpi 144 input.svg output.png
              val args = List(
                "--dpi",
                dpi.toString,
                tempSvg.toAbsolutePath.toString,
                outputPath.toAbsolutePath.toString
              )

              Processes[IO]
                .spawn(ProcessBuilder("resvg", args))
                .use { process =>
                  for
                    // Drain stdout and stderr
                    _ <- process.stdout.compile.drain
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
                .guarantee(IO.blocking(Files.deleteIfExists(tempSvg)).void)
            }
        }
      case _ =>
        IO.raiseError(RasterError.FormatNotSupported(name, format))
