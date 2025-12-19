package com.tjclp.xl.cli.raster

import java.nio.charset.StandardCharsets
import java.nio.file.Path

import cats.effect.IO
import fs2.io.process.{ProcessBuilder, Processes}

/**
 * CairoSvg integration for converting SVG to raster formats.
 *
 * CairoSvg is a Python library that provides high-quality SVG rendering using Cairo. Very portable
 * since it only requires Python 3 and can be installed via pip.
 *
 * Install: pip install cairosvg
 *
 * This is the primary fallback for native image builds since it's pre-installed in Claude.ai.
 */
object CairoSvg extends Rasterizer:

  val name: String = "cairosvg"

  // Get the Processes instance for IO
  private given Processes[IO] = Processes.forAsync[IO]

  /**
   * Check if cairosvg is available.
   *
   * Tries the `cairosvg` command first, then falls back to `python3 -m cairosvg`.
   */
  def isAvailable: IO[Boolean] =
    isCommandAvailable.flatMap {
      case true => IO.pure(true)
      case false => isPythonModuleAvailable
    }

  /**
   * Check if `cairosvg` command is available.
   */
  private def isCommandAvailable: IO[Boolean] =
    Processes[IO]
      .spawn(ProcessBuilder("cairosvg", List("--version")))
      .use { process =>
        for
          _ <- process.stdout.compile.drain
          _ <- process.stderr.compile.drain
          exitCode <- process.exitValue
        yield exitCode == 0
      }
      .handleError(_ => false)

  /**
   * Check if cairosvg is available as a Python module.
   */
  private def isPythonModuleAvailable: IO[Boolean] =
    Processes[IO]
      .spawn(ProcessBuilder("python3", List("-c", "import cairosvg; print('ok')")))
      .use { process =>
        for
          _ <- process.stdout.compile.drain
          _ <- process.stderr.compile.drain
          exitCode <- process.exitValue
        yield exitCode == 0
      }
      .handleError(_ => false)

  /**
   * Determine how to invoke cairosvg.
   */
  private def findInvocation: IO[Option[(String, List[String])]] =
    isCommandAvailable.flatMap {
      case true => IO.pure(Some(("cairosvg", Nil)))
      case false =>
        isPythonModuleAvailable.map {
          case true => Some(("python3", List("-m", "cairosvg")))
          case false => None
        }
    }

  /**
   * Convert SVG to raster format using cairosvg.
   *
   * cairosvg supports PNG, PDF, PS, and SVG output. JPEG requires additional libraries.
   */
  def convertSvgToRaster(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int = 144
  ): IO[Unit] =
    // Check format support - cairosvg doesn't support WebP
    format match
      case RasterFormat.WebP =>
        IO.raiseError(RasterError.FormatNotSupported(name, format))
      case _ =>
        val formatArg = format match
          case RasterFormat.Png => "png"
          case RasterFormat.Pdf => "pdf"
          case RasterFormat.Jpeg(_) =>
            // cairosvg doesn't support JPEG directly
            // Let the chain try another rasterizer
            "png"
          case RasterFormat.WebP => "png" // unreachable

        findInvocation.flatMap {
          case None =>
            IO.raiseError(
              RasterError.RasterizerNotFound(
                name,
                "Install: pip install cairosvg"
              )
            )

          case Some((cmd, prefix)) =>
            // Build command:
            // cairosvg - -o output.png -d 144 -f png
            // or: python3 -m cairosvg - -o output.png -d 144 -f png
            val args = prefix ++ List(
              "-", // Read from stdin
              "-o",
              outputPath.toAbsolutePath.toString,
              "-d",
              dpi.toString,
              "-f",
              formatArg
            )

            Processes[IO]
              .spawn(ProcessBuilder(cmd, args))
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
