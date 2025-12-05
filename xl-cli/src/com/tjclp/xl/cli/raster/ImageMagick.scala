package com.tjclp.xl.cli.raster

import java.nio.charset.StandardCharsets
import java.nio.file.Path

import cats.effect.IO
import fs2.io.process.{ProcessBuilder, Processes}

/**
 * ImageMagick integration for converting SVG to raster formats.
 *
 * Supports both ImageMagick 7+ (`magick` command) and ImageMagick 6 (`convert` command). Falls back
 * gracefully with installation instructions if neither is found.
 */
object ImageMagick:

  /** Supported output formats */
  sealed trait Format:
    def extension: String
    def magickFormat: String
    def extraArgs: List[String]

  object Format:
    case object Png extends Format:
      val extension = "png"
      val magickFormat = "png"
      val extraArgs = List.empty

    case class Jpeg(quality: Int) extends Format:
      val extension = "jpeg"
      val magickFormat = "jpeg"
      val extraArgs = List("-quality", quality.toString)

    case object WebP extends Format:
      val extension = "webp"
      val magickFormat = "webp"
      val extraArgs = List.empty

    case object Pdf extends Format:
      val extension = "pdf"
      val magickFormat = "pdf"
      val extraArgs = List.empty

  /** Error indicating ImageMagick is not installed */
  case class ImageMagickNotFound(message: String) extends Exception(message)

  /** Error from ImageMagick conversion */
  case class ConversionError(message: String, exitCode: Int) extends Exception(message)

  // Get the Processes instance for IO
  private given Processes[IO] = Processes.forAsync[IO]

  /** Which ImageMagick command to use */
  private sealed trait ImageMagickCommand:
    def command: String
    def versionArgs: List[String]

  private object ImageMagickCommand:
    /** ImageMagick 7+ uses unified `magick` command */
    case object Magick7 extends ImageMagickCommand:
      val command = "magick"
      val versionArgs = List("--version")

    /** ImageMagick 6 uses separate `convert` command */
    case object Convert6 extends ImageMagickCommand:
      val command = "convert"
      val versionArgs = List("-version")

  /**
   * Check if a specific ImageMagick command is available.
   */
  private def isCommandAvailable(cmd: ImageMagickCommand): IO[Boolean] =
    Processes[IO]
      .spawn(ProcessBuilder(cmd.command, cmd.versionArgs))
      .use(_.exitValue)
      .map(_ == 0)
      .handleError(_ => false)

  /**
   * Find the available ImageMagick command, preferring v7 over v6.
   */
  private def findCommand: IO[Option[ImageMagickCommand]] =
    isCommandAvailable(ImageMagickCommand.Magick7).flatMap {
      case true => IO.pure(Some(ImageMagickCommand.Magick7))
      case false =>
        isCommandAvailable(ImageMagickCommand.Convert6).map {
          case true => Some(ImageMagickCommand.Convert6)
          case false => None
        }
    }

  /**
   * Check if ImageMagick is available on the system.
   *
   * Checks for both `magick` (v7+) and `convert` (v6) commands.
   */
  def isAvailable: IO[Boolean] = findCommand.map(_.isDefined)

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
    format: Format,
    dpi: Int = 144
  ): IO[Unit] =
    findCommand.flatMap {
      case None =>
        IO.raiseError(
          ImageMagickNotFound(
            s"""ImageMagick not found. Install it to enable raster export:

  macOS:   brew install imagemagick
  Ubuntu:  sudo apt-get install imagemagick librsvg2-bin
  Windows: https://imagemagick.org/script/download.php

After installation, ensure 'magick' (v7+) or 'convert' (v6) is in your PATH."""
          )
        )
      case Some(cmd) =>
        // Build command args (same for both v6 and v7)
        // magick -density <dpi> -background white svg:- <format>:<output>
        // convert -density <dpi> -background white svg:- <format>:<output>
        val args = List(
          "-density",
          dpi.toString,
          "-background",
          "white", // SVG may have transparent background
          "svg:-" // Read SVG from stdin
        ) ++ format.extraArgs ++ List(
          s"${format.magickFormat}:${outputPath.toAbsolutePath}"
        )

        Processes[IO]
          .spawn(ProcessBuilder(cmd.command, args))
          .use { process =>
            val svgBytes = svg.getBytes(StandardCharsets.UTF_8)

            // Write SVG to stdin, then read exit code and stderr
            for
              _ <- fs2.Stream.emits(svgBytes).through(process.stdin).compile.drain
              exitCode <- process.exitValue
              _ <-
                if exitCode == 0 then IO.unit
                else
                  process.stderr.through(fs2.text.utf8.decode).compile.string.flatMap { stderr =>
                    IO.raiseError(
                      ConversionError(
                        s"ImageMagick conversion failed (exit $exitCode): $stderr",
                        exitCode
                      )
                    )
                  }
            yield ()
          }
    }

  /**
   * Get the ImageMagick version string for diagnostics.
   */
  def version: IO[Option[String]] =
    findCommand.flatMap {
      case None => IO.pure(None)
      case Some(cmd) =>
        Processes[IO]
          .spawn(ProcessBuilder(cmd.command, cmd.versionArgs))
          .use { process =>
            process.stdout
              .through(fs2.text.utf8.decode)
              .compile
              .string
              .map(_.linesIterator.nextOption)
          }
          .handleError(_ => None)
    }
