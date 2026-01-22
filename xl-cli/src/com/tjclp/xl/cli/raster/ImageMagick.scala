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
 *
 * Note: ImageMagick's SVG rendering can be inconsistent. Consider using cairosvg or rsvg-convert
 * for better SVG compatibility.
 */
object ImageMagick extends Rasterizer:

  val name: String = "ImageMagick"

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
      .use { process =>
        for
          _ <- process.stdout.compile.drain
          _ <- process.stderr.compile.drain
          exitCode <- process.exitValue
        yield exitCode == 0
      }
      .handleError(_ => false)

  /**
   * Check if a binary is available in PATH.
   */
  private def isBinaryInPath(binary: String): IO[Boolean] =
    Processes[IO]
      .spawn(ProcessBuilder("which", List(binary)))
      .use { process =>
        for
          _ <- process.stdout.compile.drain
          _ <- process.stderr.compile.drain
          exitCode <- process.exitValue
        yield exitCode == 0
      }
      .handleError(_ => false)

  /**
   * Get the SVG delegate configuration from ImageMagick.
   *
   * Runs `<cmd> -list delegate` and parses the SVG entry to find the delegate binary. Returns the
   * binary name (e.g., "rsvg-convert") if configured, or None if SVG conversion uses internal
   * rendering.
   */
  private def getSvgDelegate(cmd: ImageMagickCommand): IO[Option[String]] =
    Processes[IO]
      .spawn(ProcessBuilder(cmd.command, List("-list", "delegate")))
      .use { process =>
        for
          output <- process.stdout.through(fs2.text.utf8.decode).compile.string
          _ <- process.stderr.compile.drain
          exitCode <- process.exitValue
        yield
          if exitCode != 0 then None
          else
            // Parse delegate list for SVG entry
            // Format: "        svg =>          "rsvg-convert' --dpi-x %x ..."
            // Note: Leading whitespace varies, and we need to match "svg =>" specifically
            output.linesIterator
              .find(line => line.trim.startsWith("svg") && line.contains("=>"))
              .flatMap { line =>
                // Extract binary name from delegate command
                // The delegate line format is: [whitespace]svg => "binary' args..."
                val delegatePattern = """\bsvg\s*=>\s*"([^']+)'""".r
                delegatePattern.findFirstMatchIn(line).map(_.group(1))
              }
      }
      .handleError(_ => None)

  /**
   * Check if ImageMagick's SVG delegate is functional.
   *
   * ImageMagick may be installed but its SVG delegate (e.g., rsvg-convert) may be missing. This
   * causes confusing errors like "delegate failed 'rsvg-convert' -o '%o' '%i'". We detect this
   * upfront and skip ImageMagick if the delegate is broken.
   */
  private def isSvgDelegateFunctional(cmd: ImageMagickCommand): IO[Boolean] =
    getSvgDelegate(cmd).flatMap {
      case None =>
        // No SVG delegate configured - ImageMagick may use internal rendering
        // This is fine, let it try
        IO.pure(true)
      case Some(delegateBinary) =>
        // Delegate is configured - verify the binary exists
        isBinaryInPath(delegateBinary)
    }

  /**
   * Find the available ImageMagick command, preferring v7 over v6.
   *
   * Also verifies that the SVG delegate is functional (GH-160). If v7 is available but its delegate
   * is broken, falls back to try v6.
   */
  private def findCommand: IO[Option[ImageMagickCommand]] =
    // Helper to check v6 availability and delegate
    def tryV6: IO[Option[ImageMagickCommand]] =
      isCommandAvailable(ImageMagickCommand.Convert6).flatMap {
        case true =>
          isSvgDelegateFunctional(ImageMagickCommand.Convert6).map {
            case true => Some(ImageMagickCommand.Convert6)
            case false => None
          }
        case false => IO.pure(None)
      }

    isCommandAvailable(ImageMagickCommand.Magick7).flatMap {
      case true =>
        // v7 available, check SVG delegate
        isSvgDelegateFunctional(ImageMagickCommand.Magick7).flatMap {
          case true => IO.pure(Some(ImageMagickCommand.Magick7))
          case false => tryV6 // v7 delegate broken, fall back to v6
        }
      case false => tryV6
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
    format: RasterFormat,
    dpi: Int = 144
  ): IO[Unit] =
    findCommand.flatMap {
      case None =>
        IO.raiseError(
          RasterError.RasterizerNotFound(
            name,
            """Install ImageMagick:
              |  macOS:   brew install imagemagick
              |  Ubuntu:  sudo apt-get install imagemagick
              |  Windows: https://imagemagick.org/script/download.php""".stripMargin
          )
        )
      case Some(cmd) =>
        // Map format to ImageMagick format string and extra args
        val (magickFormat, extraArgs) = format match
          case RasterFormat.Png => ("png", List.empty)
          case RasterFormat.Jpeg(q) => ("jpeg", List("-quality", q.toString))
          case RasterFormat.WebP => ("webp", List.empty)
          case RasterFormat.Pdf => ("pdf", List.empty)

        // Build command args (same for both v6 and v7)
        // magick -density <dpi> -background white svg:- <format>:<output>
        // convert -density <dpi> -background white svg:- <format>:<output>
        val args = List(
          "-density",
          dpi.toString,
          "-background",
          "white", // SVG may have transparent background
          "svg:-" // Read SVG from stdin
        ) ++ extraArgs ++ List(
          s"$magickFormat:${outputPath.toAbsolutePath}"
        )

        Processes[IO]
          .spawn(ProcessBuilder(cmd.command, args))
          .use { process =>
            val svgBytes = svg.getBytes(StandardCharsets.UTF_8)

            // Write SVG to stdin, then read exit code and stderr
            for
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
            for
              versionOutput <- process.stdout.through(fs2.text.utf8.decode).compile.string
              _ <- process.stderr.compile.drain
            yield versionOutput.linesIterator.nextOption
          }
          .handleError(_ => None)
    }

  /**
   * Get diagnostic information about ImageMagick availability.
   *
   * Returns a human-readable string explaining the state of ImageMagick and its SVG delegate.
   * Useful for debugging rasterization issues.
   */
  def diagnostics: IO[String] =
    for
      v7Available <- isCommandAvailable(ImageMagickCommand.Magick7)
      v6Available <- isCommandAvailable(ImageMagickCommand.Convert6)
      result <-
        if v7Available then
          for
            delegate <- getSvgDelegate(ImageMagickCommand.Magick7)
            delegateOk <- isSvgDelegateFunctional(ImageMagickCommand.Magick7)
          yield
            if delegateOk then
              s"ImageMagick 7 (magick) available, SVG delegate: ${delegate.getOrElse("internal")}"
            else
              s"ImageMagick 7 (magick) found but SVG delegate '${delegate.getOrElse("unknown")}' is missing"
        else if v6Available then
          for
            delegate <- getSvgDelegate(ImageMagickCommand.Convert6)
            delegateOk <- isSvgDelegateFunctional(ImageMagickCommand.Convert6)
          yield
            if delegateOk then
              s"ImageMagick 6 (convert) available, SVG delegate: ${delegate.getOrElse("internal")}"
            else
              s"ImageMagick 6 (convert) found but SVG delegate '${delegate.getOrElse("unknown")}' is missing"
        else IO.pure("ImageMagick not found")
    yield result
