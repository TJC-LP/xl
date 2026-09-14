package com.tjclp.xl.cli.raster

import java.io.IOException
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import fs2.io.process.{ProcessBuilder, Processes}

import com.tjclp.xl.cli.MemoryGuard
import com.tjclp.xl.cli.contract.CliException

/**
 * resvg integration for converting SVG to raster formats.
 *
 * resvg is a high-quality SVG rendering library written in Rust with excellent SVG 2.0 support. It
 * produces the best quality output of all the rasterizers.
 *
 * Install: cargo install resvg
 *
 * Note: resvg requires file paths (doesn't support stdin), so the SVG goes through a scratch file —
 * in `XL_SPILL_DIR` when set (the CLI's one lever for where its scratch files land, GH-517), else
 * `java.io.tmpdir`; see [[scratchSvg]].
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
            // resvg requires file paths, so the SVG goes through a scratch file (a Resource:
            // deleted even if the spawn fails); an unparseable XL_SPILL_DIR is the same usage
            // error the writes raise
            IO.fromEither(MemoryGuard.spill.left.map(CliException(_))).flatMap { spill =>
              scratchSvg(svg, spill).use(tempSvg => run(tempSvg, outputPath, dpi))
            }
        }
      case _ =>
        IO.raiseError(RasterError.FormatNotSupported(name, format))

  /**
   * The scratch SVG resvg reads, as a Resource (deleted after use): created in `spill` — the CLI's
   * `XL_SPILL_DIR` — when set, else in the JVM's `java.io.tmpdir`. A directory that does not exist
   * or refuses the file fails as [[RasterError.ScratchFileFailed]] (the CLI's `IO_WRITE`, naming
   * the directory and the lever), never a bare path under `INTERNAL`.
   */
  private[raster] def scratchSvg(svg: String, spill: Option[Path]): Resource[IO, Path] =
    Resource.make(
      IO.blocking {
        val tempSvg = spill match
          case Some(dir) => Files.createTempFile(dir, "xl-resvg-", ".svg")
          case None => Files.createTempFile("xl-resvg-", ".svg")
        Files.write(tempSvg, svg.getBytes(StandardCharsets.UTF_8))
        tempSvg
      }.adaptError { case e: IOException => RasterError.ScratchFileFailed(name, spill, e) }
    )(path => IO.blocking(Files.deleteIfExists(path)).void)

  /** `resvg --dpi <dpi> <tempSvg> <outputPath>`, its stderr the failure's message. */
  private def run(tempSvg: Path, outputPath: Path, dpi: Int): IO[Unit] =
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
            else IO.raiseError(RasterError.ConversionFailed(name, stderr, exitCode))
        yield ()
      }
