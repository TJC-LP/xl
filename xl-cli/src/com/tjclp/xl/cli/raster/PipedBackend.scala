package com.tjclp.xl.cli.raster

import java.io.IOException

import cats.effect.IO
import cats.syntax.all.*
import fs2.Chunk
import fs2.io.process.{ProcessBuilder, Processes}

/**
 * A rasterizer backend run as a child process that reads the SVG on stdin (rsvg-convert, cairosvg,
 * ImageMagick).
 *
 * The stdin write and the stdout and stderr drains run concurrently (GH-673): written in sequence,
 * a backend that exits before consuming its input — a bad install, a crash, a refused option — made
 * the write raise `IOException` ("Stream closed", "Broken pipe"), which escaped as `INTERNAL` and
 * never showed the backend's stderr, the one thing the user can act on. Concurrent drains also keep
 * a chatty backend from filling a pipe while the write waits on it.
 */
private[raster] object PipedBackend:

  private given Processes[IO] = Processes.forAsync[IO]

  /**
   * Run `command args` with `input` on stdin. A non-zero exit is `ConversionFailed(name, stderr,
   * exit)` whether or not the backend read its input; a zero exit that left the input unread is a
   * failure too, since the output cannot be trusted.
   */
  def run(name: String, command: String, args: List[String], input: Array[Byte]): IO[Unit] =
    Processes[IO].spawn(ProcessBuilder(command, args)).use { process =>
      val write = fs2.Stream
        .chunk(Chunk.array(input))
        .through(process.stdin)
        .compile
        .drain
        .as(Option.empty[IOException])
        .recover { case closed: IOException => Some(closed) }
      val stderr = process.stderr.through(fs2.text.utf8.decode).compile.string
      val stdout = process.stdout.compile.drain
      (write, stderr, stdout).parTupled.flatMap { (unread, err, _) =>
        process.exitValue.flatMap { exit =>
          (exit, unread) match
            case (0, None) => IO.unit
            case (0, Some(closed)) =>
              val reason = Option(closed.getMessage).getOrElse(closed.toString)
              val detail = if err.isBlank then "" else s"; $err"
              IO.raiseError(
                RasterError.ConversionFailed(
                  name,
                  s"exited before reading the whole SVG ($reason)$detail",
                  exit
                )
              )
            case _ => IO.raiseError(RasterError.ConversionFailed(name, err, exit))
        }
      }
    }
