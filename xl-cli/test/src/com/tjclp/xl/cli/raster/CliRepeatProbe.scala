package com.tjclp.xl.cli.raster

import cats.effect.unsafe.implicits.global

import com.tjclp.xl.cli.{Cli, CliIO}

/**
 * Runs one `xl` command line several times in this JVM, for [[RasterSubprocessSpec]] (GH-877): `xl
 * <args>` repeated `times` times, as `times <args…>`. Each attempt prints a header line,
 * [[RasterSubprocessSpec]]'s `@@ attempt <n> exit <code>`, on stdout, then the stderr that attempt
 * produced.
 *
 * The ImageMagick race case needs an interpreted JVM (`-Xint`) and wants more than one try at the
 * race. Starting an interpreted JVM costs most of an attempt, so one child repeats the run instead
 * of three children each paying for the startup. A failed ImageMagick probe is not cached, so every
 * attempt runs the probe conversion, and the race, again.
 */
object CliRepeatProbe:

  def main(args: Array[String]): Unit =
    val times = args.headOption.flatMap(_.toIntOption).getOrElse(1)
    val cli = args.toList.drop(1)
    (1 to times).foreach { attempt =>
      val (io, collected) = CliIO.capturing("").unsafeRunSync()
      val exit = Cli.run(cli, io).unsafeRunSync()
      val (_, stderr) = collected.unsafeRunSync()
      println(s"@@ attempt $attempt exit ${exit.code}")
      print(stderr)
    }
