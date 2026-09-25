package com.tjclp.xl.cli.raster

import java.nio.file.Path

import cats.effect.unsafe.implicits.global

import com.tjclp.xl.cli.contract.CliError

/**
 * The default chain as a forked JVM runs it for [[RasterSubprocessSpec]] (GH-673): the fork's PATH
 * holds the shims, which this JVM's PATH cannot (a child process is resolved against the parent's
 * PATH). The SVG is malformed on purpose, so Batik — always available on the JVM — fails and the
 * chain reaches the subprocess backends, exactly as the native binary's chain does; it is larger
 * than any pipe buffer, so a backend that exits without reading stdin fails the write.
 *
 * Prints the classified failure as two lines — the code, then the message — or `ok <backend>`.
 */
object RasterShimProbe:

  def main(args: Array[String]): Unit =
    val out = Path.of(args.headOption.getOrElse("probe.png"))
    val svg = "<svg xmlns=\"http://www.w3.org/2000/svg\">" + ("<g>" * 200000)
    RasterizerChain.convert(svg, out, RasterFormat.Png).attempt.unsafeRunSync() match
      case Right(name) => println(s"ok $name")
      case Left(t) =>
        val error = CliError.fromThrowable(t)
        println(error.code)
        println(error.message)
