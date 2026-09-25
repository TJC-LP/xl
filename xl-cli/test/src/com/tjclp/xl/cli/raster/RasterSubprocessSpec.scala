package com.tjclp.xl.cli.raster

import java.io.File
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.nio.file.attribute.PosixFilePermissions
import java.util.concurrent.TimeUnit

import scala.jdk.CollectionConverters.*

import cats.effect.unsafe.implicits.global
import munit.FunSuite

import com.tjclp.xl.cli.contract.{CliError, ErrorCode, TestFixtures}

/**
 * GH-673: the subprocess backends wrote all of stdin, then drained stderr, then read the exit code.
 * A backend that exits before consuming stdin made the write raise `IOException: Stream closed`,
 * which surfaced as `INTERNAL` with an fs2 stack trace under a forced `--rasterizer` (and as a bare
 * "Stream closed" in the default chain's tried list) — the backend's stderr, the one thing the user
 * can act on, was never shown.
 *
 * The shims are PATH scripts that answer the availability probe (`--version`, `--help`) and exit 1
 * with a message on stderr for a real conversion, without reading stdin. A child process is
 * resolved against its parent's PATH, so the end-to-end cases fork a JVM whose PATH starts with the
 * shim directory; the SVGs sent are larger than any pipe buffer, so the write cannot complete into
 * the buffer before the shim exits.
 */
class RasterSubprocessSpec extends FunSuite:

  private val shimMessage = "shim: refusing to render"

  private val shimScript =
    s"""#!/bin/sh
       |case "$$1" in
       |  --version|--help) echo "shim 1.0"; exit 0 ;;
       |esac
       |echo "$shimMessage" >&2
       |exit 1
       |""".stripMargin

  /** A directory holding a failing shim for each stdin-fed backend (and resvg, which is not). */
  private val shims = FunFixture[Path](
    setup = _ =>
      val dir = Files.createTempDirectory("xl-raster-shims-")
      Vector("rsvg-convert", "cairosvg", "resvg").foreach { name =>
        val shim = dir.resolve(name)
        Files.writeString(shim, shimScript, StandardCharsets.UTF_8)
        Files.setPosixFilePermissions(shim, PosixFilePermissions.fromString("rwxr-xr-x"))
      }
      dir
    ,
    teardown = dir =>
      val entries = Files.list(dir)
      try entries.forEach(p => Files.deleteIfExists(p))
      finally entries.close()
      Files.deleteIfExists(dir)
  )

  /** Everything this JVM loaded its classes from: the fork runs the same code. */
  private def classpath: String =
    def urls(loader: Option[ClassLoader]): List[String] = loader match
      case Some(u: java.net.URLClassLoader) =>
        u.getURLs.toList.map(url => Path.of(url.toURI).toString) ++ urls(Option(u.getParent))
      case Some(other) => urls(Option(other.getParent))
      case None => Nil
    val own = System.getProperty("java.class.path").split(File.pathSeparator).toList
    (own ++ urls(Option(getClass.getClassLoader)))
      .filter(_.nonEmpty)
      .distinct
      .mkString(File.pathSeparator)

  final case class Forked(exit: Int, stdout: String, stderr: String)

  /** Run `mainClass args` in a fresh JVM whose PATH puts `shimDir` first. */
  private def fork(shimDir: Path, mainClass: String, args: String*): Forked =
    val java = Path.of(System.getProperty("java.home"), "bin", "java").toString
    val out = Files.createTempFile("xl-fork-", ".out")
    val err = Files.createTempFile("xl-fork-", ".err")
    try
      val builder = new ProcessBuilder(
        (List(java, "-Djava.awt.headless=true", "-cp", classpath, mainClass) ++ args).asJava
      )
      builder.environment().put("PATH", s"$shimDir${File.pathSeparator}/usr/bin:/bin")
      builder.redirectOutput(out.toFile).redirectError(err.toFile)
      val process = builder.start()
      assert(process.waitFor(180, TimeUnit.SECONDS), "the forked JVM did not finish in 180 s")
      // the JVM's own startup notices (sun.misc.Unsafe deprecation) are not the program's output
      val stderr = Files
        .readString(err, StandardCharsets.UTF_8)
        .linesWithSeparators
        .filterNot(_.startsWith("WARNING: "))
        .mkString
      Forked(process.exitValue(), Files.readString(out, StandardCharsets.UTF_8), stderr)
    finally
      Files.deleteIfExists(out)
      Files.deleteIfExists(err)

  /** `xl -f simple.xlsx view <a large range> --format <fmt> --rasterizer <backend>` in a fork. */
  private def forcedView(shimDir: Path, backend: String, format: String): Forked =
    val fixtures = TestFixtures.materialize.unsafeRunSync()
    try
      fork(
        shimDir,
        "com.tjclp.xl.cli.Main",
        "-f",
        fixtures.resolve("simple.xlsx").toString,
        "-s",
        "Data",
        "view",
        "A1:Z1000",
        "--format",
        format,
        "--raster-output",
        fixtures.resolve(s"out.$format").toString,
        "--rasterizer",
        backend
      )
    finally TestFixtures.delete(fixtures).unsafeRunSync()

  shims.test("forced rsvg-convert that exits early: its stderr and exit, not INTERNAL") { dir =>
    val run = forcedView(dir, "rsvg-convert", "png")
    assertEquals(run.exit, 3, run.stderr)
    assert(
      run.stderr.startsWith(s"Error: rsvg-convert conversion failed (exit 1): $shimMessage"),
      run.stderr
    )
    assert(run.stderr.contains(s"code: ${ErrorCode.RASTERIZER_UNAVAILABLE}"), run.stderr)
    assert(!run.stderr.contains("Stream closed"), run.stderr)
    assert(!run.stderr.contains("\tat "), s"no stack trace: ${run.stderr}")
  }

  shims.test("forced cairosvg that exits early: its stderr and exit, not INTERNAL") { dir =>
    val run = forcedView(dir, "cairosvg", "png")
    assertEquals(run.exit, 3, run.stderr)
    assert(
      run.stderr.startsWith(s"Error: cairosvg conversion failed (exit 1): $shimMessage"),
      run.stderr
    )
    assert(run.stderr.contains(s"code: ${ErrorCode.RASTERIZER_UNAVAILABLE}"), run.stderr)
    assert(!run.stderr.contains("\tat "), s"no stack trace: ${run.stderr}")
  }

  shims.test("forced rsvg-convert asked for jpeg: a typed code, not INTERNAL") { dir =>
    val run = forcedView(dir, "rsvg-convert", "jpeg")
    assertEquals(run.exit, 3, run.stderr)
    assert(run.stderr.startsWith("Error: rsvg-convert does not support jpeg format"), run.stderr)
    assert(run.stderr.contains(s"code: ${ErrorCode.RASTERIZER_UNAVAILABLE}"), run.stderr)
    assert(run.stderr.contains("--rasterizer"), run.stderr)
  }

  shims.test("the default chain lists each backend's failure with its exit, not Stream closed") {
    dir =>
      val out = Files.createTempFile("xl-probe-", ".png")
      try
        val run = fork(dir, "com.tjclp.xl.cli.raster.RasterShimProbe", out.toString)
        assertEquals(run.exit, 0, run.stderr)
        val lines = run.stdout.linesIterator.toVector
        assertEquals(lines.headOption, Some(ErrorCode.RASTERIZER_UNAVAILABLE), run.stdout)
        val message = lines.drop(1).mkString("\n")
        assert(message.contains("cairosvg (cairosvg conversion failed (exit 1)"), message)
        assert(message.contains("rsvg-convert (rsvg-convert conversion failed (exit 1)"), message)
        assert(!message.contains("Stream closed"), message)
        assert(!message.contains("Broken pipe"), message)
      finally Files.deleteIfExists(out)
  }

  test("FormatNotSupported classifies as RASTERIZER_UNAVAILABLE with the lever in the hint") {
    val err =
      CliError.fromThrowable(RasterError.FormatNotSupported("rsvg-convert", RasterFormat.WebP))
    assertEquals(err.code, ErrorCode.RASTERIZER_UNAVAILABLE)
    assertEquals(err.message, "rsvg-convert does not support webp format")
    assert(err.hint.exists(_.contains("--rasterizer")), err.hint)
  }

  test("ConversionFailed classifies as RASTERIZER_UNAVAILABLE, carrying the backend's stderr") {
    val err = CliError.fromThrowable(RasterError.ConversionFailed("resvg", "bad svg", 2))
    assertEquals(err.code, ErrorCode.RASTERIZER_UNAVAILABLE)
    assertEquals(err.message, "resvg conversion failed (exit 2): bad svg")
  }
