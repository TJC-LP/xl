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
 * shim directory. Each case runs with an SVG larger than any pipe buffer, so the write cannot
 * complete into the buffer before the shim exits, and with one that fits the buffer, so the shim
 * exits before stdin is closed (closing it then flushes into a stream the JDK has already shut).
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

  /**
   * ImageMagick's availability probe runs `-version`/`--version` and `-list delegate`, then
   * converts a 1x1 SVG on stdin; this shim passes the first two and exits on the conversion without
   * reading.
   */
  private val magickShimScript =
    s"""#!/bin/sh
       |case "$$1" in
       |  -version|--version|-list) echo "shim 1.0"; exit 0 ;;
       |esac
       |echo "$shimMessage" >&2
       |exit 1
       |""".stripMargin

  /**
   * A directory holding a failing shim for each stdin-fed backend (and resvg, which is not). Both
   * ImageMagick commands are shimmed so a system `convert` on the fork's PATH cannot answer the
   * probe instead.
   */
  private val shims = shimsWith(shimScript, magickShimScript)

  /** A shim directory: `script` for rsvg-convert, cairosvg and resvg, `magick` for ImageMagick. */
  private def shimsWith(script: String, magick: String): FunFixture[Path] = FunFixture[Path](
    setup = _ =>
      val dir = Files.createTempDirectory("xl-raster-shims-")
      def install(name: String, body: String): Unit =
        val shim = dir.resolve(name)
        Files.writeString(shim, body, StandardCharsets.UTF_8)
        Files.setPosixFilePermissions(shim, PosixFilePermissions.fromString("rwxr-xr-x"))
      Vector("rsvg-convert", "cairosvg", "resvg").foreach(install(_, script))
      Vector("magick", "convert").foreach(install(_, magick))
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
    forkWith(shimDir, Nil, mainClass, args*)

  /** [[fork]], with extra flags for the forked JVM. */
  private def forkWith(
    shimDir: Path,
    jvmFlags: List[String],
    mainClass: String,
    args: String*
  ): Forked =
    val java = Path.of(System.getProperty("java.home"), "bin", "java").toString
    val out = Files.createTempFile("xl-fork-", ".out")
    val err = Files.createTempFile("xl-fork-", ".err")
    try
      val builder = new ProcessBuilder(
        (List(java, "-Djava.awt.headless=true") ++ jvmFlags ++
          List("-cp", classpath, mainClass) ++ args).asJava
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

  /** `xl -f simple.xlsx view <range> --format <fmt> --rasterizer <backend>` in a fork. */
  private def forcedView(
    shimDir: Path,
    backend: String,
    format: String,
    range: String,
    jvmFlags: List[String] = Nil,
    output: Option[Path] = None
  ): Forked =
    val fixtures = TestFixtures.materialize.unsafeRunSync()
    try
      forkWith(
        shimDir,
        jvmFlags,
        "com.tjclp.xl.cli.Main",
        "-f",
        fixtures.resolve("simple.xlsx").toString,
        "-s",
        "Data",
        "view",
        range,
        "--format",
        format,
        "--raster-output",
        output.getOrElse(fixtures.resolve(s"out.$format")).toString,
        "--rasterizer",
        backend
      )
    finally TestFixtures.delete(fixtures).unsafeRunSync()

  /** The sizes each forced case runs at: past the pipe buffer, and inside it. */
  private val ranges = Vector("A1:Z1000" -> "large SVG", "A1:B2" -> "small SVG")

  private def assertExitedEarly(run: Forked, backend: String): Unit =
    assertEquals(run.exit, 3, run.stderr)
    assert(
      run.stderr.startsWith(s"Error: $backend conversion failed (exit 1): $shimMessage"),
      run.stderr
    )
    assert(run.stderr.contains(s"code: ${ErrorCode.RASTERIZER_UNAVAILABLE}"), run.stderr)
    assert(!run.stderr.contains("Stream closed"), run.stderr)
    assert(!run.stderr.contains("\tat "), s"no stack trace: ${run.stderr}")

  for (range, size) <- ranges do
    shims.test(s"forced rsvg-convert that exits early ($size): its stderr and exit, not INTERNAL") {
      dir => assertExitedEarly(forcedView(dir, "rsvg-convert", "png", range), "rsvg-convert")
    }

    shims.test(s"forced cairosvg that exits early ($size): its stderr and exit, not INTERNAL") {
      dir => assertExitedEarly(forcedView(dir, "cairosvg", "png", range), "cairosvg")
    }

  shims.test(
    "forced imagemagick whose probe conversion exits early: RASTERIZER_UNAVAILABLE, no trace"
  ) { dir =>
    // The probe's 1x1 SVG fits the stdin buffer, so the old fs2 write failed only when the shim
    // exited before fs2 closed stdin — a race the JIT usually won. An interpreted JVM (-Xint) loses
    // it every time (measured 4/4 against the fs2 probe, 1/16 with the JIT under load); the repeats
    // keep the case honest if a faster machine narrows the window again.
    (1 to 3).foreach { attempt =>
      val run = forcedView(dir, "imagemagick", "png", "A1:B2", List("-Xint"))
      val clue = s"attempt $attempt: ${run.stderr}"
      assertEquals(run.exit, 3, clue)
      assert(run.stderr.contains(s"code: ${ErrorCode.RASTERIZER_UNAVAILABLE}"), clue)
      assert(!run.stderr.contains("Stream closed"), clue)
      assert(!run.stderr.contains("\tat "), s"no stack trace, $clue")
    }
  }

  shims.test("forced rsvg-convert asked for jpeg: a typed code, not INTERNAL") { dir =>
    val run = forcedView(dir, "rsvg-convert", "jpeg", "A1:Z1000")
    assertEquals(run.exit, 3, run.stderr)
    assert(run.stderr.startsWith("Error: rsvg-convert does not support jpeg format"), run.stderr)
    assert(run.stderr.contains(s"code: ${ErrorCode.RASTERIZER_UNAVAILABLE}"), run.stderr)
    assert(run.stderr.contains("--rasterizer"), run.stderr)
  }

  for size <- Vector("large", "small") do
    shims.test(
      s"the default chain ($size SVG) lists each backend's failure with its exit, not Stream closed"
    ) { dir =>
      val out = Files.createTempFile("xl-probe-", ".png")
      try
        val run = fork(dir, "com.tjclp.xl.cli.raster.RasterShimProbe", out.toString, size)
        assertEquals(run.exit, 0, run.stderr)
        assert(!run.stderr.contains("\tat "), s"no stack trace: ${run.stderr}")
        assert(!run.stderr.contains("Stream closed"), run.stderr)
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

  // ===== GH-690: what a backend's exit cannot prove, the output file must =====

  /** Answers every probe; a conversion exits 0 at once, reading nothing and writing nothing. */
  private val silentZeroScript =
    """#!/bin/sh
      |case "$1" in
      |  --version|--help) echo "shim 1.0"; exit 0 ;;
      |esac
      |exit 0
      |""".stripMargin

  /**
   * ImageMagick's probe converts to stdout (`png:-`) and must see bytes; the real run gets none.
   */
  private val magickSilentZeroScript =
    """#!/bin/sh
      |case "$1" in
      |  -version|--version|-list) echo "shim 1.0"; exit 0 ;;
      |esac
      |for a in "$@"; do [ "$a" = "png:-" ] && { cat >/dev/null; echo PNG; exit 0; }; done
      |exit 0
      |""".stripMargin

  /** Writes stdin to the `-o` argument: a backend that works. */
  private val writingScript =
    """#!/bin/sh
      |case "$1" in
      |  --version|--help) echo "shim 1.0"; exit 0 ;;
      |esac
      |out=""; prev=""
      |for a in "$@"; do [ "$prev" = "-o" ] && out="$a"; prev="$a"; done
      |cat > "$out"
      |exit 0
      |""".stripMargin

  /** Exits 1 with nothing on stderr. */
  private val muteFailureScript =
    """#!/bin/sh
      |case "$1" in
      |  --version|--help|-version|-list) echo "shim 1.0"; exit 0 ;;
      |esac
      |exit 1
      |""".stripMargin

  /** Exits 1 with a 200-character stderr message, past the chain's old 50-character cut. */
  private val longMessage = "shim: " + ("detail-" * 28) + "END"
  private val longFailureScript =
    s"""#!/bin/sh
       |case "$$1" in
       |  --version|--help) echo "shim 1.0"; exit 0 ;;
       |esac
       |echo "$longMessage" >&2
       |exit 1
       |""".stripMargin

  private val silentZero = shimsWith(silentZeroScript, magickSilentZeroScript)
  private val writing = shimsWith(writingScript, magickSilentZeroScript)
  private val mute = shimsWith(muteFailureScript, muteFailureScript)
  private val chatty = shimsWith(longFailureScript, magickShimScript)

  /** A scratch directory for an output path, emptied and removed afterwards. */
  private def withOutputDir[A](body: Path => A): A =
    val dir = Files.createTempDirectory("xl-raster-out-")
    try body(dir)
    finally
      val entries = Files.list(dir)
      try entries.forEach(p => Files.deleteIfExists(p))
      finally entries.close()
      Files.deleteIfExists(dir)

  private def listing(dir: Path): List[String] =
    val entries = Files.list(dir)
    try entries.iterator().asScala.map(_.getFileName.toString).toList.sorted
    finally entries.close()

  for (backend, format) <- Vector(
      "rsvg-convert" -> "png",
      "cairosvg" -> "png",
      "imagemagick" -> "png"
    )
  do
    silentZero.test(s"GH-690: forced $backend that exits 0 without output fails, no file") { dir =>
      withOutputDir { outDir =>
        val out = outDir.resolve(s"out.$format")
        val run = forcedView(dir, backend, format, "A1:B2", output = Some(out))
        assertEquals(run.exit, 3, run.stderr)
        val name = if backend == "imagemagick" then "ImageMagick" else backend
        assert(
          run.stderr.startsWith(
            s"Error: $name conversion failed (exit 0): exited 0 without writing any output"
          ),
          run.stderr
        )
        assert(!run.stdout.contains("Exported"), run.stdout)
        assertEquals(listing(outDir), Nil, "no output and no staging file left behind")
      }
    }

  silentZero.test("GH-690: an earlier run's output cannot pass for this run's, and survives") {
    dir =>
      withOutputDir { outDir =>
        val out = outDir.resolve("out.png")
        Files.writeString(out, "OLD", StandardCharsets.UTF_8)
        val run = forcedView(dir, "rsvg-convert", "png", "A1:B2", output = Some(out))
        assertEquals(run.exit, 3, run.stderr)
        assertEquals(Files.readString(out, StandardCharsets.UTF_8), "OLD")
        assertEquals(listing(outDir), List("out.png"))
      }
  }

  silentZero.test("GH-690: exit 0 before reading a large SVG: no JDK wording in the message") {
    dir =>
      val run = forcedView(dir, "rsvg-convert", "png", "A1:Z1000")
      assertEquals(run.exit, 3, run.stderr)
      assert(
        run.stderr.startsWith(
          "Error: rsvg-convert conversion failed (exit 0): exited before reading the whole SVG"
        ),
        run.stderr
      )
      assert(!run.stderr.contains("Stream closed"), run.stderr)
      assert(!run.stderr.contains("Broken pipe"), run.stderr)
  }

  writing.test("GH-690: a backend that writes replaces the old output and leaves no staging file") {
    dir =>
      withOutputDir { outDir =>
        val out = outDir.resolve("out.png")
        Files.writeString(out, "OLD", StandardCharsets.UTF_8)
        val run = forcedView(dir, "rsvg-convert", "png", "A1:B2", output = Some(out))
        assertEquals(run.exit, 0, run.stderr)
        assert(run.stdout.contains(s"Exported: $out"), run.stdout)
        assert(Files.readString(out, StandardCharsets.UTF_8).startsWith("<svg"))
        assertEquals(listing(outDir), List("out.png"))
      }
  }

  mute.test("GH-690: an empty stderr ends the message at the exit code, no dangling colon") { dir =>
    val run = forcedView(dir, "rsvg-convert", "png", "A1:B2")
    assertEquals(run.exit, 3, run.stderr)
    assertEquals(
      run.stderr.linesIterator.nextOption(),
      Some("Error: rsvg-convert conversion failed (exit 1)")
    )
  }

  chatty.test("GH-690: the default chain keeps each backend's whole message") { dir =>
    val out = Files.createTempFile("xl-probe-", ".png")
    try
      val run = fork(dir, "com.tjclp.xl.cli.raster.RasterShimProbe", out.toString, "small")
      assertEquals(run.exit, 0, run.stderr)
      val message = run.stdout.linesIterator.drop(1).mkString("\n")
      assert(
        message.contains(s"cairosvg (cairosvg conversion failed (exit 1): $longMessage)"),
        message
      )
      assert(
        message.contains(s"rsvg-convert (rsvg-convert conversion failed (exit 1): $longMessage)"),
        message
      )
    finally Files.deleteIfExists(out)
  }

  silentZero.test("GH-690: the default chain does not report Exported when no backend wrote") {
    dir =>
      val out = Files.createTempFile("xl-probe-", ".png")
      Files.delete(out)
      try
        val run = fork(dir, "com.tjclp.xl.cli.raster.RasterShimProbe", out.toString, "small")
        assertEquals(run.exit, 0, run.stderr)
        val lines = run.stdout.linesIterator.toVector
        assertEquals(lines.headOption, Some(ErrorCode.RASTERIZER_UNAVAILABLE), run.stdout)
        val message = lines.drop(1).mkString("\n")
        Vector("cairosvg", "rsvg-convert", "resvg").foreach { backend =>
          assert(
            message.contains(
              s"$backend ($backend conversion failed (exit 0): exited 0 without writing any output)"
            ),
            message
          )
        }
        assert(!Files.exists(out))
      finally Files.deleteIfExists(out)
  }
