package com.tjclp.xl.cli.raster

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import munit.FunSuite

import com.tjclp.xl.cli.MemoryGuard
import com.tjclp.xl.cli.contract.{CliError, ErrorCode}

/**
 * resvg takes file paths only, so its SVG goes through a scratch file: the CLI's one scratch file
 * outside the streaming writer's spill. It lands where `XL_SPILL_DIR` points when set (the same
 * lever, GH-517), else in `java.io.tmpdir`; a directory that cannot take it fails as `IO_WRITE`
 * naming the directory and the lever — it used to surface as `INTERNAL` with a bare path.
 */
class ResvgScratchSpec extends FunSuite:

  private val svg = """<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>"""

  private def failure(spill: Option[Path]): Throwable =
    Resvg.scratchSvg(svg, spill).use(_ => IO.unit).attempt.unsafeRunSync() match
      case Left(t) => t
      case Right(()) => fail("an unusable scratch directory must fail the render")

  test("the scratch SVG lands in XL_SPILL_DIR when set, holds the SVG, and is deleted after use") {
    val spill = Files.createTempDirectory("xl-resvg-spill-")
    try
      val (path, content) = Resvg
        .scratchSvg(svg, Some(spill))
        .use(p => IO((p, Files.readString(p, StandardCharsets.UTF_8))))
        .unsafeRunSync()
      assertEquals(path.getParent, spill)
      assert(path.getFileName.toString.startsWith("xl-resvg-"), path.toString)
      assertEquals(content, svg)
      assert(!Files.exists(path), "the scratch file is deleted when the render ends")
    finally Files.deleteIfExists(spill)
  }

  test("without XL_SPILL_DIR the scratch SVG lands in java.io.tmpdir") {
    val tmpdir = Path.of(System.getProperty("java.io.tmpdir")).toRealPath()
    val parent = Resvg
      .scratchSvg(svg, None)
      .use(p => IO(p.toRealPath().getParent))
      .unsafeRunSync()
    assertEquals(parent, tmpdir)
  }

  test(
    "a spill directory that does not exist fails as IO_WRITE naming XL_SPILL_DIR, not INTERNAL"
  ) {
    val missing = Files.createTempDirectory("xl-resvg-").resolve("missing").resolve("deeper")
    val t = failure(Some(missing))
    t match
      case RasterError.ScratchFileFailed("resvg", Some(dir), _) => assertEquals(dir, missing)
      case other => fail(s"expected ScratchFileFailed, got $other")
    val err = CliError.fromThrowable(t)
    assertEquals(err.code, ErrorCode.IO_WRITE)
    assert(err.message.startsWith("cannot write the resvg scratch SVG: "), err.message)
    assertEquals(
      err.hint,
      Some(s"check that ${MemoryGuard.SpillDirVar} ($missing) exists, is writable and has room")
    )
  }

  test("an unusable default tmpdir classifies with the lever in the hint") {
    val cause = new java.nio.file.NoSuchFileException("/nonexistent/xl-resvg-1.svg")
    val err = CliError.fromThrowable(RasterError.ScratchFileFailed("resvg", None, cause))
    assertEquals(err.code, ErrorCode.IO_WRITE)
    assertEquals(err.message, "cannot write the resvg scratch SVG: /nonexistent/xl-resvg-1.svg")
    assertEquals(
      err.hint,
      Some(
        s"check that the default temp directory java.io.tmpdir (${MemoryGuard.SpillDirVar}=<dir> redirects the scratch file) exists, is writable and has room"
      )
    )
  }
