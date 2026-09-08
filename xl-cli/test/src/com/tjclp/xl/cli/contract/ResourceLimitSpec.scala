package com.tjclp.xl.cli.contract

import cats.effect.IO
import munit.CatsEffectSuite

import com.tjclp.xl.cli.{Cli, CliIO, MemoryGuard}

/**
 * GH-636: memory exhaustion is a typed failure — exit 3, `RESOURCE_LIMIT`, the heap in the message,
 * `--stream` and `-Xmx` in the hint, the envelope under `--json` — never a raw
 * `java.lang.OutOfMemoryError` on stderr with exit 1 and an empty stdout.
 *
 * cats-effect treats an `OutOfMemoryError` as fatal (it halts the runtime), so it cannot be
 * injected into the program as an `IO` error; the seam is [[MemoryGuard.blocking]], the same thunk
 * guard the in-memory load runs under, here as the run's `stdin` so the real program — parser,
 * handler, classification, rendering — carries the failure to the streams.
 */
class ResourceLimitSpec extends CatsEffectSuite:

  private def run(args: List[String], stdin: IO[String]): IO[(Int, String, String)] =
    CliIO.capturing("").flatMap { (base, collect) =>
      Cli.run(args, base.copy(stdin = stdin)).flatMap { code =>
        collect.map((out, err) => (code.code, out, err))
      }
    }

  private val exhausted: IO[String] =
    MemoryGuard.blocking[String](throw new OutOfMemoryError("Garbage-collected heap size exceeded"))

  test("text: exit 3, stdout empty, Error + code RESOURCE_LIMIT + the --stream/-Xmx hint") {
    run(List("batch", "--dry-run", "-"), exhausted).map { (exit, out, err) =>
      assertEquals(exit, 3, err)
      assertEquals(out, "", "stdout must be empty on failure")
      val lines = err.split("\n", -1).toVector
      assertEquals(lines.headOption, Some(s"Error: ${MemoryGuard.exhausted.message}"), err)
      assert(lines.contains(s"  code: ${ErrorCode.RESOURCE_LIMIT}"), err)
      assert(lines.contains(s"  hint: ${MemoryGuard.hint}"), err)
      assert(err.contains("--stream") && err.contains("-Xmx"), err)
      assert(!err.contains("java.lang.OutOfMemoryError"), s"no raw error may leak:\n$err")
    }
  }

  test("--json: one schema-valid envelope, ok:false, exitCode 3, error.code RESOURCE_LIMIT") {
    run(List("--json", "batch", "--dry-run", "-"), exhausted).map { (exit, out, err) =>
      assertEquals(exit, 3, out + err)
      val envelope = ujson.read(out)
      EnvelopeSchema.assertValid(envelope)
      assertEquals(envelope("ok"), ujson.False)
      assertEquals(envelope("exitCode").num.toInt, 3)
      assertEquals(envelope("verb").str, "batch")
      assertEquals(envelope("data"), ujson.Null)
      val error = envelope("error")
      assertEquals(error("code").str, ErrorCode.RESOURCE_LIMIT)
      assertEquals(error("message").str, MemoryGuard.exhausted.message)
      assertEquals(error("hint").str, MemoryGuard.hint)
      assert(
        error("message").str.contains(MemoryGuard.human(MemoryGuard.maxHeapBytes)),
        s"the message names the heap: ${error("message").str}"
      )
      assertEquals(err, s"Error: ${MemoryGuard.exhausted.message}\n")
    }
  }

  test("a different Error subtype is still INTERNAL, exit 3 — only memory exhaustion is typed") {
    run(List("--json", "batch", "--dry-run", "-"), IO.raiseError(new AssertionError("boom"))).map {
      (exit, out, _) =>
        assertEquals(exit, 3, out)
        val envelope = ujson.read(out)
        EnvelopeSchema.assertValid(envelope)
        assertEquals(envelope("error")("code").str, ErrorCode.INTERNAL)
        assertEquals(envelope("error")("message").str, "boom")
    }
  }

  test("RESOURCE_LIMIT is published: in ErrorCode.cli, exit 3, in schema --json's errorCodes") {
    assert(ErrorCode.cli.contains(ErrorCode.RESOURCE_LIMIT))
    assertEquals(ExitCodes.forCode(ErrorCode.RESOURCE_LIMIT), ExitCodes.failed)
    val codes = Schema
      .json("test")("errorCodes")
      .arr
      .map(e => (e("code").str, e("exit").num.toInt))
    assert(codes.contains((ErrorCode.RESOURCE_LIMIT, 3)), codes.toString)
  }
