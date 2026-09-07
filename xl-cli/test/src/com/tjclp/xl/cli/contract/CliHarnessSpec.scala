package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path}
import java.util.Locale

import cats.effect.IO
import cats.syntax.all.*
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.{BuildInfo, Cli, CliIO}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * The harness itself: every property the golden corpus relies on to be meaningful.
 */
class CliHarnessSpec extends CatsEffectSuite:

  private def withTempExcelFile[A](test: Path => IO[A]): IO[A] =
    IO.blocking {
      val tempFile = Files.createTempFile("xl-harness-", ".xlsx")
      tempFile.toFile.deleteOnExit()
      tempFile
    }.flatMap { tempFile =>
      val sheet = Sheet("Test")
        .put(ref"A1", CellValue.Text("Hello"))
        .put(ref"A2", CellValue.Number(BigDecimal("1.5")))
      ExcelIO.instance[IO].write(Workbook(Vector(sheet)), tempFile) *> test(tempFile)
    }

  private def occurrences(text: String, needle: String): Int =
    text.sliding(needle.length).count(_ == needle)

  test("stdout and stderr are captured separately") {
    for
      version <- CliHarness.run("--version")
      help <- CliHarness.run("--help")
    yield
      assertEquals(version.exit, 0)
      assertEquals(version.stdout, s"${BuildInfo.version}\n")
      assertEquals(version.stderr, "")
      // decline-effect's contract, kept as-is in this cluster: help renders on stderr, exit 0
      assertEquals(help.exit, 0)
      assertEquals(help.stdout, "")
      assert(help.stderr.startsWith("Usage:"), s"help must render on stderr, got:\n${help.stderr}")
  }

  test("stdin is delivered to `batch -`") {
    CliHarness
      .run(List("batch", "--dry-run", "-"), """[{"op":"put","ref":"A1","value":"x"}]""")
      .map { run =>
        assertEquals(run.exit, 0)
        assert(
          run.stdout.contains("Dry run - 1 operations parsed"),
          s"stdin was not consumed by batch -:\n${run.stdout}\n${run.stderr}"
        )
        assert(run.stdout.contains("PUT A1 = Text(x)"), run.stdout)
      }
  }

  test("exit code propagates") {
    for
      unknown <- CliHarness.run("frob")
      lintMissing <- CliHarness.run("lint", "/nonexistent/no-such-file.xlsx")
      eval <- CliHarness.run("eval", "=1+1")
    yield
      // ADR-017 §2.3: an unknown verb is usage (2); an unreadable file is a failure (3)
      assertEquals(unknown.exit, 2)
      assert(unknown.stderr.contains("unknown verb 'frob'"), unknown.stderr)
      assertEquals(lintMissing.exit, 3)
      assertEquals(eval.exit, 0)
      assertEquals(eval.stdout, "Formula: =1+1\nResult: 2 (number)\n")
  }

  test("batch parse warnings go through the run's warning sink, in both modes") {
    withTempExcelFile { file =>
      val textOut = file.resolveSibling(file.getFileName.toString + ".text.xlsx")
      val jsonOut = file.resolveSibling(file.getFileName.toString + ".json.xlsx")
      val ops = """[{"op":"put","ref":"A1","value":"x","bogus":true}]"""
      def args(out: Path): List[String] =
        List("-f", file.toString, "-s", "Test", "-o", out.toString, "batch", "-")
      for
        text <- CliHarness.run(args(textOut), ops)
        json <- CliHarness.run("--json" :: args(jsonOut), ops)
      yield
        // Text mode: the warning is a `Warning[CODE]:` line on stderr, never a bare handler print
        assertEquals(text.exit, 0, text.stdout + text.stderr)
        assert(
          text.stderr.contains(
            "Warning[UNKNOWN_PROPERTY]: Object 1 (put): unknown properties ignored: bogus"
          ),
          s"warning did not reach stderr through the sink:\n${text.stderr}"
        )
        assert(!text.stderr.contains("Warning: Object"), s"bypassed the sink:\n${text.stderr}")
        assert(!text.stdout.contains("bogus"), s"warning leaked to stdout:\n${text.stdout}")
        assert(text.stdout.contains("Applied 1 operations"), text.stdout)
        // --json: the warning rides in the envelope's warnings[] with the op's index; stderr is empty
        assertEquals(json.exit, 0, json.stdout + json.stderr)
        assertEquals(json.stderr, "", "under --json a warning belongs to the envelope")
        val warnings = ujson.read(json.stdout)("warnings").arr
        assertEquals(warnings.size, 1, json.stdout)
        assertEquals(warnings(0)("code").str, "UNKNOWN_PROPERTY")
        assertEquals(warnings(0)("location")("opIndex").num.toInt, 1)
        assert(
          warnings(0)("message").str
            .startsWith("Object 1 (put): unknown properties ignored: bogus"),
          warnings(0)("message").str
        )
        assert(!json.stdout.contains("\"warnings\": [],"), "the envelope must carry the warning")
    }
  }

  test("the harness restores the JVM streams") {
    val before = (System.out, System.err)
    CliHarness.run("--version").map { _ =>
      assert(System.out eq before._1, "System.out was not restored")
      assert(System.err eq before._2, "System.err was not restored")
    }
  }

  test("two concurrent harness runs do not interleave stderr") {
    val helpArgs = List("--help")
    val frobArgs = List("frob")
    for
      help <- CliHarness.run(helpArgs, "")
      frob <- CliHarness.run(frobArgs, "")
      concurrent <- List(helpArgs, frobArgs, helpArgs, frobArgs, helpArgs, frobArgs)
        .parTraverse(args => CliHarness.run(args, "").map(args -> _))
    yield
      concurrent.foreach { (args, run) =>
        val sequential = if args == helpArgs then help else frob
        assertEquals(run, sequential, s"concurrent run of $args differs from its sequential result")
      }
      // Nothing crosses channels or runs: each stderr carries exactly its own block
      assertEquals(help.stdout, "")
      assertEquals(frob.stdout, "")
      assertEquals(occurrences(help.stderr, "Usage:"), 1)
      assertEquals(occurrences(frob.stderr, "usage: xl"), 1)
      assert(!help.stderr.contains("unknown verb"), help.stderr)
      assertEquals(occurrences(frob.stderr, "unknown verb 'frob'"), 1)
  }

  test("the harness pins the default locale to US for the run and restores the caller's") {
    withTempExcelFile { file =>
      val original = Locale.getDefault
      // Premise: stats renders through f"%.2f", which under Locale.GERMANY prints a decimal comma
      assertEquals(String.format(Locale.GERMANY, "%.2f", Double.box(1.5)), "1,50")
      IO.blocking(Locale.setDefault(Locale.GERMANY))
        .bracket { _ =>
          CliHarness.run("-f", file.toString, "stats", "Test!A1:A2").map { run =>
            assertEquals(run.exit, 0, run.stdout + run.stderr)
            assert(run.stdout.contains("sum: 1.50"), s"locale not pinned to US:\n${run.stdout}")
            assertEquals(Locale.getDefault, Locale.GERMANY, "caller's locale was not restored")
          }
        }(_ => IO.blocking(Locale.setDefault(original)))
    }
  }

  test("CliIO.capturing keeps the channels apart without touching the JVM streams") {
    val before = (System.out, System.err)
    CliIO.capturing("").flatMap { (io, collect) =>
      Cli.run(List("--version"), io).flatMap { code =>
        collect.map { (out, err) =>
          assertEquals(code.code, 0)
          assertEquals(out, s"${BuildInfo.version}\n")
          assertEquals(err, "")
          assert(System.out eq before._1)
          assert(System.err eq before._2)
        }
      }
    }
  }

  test("CliIO.capturing feeds its stdin to `batch -`") {
    CliIO.capturing("""[{"op":"merge","range":"A1:B1"}]""").flatMap { (io, collect) =>
      Cli.run(List("batch", "--dry-run", "-"), io).flatMap { code =>
        collect.map { (out, err) =>
          assertEquals(code.code, 0)
          assert(out.contains("MERGE A1:B1"), out)
          assertEquals(err, "")
        }
      }
    }
  }
