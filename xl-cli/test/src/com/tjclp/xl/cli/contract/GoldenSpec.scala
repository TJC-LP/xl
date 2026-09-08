package com.tjclp.xl.cli.contract

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

/**
 * The CLI contract corpus: every `<name>.golden` under `xl-cli/test/resources/golden/` runs through
 * the in-process harness and must reproduce its recorded exit code, stdout and stderr byte for byte
 * (after [[Golden.normalize]]).
 *
 * These goldens ARE the agent-visible contract of `xl`: a diff here is a contract change and needs
 * a CHANGELOG line. Re-record deliberately with `XL_UPDATE_GOLDEN=1` (e.g.
 * `XL_UPDATE_GOLDEN=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.GoldenSpec`) and review
 * the resulting diff like code. `XL_GOLDEN_KEEP=1` leaves the fixture directory behind and prints
 * its path, for reproducing a case by hand against the same files.
 *
 * Fixtures are written fresh per run ([[TestFixtures]]); cases that write do so under the same
 * directory, so `<DIR>` in a golden's args stands for both inputs and outputs.
 */
class GoldenSpec extends CatsEffectSuite:

  private val update = sys.env.get("XL_UPDATE_GOLDEN").contains("1")
  private val keep = sys.env.get("XL_GOLDEN_KEEP").contains("1")

  private val fixtures = ResourceSuiteLocalFixture(
    "golden-fixtures",
    Resource.make(TestFixtures.materialize) { dir =>
      if keep then IO.println(s"XL_GOLDEN_KEEP: fixtures left at $dir")
      else TestFixtures.delete(dir)
    }
  )

  override def munitFixtures = List(fixtures)

  private val cases: Vector[Path] = Golden.caseFiles

  test("the corpus is present") {
    assert(cases.nonEmpty, s"no *.golden files under ${Golden.corpusDir}")
  }

  cases.foreach { file =>
    val name = file.getFileName.toString.stripSuffix(".golden")
    test(s"golden: $name") {
      val dir = fixtures()
      val recorded = Golden.readCase(file).fold(msg => fail(msg), identity)
      val args = recorded.args.map(_.replace("<DIR>", dir.toString))
      CliHarness.run(args, recorded.stdin.getOrElse("")).flatMap { run =>
        val observed = recorded.observed(
          run.exit,
          Golden.normalize(run.stdout, dir),
          Golden.normalize(run.stderr, dir)
        )
        if update then
          IO.blocking(Files.writeString(file, GoldenCase.render(observed), StandardCharsets.UTF_8))
            .void
        else IO(check(recorded, observed))
      }
    }
  }

  private def check(recorded: GoldenCase, observed: GoldenCase): Unit =
    val problems = Vector(
      recorded.exit match
        case None => Some("exit: not recorded")
        case Some(expected) if observed.exit.contains(expected) => None
        case Some(expected) => Some(s"exit: expected $expected, got ${observed.exit.getOrElse(-1)}")
      ,
      channel("stdout", recorded.stdout, observed.stdout.getOrElse("")),
      channel("stderr", recorded.stderr, observed.stderr.getOrElse(""))
    ).flatten
    if problems.nonEmpty then
      fail(
        s"golden '${recorded.name}' differs from the recorded contract " +
          s"(re-record deliberately with XL_UPDATE_GOLDEN=1 and add a CHANGELOG line):\n" +
          problems.mkString("\n")
      )

  private def channel(label: String, recorded: Option[String], actual: String): Option[String] =
    recorded match
      case None => Some(s"$label: not recorded")
      case Some(expected) if expected == actual => None
      case Some(expected) => Some(s"$label:\n${UnifiedDiff.render(expected, actual)}")
