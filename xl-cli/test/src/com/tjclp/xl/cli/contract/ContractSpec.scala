package com.tjclp.xl.cli.contract

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import scala.jdk.CollectionConverters.*
import scala.util.matching.Regex

import munit.FunSuite

import com.tjclp.xl.cli.BuildInfo
import com.tjclp.xl.error.XLError

/**
 * The code vocabulary enumerated end to end (ADR-017 §2.3, GH-592). [[DocsGenSpec]] proves the
 * generated pages equal what `Schema.json` renders; this suite reads
 * `docs/reference/generated/error-codes.md` back the way an agent does — as table rows — and holds
 * it against the source of truth ([[ErrorCode.all]], [[WarningCode.all]], [[ExitCodes.forCode]])
 * and against every code the golden corpus records. A code the CLI emitted that the table does not
 * publish, or an exit the table does not predict from the code alone, is a contract breach.
 */
class ContractSpec extends FunSuite:

  private val page: Path = Golden.repoRoot.resolve("docs/reference/generated/error-codes.md")

  /** `| \`CODE\` | 3 |` and `| \`CODE\` |` rows; the header and rule rows carry no backticks. */
  private val ErrorRow: Regex = """^\|\s*`([A-Z_]+)`\s*\|\s*(\d)\s*\|\s*$""".r
  private val WarningRow: Regex = """^\|\s*`([A-Z_]+)`\s*\|\s*$""".r

  /** `  code: <CODE>` on stderr (text mode). */
  private val CodeLine: Regex = """^\s*code: ([A-Z_]+)\s*$""".r

  /** `Warning[<CODE>]: <message>` on stderr (text mode). */
  private val WarningLine: Regex = """^Warning\[([A-Z_]+)\]: .*$""".r

  private lazy val text: String = Files.readString(page, StandardCharsets.UTF_8)

  private def section(heading: String): Vector[String] =
    val lines = text.linesIterator.toVector
    val start = lines.indexWhere(_.trim == s"## $heading")
    assert(start >= 0, s"no '## $heading' section in $page")
    lines.drop(start + 1).takeWhile(l => !l.startsWith("## "))

  private lazy val errorRows: Vector[(String, Int)] =
    section("Error codes").collect { case ErrorRow(code, exit) => (code, exit.toInt) }

  private lazy val warningRows: Vector[String] =
    section("Warning codes").collect { case WarningRow(code) => code }

  private lazy val published: Set[String] = errorRows.map(_._1).toSet

  private lazy val exitOf: Map[String, Int] = errorRows.toMap

  // --- the page against the vocabulary --------------------------------------------------------

  test("error-codes.md publishes ErrorCode.all: every code once, in order, with ExitCodes' exit") {
    assertEquals(errorRows.map(_._1), ErrorCode.all)
    errorRows.foreach { (code, exit) =>
      assertEquals(exit, ExitCodes.forCode(code).code, s"$code's exit in the table")
    }
  }

  test("error-codes.md publishes WarningCode.all: every code once, in order") {
    assertEquals(warningRows, WarningCode.all)
    assertEquals(warningRows.distinct, warningRows)
  }

  test(
    "the table is exactly the CLI codes plus every XLError code, and no code is also a warning"
  ) {
    assert(ErrorCode.cli.forall(published.contains), ErrorCode.cli.filterNot(published.contains))
    assert(XLError.codes.forall(published.contains), XLError.codes.filterNot(published.contains))
    assertEquals(published, (ErrorCode.cli ++ XLError.codes).toSet)
    assertEquals(published.intersect(warningRows.toSet), Set.empty[String])
  }

  test("every published code exits 1, 2 or 3; exit 1 is exactly the findings and gates") {
    assert(errorRows.forall((_, exit) => Set(1, 2, 3).contains(exit)))
    assertEquals(
      errorRows.collect { case (code, 1) => code }.toSet,
      Set(
        ErrorCode.RECALC_GATE,
        ErrorCode.DIFFERENCES_FOUND,
        ErrorCode.LINT_FINDINGS,
        ErrorCode.AUDIT_FINDINGS
      )
    )
  }

  test("`xl schema --json` carries the same two tables, row for row") {
    val schema = Schema.json(BuildInfo.version)
    val schemaErrors = schema("errorCodes").arr.toVector.map { e =>
      (e("code").str, e("exit").num.toInt)
    }
    val schemaWarnings = schema("warningCodes").arr.toVector.map(_.str)
    assertEquals(schemaErrors, errorRows)
    assertEquals(schemaWarnings, warningRows)
  }

  // --- the golden corpus against the page ------------------------------------------------------

  private lazy val goldens: Vector[GoldenCase] =
    val listing = Files.list(Golden.corpusDir)
    try
      listing.iterator.asScala
        .filter(_.toString.endsWith(".golden"))
        .toVector
        .sortBy(_.toString)
        .map { file =>
          val name = file.getFileName.toString.stripSuffix(".golden")
          GoldenCase
            .parse(name, Files.readString(file, StandardCharsets.UTF_8))
            .fold(msg => fail(msg), identity)
        }
    finally listing.close()

  /** A golden's stdout as the `--json` envelope, when it is one. */
  private def envelopeOf(g: GoldenCase): Option[ujson.Obj] =
    g.stdout.filter(_.startsWith("{")).flatMap { out =>
      scala.util.Try(ujson.read(out)).toOption.collect {
        case obj: ujson.Obj if obj.value.contains("ok") && obj.value.contains("error") => obj
      }
    }

  /** (golden, code, recorded exit) for every error the corpus records, on either channel. */
  private lazy val recordedErrors: Vector[(String, String, Int)] =
    goldens.flatMap { g =>
      val exit = g.exit.getOrElse(fail(s"${g.name}.golden: exit not recorded"))
      val text = g.stderr.toList.flatMap(_.linesIterator).collect { case CodeLine(code) => code }
      val json = envelopeOf(g).toList.flatMap { env =>
        env("error") match
          case err: ujson.Obj =>
            assertEquals(env("exitCode").num.toInt, exit, s"${g.name}: envelope exitCode")
            List(err("code").str)
          case _ => Nil
      }
      (text ++ json).map(code => (g.name, code, exit))
    }

  /** (golden, code) for every warning the corpus records, on either channel. */
  private lazy val recordedWarnings: Vector[(String, String)] =
    goldens.flatMap { g =>
      val text =
        g.stderr.toList.flatMap(_.linesIterator).collect { case WarningLine(code) => code }
      val json = envelopeOf(g).toList.flatMap(_("warnings").arr.map(_("code").str))
      (text ++ json).map(code => (g.name, code))
    }

  test("the corpus is present and records errors and warnings on both channels") {
    assert(goldens.nonEmpty, s"no *.golden files under ${Golden.corpusDir}")
    assert(recordedErrors.nonEmpty, "no error recorded in any golden")
    assert(recordedWarnings.nonEmpty, "no warning recorded in any golden")
    assert(goldens.exists(g => envelopeOf(g).isDefined), "no --json golden")
  }

  test("every error code the corpus records is published, and its exit is the table's") {
    recordedErrors.foreach { (golden, code, exit) =>
      assert(published.contains(code), s"$golden.golden records unpublished code $code")
      assertEquals(exitOf.get(code), Some(exit), s"$golden.golden: exit for $code")
    }
  }

  test("every warning code the corpus records is published") {
    recordedWarnings.foreach { (golden, code) =>
      assert(warningRows.contains(code), s"$golden.golden records unpublished warning $code")
    }
  }

  test(
    "every failure (exit 2 or 3) in the corpus carries a code — an agent never branches on text"
  ) {
    // Exit 1 is "completed with findings or a failed gate": in text mode the findings ARE the
    // report and no `code:` line is printed; the `--json` twin carries `error.code`.
    val withCode = recordedErrors.map(_._1).toSet
    goldens.filter(_.exit.exists(e => e == 2 || e == 3)).foreach { g =>
      assert(withCode.contains(g.name), s"${g.name}.golden exits ${g.exit} without a code")
    }
  }

  test("every --json golden that exits 1 carries error.code (findings are addressable by code)") {
    goldens.filter(_.exit.contains(1)).flatMap(g => envelopeOf(g).map(g.name -> _)).foreach {
      (name, env) =>
        env("error") match
          case err: ujson.Obj =>
            assert(published.contains(err("code").str), s"$name.golden: unpublished code")
          case other => fail(s"$name.golden: exit 1 envelope without error: ${ujson.write(other)}")
    }
  }

  test("the corpus exercises each exit class the table defines (the scans above are not vacuous)") {
    val exits = recordedErrors.map(_._3).toSet
    assertEquals(exits, Set(1, 2, 3))
  }
