package com.tjclp.xl.cli.contract

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import scala.util.matching.Regex

import munit.FunSuite

import com.tjclp.xl.cli.BuildInfo

/**
 * The code vocabulary held against what agents actually see (ADR-017 §2.3, GH-592).
 *
 * The vocabulary itself is pinned elsewhere: [[SchemaSpec]] proves `schema --json` publishes
 * [[ErrorCode.all]] in order with [[ExitCodes]]' exits and [[WarningCode.all]]; [[DocsGenSpec]]
 * proves `docs/reference/generated/error-codes.md` is that schema rendered; [[ExitCodesSpec]] pins
 * which codes exit 1, 2 and 3; [[CliErrorSpec]] that the vocabulary is the CLI codes plus every
 * `XLError` code. This suite adds the two things none of them look at. The page read back the way
 * an agent reads it — as table rows — must be the schema's rows. And every code the golden corpus
 * records, on either channel, must be published with the exit the table predicts from the code
 * alone: a code the CLI emitted that the table does not publish, or an exit the table does not
 * predict, is a contract breach.
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

  /** The lines under `## <heading>`; empty when the page has no such section. */
  private def section(heading: String): Vector[String] =
    val lines = text.linesIterator.toVector
    val start = lines.indexWhere(_.trim == s"## $heading")
    if start < 0 then Vector.empty else lines.drop(start + 1).takeWhile(l => !l.startsWith("## "))

  private lazy val errorRows: Vector[(String, Int)] =
    section("Error codes").collect { case ErrorRow(code, exit) => (code, exit.toInt) }

  private lazy val warningRows: Vector[String] =
    section("Warning codes").collect { case WarningRow(code) => code }

  private lazy val published: Set[String] = errorRows.map(_._1).toSet

  private lazy val exitOf: Map[String, Int] = errorRows.toMap

  // --- the page, read as rows, against the schema ----------------------------------------------

  test("error-codes.md read as table rows is `xl schema --json`'s errorCodes and warningCodes") {
    val schema = Schema.json(BuildInfo.version)
    val schemaErrors =
      schema("errorCodes").arr.toVector.map(e => (e("code").str, e("exit").num.toInt))
    val schemaWarnings = schema("warningCodes").arr.toVector.map(_.str)
    assert(errorRows.nonEmpty, s"no `| CODE | exit |` rows under '## Error codes' in $page")
    assert(warningRows.nonEmpty, s"no `| CODE |` rows under '## Warning codes' in $page")
    assertEquals(errorRows, schemaErrors)
    assertEquals(warningRows, schemaWarnings)
    assertEquals(published.intersect(warningRows.toSet), Set.empty[String], "a code on both tables")
  }

  // --- the golden corpus against the page ------------------------------------------------------

  private lazy val parsed: Vector[Either[String, GoldenCase]] =
    Golden.caseFiles.map(Golden.readCase)

  private lazy val goldens: Vector[GoldenCase] = parsed.collect { case Right(g) => g }

  private def field(obj: ujson.Obj, key: String): Option[ujson.Value] = obj.value.get(key)

  /**
   * A golden's stdout as the `--json` envelope, when it is one: a JSON object carrying `ok`. A
   * `--format json` payload (`view`, `diff`) is a JSON object too, but has no `ok` — it is a verb's
   * data, not the envelope, and is not held to the envelope's shape.
   */
  private def envelopeOf(g: GoldenCase): Option[ujson.Obj] =
    g.stdout.map(_.trim).filter(_.startsWith("{")).flatMap { out =>
      scala.util.Try(ujson.read(out)).toOption.collect {
        case obj: ujson.Obj if obj.value.contains("ok") => obj
      }
    }

  private lazy val envelopes: Vector[(GoldenCase, ujson.Obj)] =
    goldens.flatMap(g => envelopeOf(g).map(g -> _))

  private def recordedExit(g: GoldenCase): Int = g.exit.getOrElse(-1)

  /** The `code` of an envelope's `error` object, when there is one. */
  private def errorCode(env: ujson.Obj): Option[String] =
    field(env, "error").collect { case err: ujson.Obj => err }.flatMap { err =>
      field(err, "code").collect { case ujson.Str(code) => code }
    }

  /** (golden, code, recorded exit) for every error the corpus records, on either channel. */
  private lazy val recordedErrors: Vector[(String, String, Int)] =
    goldens.flatMap { g =>
      val text = g.stderr.toList.flatMap(_.linesIterator).collect { case CodeLine(code) => code }
      val json = envelopeOf(g).flatMap(errorCode).toList
      (text ++ json).map(code => (g.name, code, recordedExit(g)))
    }

  /** (golden, code) for every warning the corpus records, on either channel. */
  private lazy val recordedWarnings: Vector[(String, String)] =
    goldens.flatMap { g =>
      val text =
        g.stderr.toList.flatMap(_.linesIterator).collect { case WarningLine(code) => code }
      val json = envelopeOf(g).toList.flatMap { env =>
        field(env, "warnings").toList.flatMap {
          case arr: ujson.Arr =>
            arr.value.toList.collect { case w: ujson.Obj => field(w, "code") }.flatten.collect {
              case ujson.Str(code) => code
            }
          case _ => Nil
        }
      }
      (text ++ json).map(code => (g.name, code))
    }

  test("the corpus is present, every golden parses and records its exit") {
    assert(Golden.caseFiles.nonEmpty, s"no *.golden files under ${Golden.corpusDir}")
    assertEquals(parsed.collect { case Left(msg) => msg }, Vector.empty[String])
    goldens.foreach(g => assert(g.exit.isDefined, s"${g.name}.golden: exit not recorded"))
  }

  test(
    "every --json golden is an envelope: exitCode is the recorded exit, error is coded or null"
  ) {
    assert(envelopes.nonEmpty, "no --json golden")
    envelopes.foreach { (g, env) =>
      assertEquals(
        field(env, "exitCode").collect { case ujson.Num(n) => n.toInt },
        Some(recordedExit(g)),
        s"${g.name}.golden: envelope exitCode"
      )
      field(env, "error") match
        case Some(ujson.Null) | None => ()
        case Some(err: ujson.Obj) =>
          assert(errorCode(env).isDefined, s"${g.name}.golden: error without a string code")
        case Some(other) => fail(s"${g.name}.golden: error is ${ujson.write(other)}")
    }
  }

  test("the corpus records errors and warnings, on the text channel and in envelopes") {
    assert(recordedErrors.nonEmpty, "no error recorded in any golden")
    assert(recordedWarnings.nonEmpty, "no warning recorded in any golden")
    assert(envelopes.exists((_, env) => errorCode(env).isDefined), "no --json golden that failed")
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
    goldens.filter(g => Set(2, 3).contains(recordedExit(g))).foreach { g =>
      assert(withCode.contains(g.name), s"${g.name}.golden exits ${g.exit} without a code")
    }
  }

  test("every --json golden that exits 1 carries error.code (findings are addressable by code)") {
    envelopes.filter((g, _) => recordedExit(g) == 1).foreach { (g, env) =>
      errorCode(env) match
        case Some(code) => assert(published.contains(code), s"${g.name}.golden: unpublished $code")
        case None => fail(s"${g.name}.golden: exit 1 envelope without error.code")
    }
  }

  test("the corpus exercises each exit class the table defines (the scans above are not vacuous)") {
    assertEquals(recordedErrors.map(_._3).toSet, Set(1, 2, 3))
  }
