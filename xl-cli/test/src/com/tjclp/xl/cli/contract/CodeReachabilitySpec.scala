package com.tjclp.xl.cli.contract

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import scala.jdk.CollectionConverters.*
import scala.util.matching.Regex

import munit.FunSuite

/**
 * Every published code is one the CLI can actually emit (ADR-017 §2.3): each constant in
 * [[ErrorCode.cli]] and [[WarningCode.all]] has at least one raising site under `xl-cli/src` — a
 * `ErrorCode.<CODE>` / `WarningCode.<CODE>` reference outside the files that merely define or
 * tabulate the vocabulary. A code with no site is a promise the schema makes and the binary cannot
 * keep; a deliberate reservation goes in [[CodeReachabilitySpec.reserved]] with its reason, and the
 * spec also fails when a reserved code gains a site, so the list cannot go stale.
 */
class CodeReachabilitySpec extends FunSuite:

  /** Codes published ahead of their emission site, and why. */
  private val reserved: Map[String, String] = Map(
    WarningCode.UNKNOWN_PROPERTY -> "wired by the batch warning sink",
    WarningCode.FORMAT_HINT_IGNORED -> "wired by the batch warning sink"
  )

  private val sourceRoot: Path = Golden.repoRoot.resolve("xl-cli/src")

  /** The vocabulary's own files: definitions and tables, never emissions. */
  private val vocabulary: Set[String] = Set(
    "contract/ErrorCode.scala",
    "contract/Warning.scala",
    "contract/ExitCodes.scala",
    "contract/Schema.scala"
  )

  private lazy val sources: Vector[(String, String)] =
    val walk = Files.walk(sourceRoot)
    try
      walk.iterator.asScala
        .filter(p => p.toString.endsWith(".scala"))
        .map(p => (sourceRoot.relativize(p).toString.replace('\\', '/'), read(p)))
        .filterNot((name, _) => vocabulary.exists(name.endsWith))
        .toVector
    finally walk.close()

  private def read(p: Path): String = Files.readString(p, StandardCharsets.UTF_8)

  private def sites(owner: String, code: String): Vector[String] =
    val pattern = new Regex(s"\\b$owner\\.$code\\b")
    sources.collect { case (name, text) if pattern.findFirstIn(text).isDefined => name }

  private def check(owner: String, codes: Vector[String]): Unit =
    codes.foreach { code =>
      val where = sites(owner, code)
      reserved.get(code) match
        case Some(reason) =>
          assert(
            where.isEmpty,
            s"$owner.$code is reserved ($reason) but now has a site: ${where.mkString(", ")} — drop the reservation"
          )
        case None =>
          assert(
            where.nonEmpty,
            s"$owner.$code is published but nothing under xl-cli/src raises it (reserve it with a reason, or remove it)"
          )
    }

  test("the source tree is where we think it is") {
    assert(Files.isDirectory(sourceRoot), sourceRoot.toString)
    assert(sources.nonEmpty, "no Scala sources found")
  }

  test("every CLI error code has an emission site (or a stated reservation)") {
    check("ErrorCode", ErrorCode.cli)
  }

  test("every warning code has an emission site (or a stated reservation)") {
    check("WarningCode", WarningCode.all)
  }

  test("every reservation names a published code") {
    reserved.keys.foreach { code =>
      assert(
        ErrorCode.cli.contains(code) || WarningCode.all.contains(code),
        s"reserved code $code is not published"
      )
    }
  }
