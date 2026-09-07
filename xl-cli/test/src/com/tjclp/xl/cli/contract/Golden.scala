package com.tjclp.xl.cli.contract

import java.nio.file.{Files, Path, Paths}
import java.util.regex.Pattern

import scala.annotation.tailrec

import com.tjclp.xl.cli.BuildInfo

/**
 * One case of the CLI contract corpus: `xl-cli/test/resources/golden/<name>.golden`.
 *
 * File layout, one section per `## ` header, in this order:
 * {{{
 * ## args        one argv token per line; `<DIR>` is the per-run fixture directory
 * ## stdin       (optional) text fed to standard input
 * ## exit        the exit code
 * ## stdout      normalized standard output
 * ## stderr      normalized standard error
 * }}}
 * A file may carry only `## args` (and `## stdin`): recording with `XL_UPDATE_GOLDEN=1` fills in
 * the rest. Output is normalized by [[Golden.normalize]] before it is compared or written.
 */
final case class GoldenCase(
  name: String,
  args: List[String],
  stdin: Option[String],
  exit: Option[Int],
  stdout: Option[String],
  stderr: Option[String]
):
  def observed(exit: Int, stdout: String, stderr: String): GoldenCase =
    copy(exit = Some(exit), stdout = Some(stdout), stderr = Some(stderr))

object GoldenCase:

  private val headers = Set("args", "stdin", "exit", "stdout", "stderr")

  /** Fold state while scanning a golden file line by line. */
  private final case class Scan(
    sections: Map[String, Vector[String]],
    current: Option[String],
    error: Option[String]
  )

  def parse(name: String, text: String): Either[String, GoldenCase] =
    val scanned =
      text.split("\n", -1).toVector.foldLeft(Scan(Map.empty, None, None)) { (scan, line) =>
        val header = line.stripPrefix("## ")
        if scan.error.isDefined then scan
        else if line.startsWith("## ") && headers.contains(header) then
          scan.copy(sections = scan.sections.updated(header, Vector.empty), current = Some(header))
        else
          scan.current match
            case Some(section) =>
              val body = scan.sections.getOrElse(section, Vector.empty) :+ line
              scan.copy(sections = scan.sections.updated(section, body))
            case None if line.trim.isEmpty => scan
            case None =>
              scan.copy(error = Some(s"$name.golden: text before the first '## ' header: $line"))
      }
    scanned.error.toLeft(scanned.sections).flatMap { sections =>
      def body(section: String): Option[String] =
        sections.get(section).map(v => Golden.trimTrailing(v.mkString("\n")))
      val exit: Either[String, Option[Int]] = body("exit").map(_.trim) match
        case None => Right(None)
        case Some(raw) =>
          raw.toIntOption.toRight(s"$name.golden: '## exit' is not an integer: $raw").map(Some(_))
      for
        argLines <- sections.get("args").toRight(s"$name.golden: missing '## args' section")
        code <- exit
      yield GoldenCase(
        name,
        Golden.dropTrailingBlank(argLines).toList,
        sections.get("stdin").map(v => Golden.dropTrailingBlank(v).mkString("", "\n", "\n")),
        code,
        body("stdout"),
        body("stderr")
      )
    }

  def render(c: GoldenCase): String =
    val sb = new StringBuilder
    def section(header: String, body: String): Unit =
      sb.append("## ").append(header).append('\n')
      if body.nonEmpty then sb.append(body).append('\n')
    section("args", c.args.mkString("\n"))
    c.stdin.foreach(text => section("stdin", Golden.trimTrailing(text)))
    c.exit.foreach(code => section("exit", code.toString))
    c.stdout.foreach(text => section("stdout", text))
    c.stderr.foreach(text => section("stderr", text))
    sb.toString

object Golden:

  /** The checkout root: the nearest ancestor of the working directory holding `build.mill`. */
  lazy val repoRoot: Path =
    val start = Paths.get(sys.props.getOrElse("user.dir", ".")).toAbsolutePath
    Iterator
      .iterate(Option(start))(_.flatMap(p => Option(p.getParent)))
      .takeWhile(_.isDefined)
      .flatten
      .find(p => Files.isRegularFile(p.resolve("build.mill")))
      .getOrElse(
        throw new IllegalStateException(s"no build.mill at or above $start; cannot locate goldens")
      )

  val corpusDir: Path = repoRoot.resolve("xl-cli/test/resources/golden")

  /**
   * Make output byte-stable across machines and runs: the fixture directory becomes `<DIR>`, any
   * other path under the JVM temp directory becomes `<TMP>`, the build version becomes `<VERSION>`,
   * trailing whitespace is dropped from every line and trailing blank lines from the whole text.
   * Exactly what is written to a golden and what an actual run is reduced to.
   */
  def normalize(text: String, dir: Path): String =
    val realDir = scala.util.Try(dir.toRealPath()).map(_.toString).getOrElse(dir.toString)
    val tmp = Paths.get(sys.props.getOrElse("java.io.tmpdir", "/tmp")).toAbsolutePath.toString
    val tmpPath = Pattern.compile(Pattern.quote(tmp) + "/[^\\s\"'`|]*")
    val withDir = text.replace(dir.toString, "<DIR>").replace(realDir, "<DIR>")
    val withTmp = tmpPath.matcher(withDir).replaceAll("<TMP>")
    trimTrailing(withTmp.replace(BuildInfo.version, "<VERSION>"))

  def trimTrailing(text: String): String =
    dropTrailingBlank(text.split("\n", -1).toVector.map(_.replaceAll("\\s+$", ""))).mkString("\n")

  def dropTrailingBlank(lines: Vector[String]): Vector[String] =
    lines.reverse.dropWhile(_.isEmpty).reverse

/** Dependency-free line diff in unified format, for golden mismatch messages. */
object UnifiedDiff:

  private enum Op derives CanEqual:
    case Keep(line: String)
    case Del(line: String)
    case Ins(line: String)

    /** Whether this op consumes a line of the expected text. */
    def inExpected: Boolean = this match
      case Ins(_) => false
      case _ => true

    /** Whether this op consumes a line of the actual text. */
    def inActual: Boolean = this match
      case Del(_) => false
      case _ => true

  def render(expected: String, actual: String, context: Int = 3): String =
    val ops = diffOps(expected.split("\n", -1).toVector, actual.split("\n", -1).toVector)
    val changes = ops.indices.filter(i => !(ops(i).inExpected && ops(i).inActual)).toVector
    if changes.isEmpty then ""
    else
      val shown =
        changes.flatMap(i => (i - context) to (i + context)).filter(i => i >= 0 && i < ops.length)
      val hunks = shown.distinct.sorted.foldLeft(Vector.empty[Vector[Int]]) { (acc, i) =>
        acc.lastOption match
          case Some(group) if group.lastOption.contains(i - 1) => acc.dropRight(1) :+ (group :+ i)
          case _ => acc :+ Vector(i)
      }
      val positions = ops.scanLeft((1, 1)) { case ((e, a), op) =>
        op match
          case Op.Keep(_) => (e + 1, a + 1)
          case Op.Del(_) => (e + 1, a)
          case Op.Ins(_) => (e, a + 1)
      }
      val rendered = hunks.map { group =>
        val (eStart, aStart) = positions(group(0))
        val eCount = group.count(i => ops(i).inExpected)
        val aCount = group.count(i => ops(i).inActual)
        val body = group.map(i =>
          ops(i) match
            case Op.Keep(l) => s" $l"
            case Op.Del(l) => s"-$l"
            case Op.Ins(l) => s"+$l"
        )
        (s"@@ -$eStart,$eCount +$aStart,$aCount @@" +: body).mkString("\n")
      }
      ("--- expected" +: "+++ actual" +: rendered).mkString("\n")

  private def diffOps(a: Vector[String], b: Vector[String]): Vector[Op] =
    val n = a.length
    val m = b.length
    val lcs = Array.ofDim[Int](n + 1, m + 1)
    for i <- (n - 1) to 0 by -1 do
      for j <- (m - 1) to 0 by -1 do
        lcs(i)(j) =
          if a(i) == b(j) then lcs(i + 1)(j + 1) + 1
          else math.max(lcs(i + 1)(j), lcs(i)(j + 1))
    // On a tie, drop the expected line before inserting the actual one: `-` then `+`, as diff does.
    @tailrec
    def walk(i: Int, j: Int, acc: Vector[Op]): Vector[Op] =
      if i < n && j < m && a(i) == b(j) then walk(i + 1, j + 1, acc :+ Op.Keep(a(i)))
      else if i < n && (j >= m || lcs(i + 1)(j) >= lcs(i)(j + 1)) then
        walk(i + 1, j, acc :+ Op.Del(a(i)))
      else if j < m then walk(i, j + 1, acc :+ Op.Ins(b(j)))
      else acc
    walk(0, 0, Vector.empty)
