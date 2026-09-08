// GH-589: the failure path of `orExit` ends the JVM, so it is pinned from outside — a child JVM
// on this test classpath runs a script-shaped main whose first step is a Left, and the parent
// asserts exit status 1, an empty stdout and the CLI's exact envelope on stderr.
package xlprelude

import java.nio.charset.StandardCharsets.UTF_8
import java.nio.file.{Files, Paths}

import munit.FunSuite

import com.tjclp.xl.scripting.{*, given}

/** What a script does at its edge: unwrap, or print the diagnostic and exit 1. */
object OrExitSmokeMain:
  val failure: XLError = XLError.SheetNotFound("Sumary", Vector("Data", "Summary"))

  def main(args: Array[String]): Unit =
    val sheet: Sheet = orExit(Left(failure))
    println(s"unreachable: ${sheet.name.value}")

/**
 * GH-615: the documented script shape — `Excel.readSheet` on a misspelled sheet, unwrapped by
 * `orExit` — against a real file: the candidates come from the workbook the reader loaded.
 */
object OrExitReadSheetMain:
  def main(args: Array[String]): Unit =
    val path = args.headOption.getOrElse(sys.error("path required"))
    Excel.write(Workbook(Sheet("Data").put(ref"A1", 1), Sheet("Summary").put(ref"A1", 2)), path)
    val sheet: Sheet = orExit(Excel.readSheet(path, "Sumary"))
    println(s"unreachable: ${sheet.name.value}")

class OrExitSmokeTest extends FunSuite:

  test("orExit on Left prints exitMessage to stderr, nothing to stdout, and exits 1"):
    val java = Paths.get(System.getProperty("java.home"), "bin", "java").toString
    val classpath = System.getProperty("java.class.path")
    val dir = Files.createTempDirectory("xl-orexit-smoke")
    val out = dir.resolve("stdout.txt")
    val err = dir.resolve("stderr.txt")
    try
      val process = new ProcessBuilder(java, "-cp", classpath, "xlprelude.OrExitSmokeMain")
        .redirectOutput(out.toFile)
        .redirectError(err.toFile)
        .start()
      val status = process.waitFor()
      val stderr = Files.readString(err, UTF_8)
      assertEquals(status, 1, stderr)
      assertEquals(Files.readString(out, UTF_8), "")
      assertEquals(stderr.stripTrailing, exitMessage(OrExitSmokeMain.failure))
      assertEquals(
        stderr.stripTrailing,
        """Error: Sheet not found: 'Sumary'. Available: Data, Summary
          |  code: SHEET_NOT_FOUND
          |  did you mean: Summary
          |  hint: list sheets with `xl -f <file> sheets`""".stripMargin
      )
    finally
      Files.deleteIfExists(out)
      Files.deleteIfExists(err)
      Files.deleteIfExists(dir)

  test("GH-615: orExit(Excel.readSheet(path, misspelled)) prints did-you-mean from the file"):
    val java = Paths.get(System.getProperty("java.home"), "bin", "java").toString
    val classpath = System.getProperty("java.class.path")
    val dir = Files.createTempDirectory("xl-orexit-readsheet")
    val book = dir.resolve("book.xlsx")
    val out = dir.resolve("stdout.txt")
    val err = dir.resolve("stderr.txt")
    try
      val process =
        new ProcessBuilder(java, "-cp", classpath, "xlprelude.OrExitReadSheetMain", book.toString)
          .redirectOutput(out.toFile)
          .redirectError(err.toFile)
          .start()
      val status = process.waitFor()
      // the JVM's own sun.misc.Unsafe deprecation notice (scala.runtime.LazyVals) is not ours
      val stderr = Files
        .readString(err, UTF_8)
        .linesIterator
        .filterNot(_.startsWith("WARNING:"))
        .mkString("\n")
      assertEquals(status, 1, stderr)
      assertEquals(Files.readString(out, UTF_8), "")
      assertEquals(
        stderr.stripTrailing,
        """Error: Sheet not found: 'Sumary'. Available: Data, Summary
          |  code: SHEET_NOT_FOUND
          |  did you mean: Summary
          |  hint: list sheets with `xl -f <file> sheets`""".stripMargin
      )
    finally
      Files.deleteIfExists(book)
      Files.deleteIfExists(out)
      Files.deleteIfExists(err)
      Files.deleteIfExists(dir)
