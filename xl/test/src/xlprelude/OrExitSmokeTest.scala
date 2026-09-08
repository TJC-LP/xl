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
