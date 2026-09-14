package com.tjclp.xl.cli

import munit.FunSuite

import java.io.File
import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.concurrent.TimeUnit

import scala.concurrent.duration.*
import scala.jdk.CollectionConverters.*

import cats.effect.{IO, unsafe}
import com.tjclp.xl.{Sheet, Workbook, given}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.writer.WriterConfig

/**
 * LibreOffice as the DISPLAYED-VALUE oracle for `--no-recalc` structural writes (review of #509).
 *
 * LibreOffice's shipped default for xlsx ("Recalculation on File Load: never") ignores
 * `<calcPr fullCalcOnLoad="1"/>`: it displays whatever `<v>` a cell carries and computes only the
 * cells that have none. `xl view --eval` cannot stand in for it — the question is what a reader
 * that trusts the file shows, so the file is opened by one. Every test here is skipped, with a
 * message, when `soffice` is not on `PATH`.
 */
class LibreOfficeOracleSpec extends FunSuite:

  given unsafe.IORuntime = unsafe.IORuntime.global

  override def munitTimeout: Duration = 5.minutes

  private val config: WriterConfig = WriterConfig.default

  /** `soffice` on PATH, if any (the macOS app bundle is not searched: PATH is the contract). */
  private val soffice: Option[Path] =
    sys.env
      .get("PATH")
      .toList
      .flatMap(_.split(File.pathSeparator).toList)
      .filter(_.nonEmpty)
      .map(dir => Path.of(dir).resolve("soffice"))
      .find(p => Files.isRegularFile(p) && Files.isExecutable(p))

  /**
   * Convert every given xlsx to CSV in one headless LibreOffice run under a fresh user profile
   * (factory-default recalculation settings, no interference from a desktop instance) and return
   * each file's CSV text, keyed by its base name.
   */
  private def libreOfficeCsv(workDir: Path, books: Path*): Map[String, String] =
    soffice match
      case None => fail("premise: soffice must be on PATH (the test is gated on it)")
      case Some(bin) =>
        val profile = workDir.resolve("lo-profile")
        val outDir = workDir.resolve("csv")
        Files.createDirectories(outDir)
        val command = List(
          bin.toString,
          s"-env:UserInstallation=${profile.toUri}",
          "--headless",
          "--convert-to",
          "csv",
          "--outdir",
          outDir.toString
        ) ++ books.map(_.toString)
        val process = new ProcessBuilder(command.asJava).redirectErrorStream(true).start()
        val output = new String(process.getInputStream.readAllBytes(), StandardCharsets.UTF_8)
        val finished = process.waitFor(4, TimeUnit.MINUTES)
        if !finished then
          process.destroyForcibly()
          fail(s"soffice did not finish within 4 minutes; output so far:\n$output")
        assertEquals(process.exitValue(), 0, s"soffice failed:\n$output")
        books.map { book =>
          val name = book.getFileName.toString.stripSuffix(".xlsx")
          val csv = outDir.resolve(s"$name.csv")
          assert(Files.isRegularFile(csv), s"soffice wrote no $csv; output:\n$output")
          name -> new String(Files.readAllBytes(csv), StandardCharsets.UTF_8).trim
        }.toMap

  private def num(n: Int): Option[CellValue] = Some(CellValue.Number(BigDecimal(n)))

  /**
   * The review's reproduction: A1=5, A2=10, B1=SUM(A1:A2) cached 15, B2=ROW()*10 cached 20,
   * C1=B1+B2 cached 35. Deleting row 2 makes B1 `SUM(A1:A1)` = 5 and C1 `B1+#REF!` = #REF!.
   */
  private def reviewWorkbook(): Workbook =
    Workbook(
      Sheet("Sheet1")
        .put(ref"A1" -> 5, ref"A2" -> 10)
        .put(ref"B1", CellValue.Formula("SUM(A1:A2)", num(15)))
        .put(ref"B2", CellValue.Formula("ROW()*10", num(20)))
        .put(ref"C1", CellValue.Formula("B1+B2", num(35)))
    )

  test(
    "--no-recalc delete-rows: LibreOffice displays no stale number (matches the recalculated write)"
  ) {
    assume(soffice.isDefined, "soffice not on PATH - skipping LibreOffice oracle test")
    val workDir = Files.createTempDirectory("xl-lo-oracle")
    val noRecalc = workDir.resolve("no-recalc.xlsx")
    val recalculated = workDir.resolve("recalculated.xlsx")
    try
      val wb = reviewWorkbook()
      val summary = WriteCommands
        .deleteRows(
          wb,
          wb.sheets.headOption,
          2,
          1,
          noRecalc,
          config,
          false,
          WritePolicy(noRecalc = true)
        )
        .unsafeRunSync()
      WriteCommands.deleteRows(wb, wb.sheets.headOption, 2, 1, recalculated, config).unsafeRunSync()
      assert(summary.contains("fullCalcOnLoad"), s"premise: the marker is set: $summary")

      val csv = libreOfficeCsv(workDir, noRecalc, recalculated)
      val truth = csv("recalculated")
      assertEquals(truth, "5,5,#REF!", "premise: LibreOffice shows the recalculated write's values")
      assertEquals(
        csv("no-recalc"),
        truth,
        "LibreOffice must display the same values for the --no-recalc write as for the " +
          "recalculated one: a carried pre-edit <v> is displayed as-is, fullCalcOnLoad or not"
      )
      assert(
        !csv("no-recalc").contains("15"),
        s"stale SUM(A1:A2) cache displayed: ${csv("no-recalc")}"
      )
      assert(!csv("no-recalc").contains("35"), s"stale B1+B2 cache displayed: ${csv("no-recalc")}")
    finally
      Files.walk(workDir).iterator().asScala.toList.reverse.foreach(p => Files.deleteIfExists(p))
  }
