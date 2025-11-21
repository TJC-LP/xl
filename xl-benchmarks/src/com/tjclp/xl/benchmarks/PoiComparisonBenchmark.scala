package com.tjclp.xl.benchmarks

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.workbooks.Workbook
import org.apache.poi.ss.usermodel.{
  Cell,
  CellType,
  Row,
  Sheet,
  Workbook as PoiWorkbook,
  WorkbookFactory
}
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.openjdk.jmh.annotations.*
import java.nio.file.{Files, Path}
import java.util.concurrent.TimeUnit
import scala.compiletime.uninitialized

/**
 * Comparison benchmarks: XL vs Apache POI.
 *
 * Tests equivalent operations:
 *   - Write workbook with N rows
 *   - Read workbook with N rows
 *   - Memory footprint (via GC profiler)
 *
 * Fair comparison: same data structures, same operations.
 */
@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.AverageTime))
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 3, time = 1, timeUnit = TimeUnit.SECONDS)
@Measurement(iterations = 5, time = 2, timeUnit = TimeUnit.SECONDS)
@Fork(value = 1)
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.Null"))
class PoiComparisonBenchmark {

  @Param(Array("1000", "10000"))
  var rows: Int = uninitialized

  var xlWorkbook: Workbook = uninitialized
  var poiWorkbook: PoiWorkbook = uninitialized
  var sharedTestFile: Path = uninitialized // Shared file for fair comparison

  @Setup(Level.Trial)
  def setup(): Unit = {
    // Generate verifiable workbook (cell A{i} = i for arithmetic series sum validation)
    xlWorkbook = BenchmarkUtils.createVerifiableWorkbook(rows)

    // Generate equivalent POI workbook with same verifiable data
    poiWorkbook = new XSSFWorkbook()
    val sheet = poiWorkbook.createSheet("Data")

    // Data rows (cell A{i} = i)
    for (i <- 1 to rows) {
      val row = sheet.createRow(i - 1) // 0-indexed in POI
      row.createCell(0, CellType.NUMERIC).setCellValue(i.toDouble)
    }

    // Create ONE shared temp file for fair comparison
    sharedTestFile = Files.createTempFile("shared-benchmark-", ".xlsx")

    // PRE-CREATE file for read benchmarks (write cost separated from read measurements)
    // Use XL to write (arbitrary choice - both libraries must read same file)
    ExcelIO.instance[IO].write(xlWorkbook, sharedTestFile).unsafeRunSync()
  }

  @TearDown(Level.Trial)
  def teardown(): Unit = {
    if (sharedTestFile != null && Files.exists(sharedTestFile)) Files.delete(sharedTestFile)
    if (poiWorkbook != null) poiWorkbook.close()
  }

  @Benchmark
  def xlWrite(): Unit = {
    // XL write performance
    ExcelIO.instance[IO].write(xlWorkbook, sharedTestFile).unsafeRunSync()
  }

  @Benchmark
  def poiWrite(): Unit = {
    // POI write performance
    val fos = new java.io.FileOutputStream(sharedTestFile.toFile)
    try {
      poiWorkbook.write(fos)
    } finally {
      fos.close()
    }
  }

  @Benchmark
  def xlRead(): Workbook = {
    // XL in-memory read performance (file pre-created in setup)
    ExcelIO.instance[IO].read(sharedTestFile).unsafeRunSync()
  }

  @Benchmark
  def poiRead(): PoiWorkbook = {
    // POI in-memory read performance (file pre-created in setup)
    val fis = new java.io.FileInputStream(sharedTestFile.toFile)
    try {
      WorkbookFactory.create(fis)
    } finally {
      fis.close()
    }
  }

  @Benchmark
  def xlReadStream(): Double = {
    // XL streaming read performance - constant memory with verifiable sum computation
    // Computes sum of all numeric values in column A (index 0)
    // Expected sum: N × (N+1) / 2 for N rows (arithmetic series)
    ExcelIO
      .instance[IO]
      .readStream(sharedTestFile)
      .evalMap { rowData =>
        IO.pure {
          // Extract numeric value from column A (index 0)
          rowData.cells.get(0) match {
            case Some(CellValue.Number(n)) => n.toDouble
            case _ => 0.0
          }
        }
      }
      .compile
      .fold(0.0)(_ + _)
      .unsafeRunSync()
  }

  @Benchmark
  def poiReadStream(): Double = {
    // POI streaming read performance - constant memory via XSSFReader (file pre-created in setup)
    // Computes sum of all numeric values in column A
    import org.apache.poi.xssf.eventusermodel.{XSSFReader, XSSFSheetXMLHandler}
    import org.apache.poi.xssf.model.SharedStrings
    import org.apache.poi.openxml4j.opc.OPCPackage
    import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler
    import org.xml.sax.InputSource
    import javax.xml.parsers.SAXParserFactory

    // Stream read with XSSFReader
    val pkg = OPCPackage.open(sharedTestFile.toFile)
    try {
      val reader = new XSSFReader(pkg)
      val sst = reader.getSharedStringsTable()

      // Sum accumulator handler
      var sum = 0.0
      val handler = new SheetContentsHandler {
        override def startRow(rowNum: Int): Unit = ()
        override def endRow(rowNum: Int): Unit = ()
        override def cell(
          cellReference: String,
          formattedValue: String,
          comment: org.apache.poi.xssf.usermodel.XSSFComment
        ): Unit = {
          // Only accumulate values from column A
          if (cellReference.startsWith("A")) {
            try {
              sum += formattedValue.toDouble
            } catch {
              case _: NumberFormatException => () // Skip non-numeric values
            }
          }
        }
        override def headerFooter(text: String, isHeader: Boolean, tagName: String): Unit = ()
      }

      val sheetHandler = new XSSFSheetXMLHandler(reader.getStylesTable(), sst, handler, false)
      val parser = SAXParserFactory.newInstance().newSAXParser()
      val xmlReader = parser.getXMLReader
      xmlReader.setContentHandler(sheetHandler)

      val sheetsIter = reader.getSheetsData()
      // WartRemover: while acceptable in benchmark code for POI interop
      while (sheetsIter.hasNext) {
        val sheetStream = sheetsIter.next()
        xmlReader.parse(new InputSource(sheetStream))
        sheetStream.close()
      }

      sum
    } finally {
      pkg.close()
    }
  }
}
