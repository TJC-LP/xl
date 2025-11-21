package com.tjclp.xl.benchmarks

import cats.effect.IO
import cats.effect.unsafe.implicits.global
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
  var xlTempFile: Path = uninitialized
  var poiTempFile: Path = uninitialized

  @Setup(Level.Trial)
  def setup(): Unit = {
    // Generate XL workbook
    xlWorkbook = BenchmarkUtils.generateWorkbook(rows, styled = false)

    // Generate equivalent POI workbook
    poiWorkbook = new XSSFWorkbook()
    val sheet = poiWorkbook.createSheet("Data")

    // Header row
    val headerRow = sheet.createRow(0)
    headerRow.createCell(0, CellType.STRING).setCellValue("ID")
    headerRow.createCell(1, CellType.STRING).setCellValue("Name")
    headerRow.createCell(2, CellType.STRING).setCellValue("Date")
    headerRow.createCell(3, CellType.STRING).setCellValue("Amount")
    headerRow.createCell(4, CellType.BOOLEAN).setCellValue("Active")

    // Data rows (equivalent to BenchmarkUtils.generateSheet)
    val random = new scala.util.Random(42)
    for (i <- 1 to rows) {
      val row = sheet.createRow(i)
      row.createCell(0, CellType.NUMERIC).setCellValue(i.toDouble)
      row.createCell(1, CellType.STRING).setCellValue(s"User_${i % 100}")
      row
        .createCell(2, CellType.NUMERIC)
        .setCellValue(java.time.LocalDate.of(2024, 1, 1).plusDays(i % 365).toEpochDay.toDouble)
      row.createCell(3, CellType.NUMERIC).setCellValue(random.nextDouble() * 10000)
      row.createCell(4, CellType.BOOLEAN).setCellValue(random.nextBoolean())
    }

    // Create temp files
    xlTempFile = Files.createTempFile("xl-poi-benchmark-", ".xlsx")
    poiTempFile = Files.createTempFile("poi-benchmark-", ".xlsx")
  }

  @TearDown(Level.Trial)
  def teardown(): Unit = {
    if (xlTempFile != null && Files.exists(xlTempFile)) Files.delete(xlTempFile)
    if (poiTempFile != null && Files.exists(poiTempFile)) Files.delete(poiTempFile)
    if (poiWorkbook != null) poiWorkbook.close()
  }

  @Benchmark
  def xlWrite(): Unit = {
    // XL write performance
    ExcelIO.instance[IO].write(xlWorkbook, xlTempFile).unsafeRunSync()
  }

  @Benchmark
  def poiWrite(): Unit = {
    // POI write performance
    val fos = new java.io.FileOutputStream(poiTempFile.toFile)
    try {
      poiWorkbook.write(fos)
    } finally {
      fos.close()
    }
  }

  @Benchmark
  def xlRead(): Workbook = {
    // XL read performance
    ExcelIO.instance[IO].write(xlWorkbook, xlTempFile).unsafeRunSync()
    ExcelIO.instance[IO].read(xlTempFile).unsafeRunSync()
  }

  @Benchmark
  def poiRead(): PoiWorkbook = {
    // POI read performance (in-memory)
    val fos = new java.io.FileOutputStream(poiTempFile.toFile)
    try {
      poiWorkbook.write(fos)
    } finally {
      fos.close()
    }

    val fis = new java.io.FileInputStream(poiTempFile.toFile)
    try {
      WorkbookFactory.create(fis)
    } finally {
      fis.close()
    }
  }

  @Benchmark
  def xlReadStream(): Long = {
    // XL streaming read performance (constant memory)
    ExcelIO.instance[IO].write(xlWorkbook, xlTempFile).unsafeRunSync()
    ExcelIO.instance[IO].readStream(xlTempFile).compile.count.unsafeRunSync()
  }

  @Benchmark
  def poiReadStream(): Long = {
    // POI streaming read performance (constant memory via XSSFReader)
    import org.apache.poi.xssf.eventusermodel.{XSSFReader, XSSFSheetXMLHandler}
    import org.apache.poi.xssf.model.SharedStrings
    import org.apache.poi.openxml4j.opc.OPCPackage
    import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler
    import org.xml.sax.InputSource
    import javax.xml.parsers.SAXParserFactory

    // Write file first
    val fos = new java.io.FileOutputStream(poiTempFile.toFile)
    try {
      poiWorkbook.write(fos)
    } finally {
      fos.close()
    }

    // Stream read with XSSFReader
    val pkg = OPCPackage.open(poiTempFile.toFile)
    try {
      val reader = new XSSFReader(pkg)
      val sst = reader.getSharedStringsTable()

      // Row counter handler
      var rowCount = 0L
      val handler = new SheetContentsHandler {
        override def startRow(rowNum: Int): Unit = rowCount += 1
        override def endRow(rowNum: Int): Unit = ()
        override def cell(
          cellReference: String,
          formattedValue: String,
          comment: org.apache.poi.xssf.usermodel.XSSFComment
        ): Unit = ()
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

      rowCount
    } finally {
      pkg.close()
    }
  }
}
