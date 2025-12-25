package com.tjclp.xl.benchmarks

import cats.effect.IO
import cats.effect.unsafe.implicits.global
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.addressing.Column
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
import scala.jdk.CollectionConverters.*
import scala.collection.mutable
import scala.compiletime.uninitialized
import javax.xml.parsers.SAXParserFactory
import org.xml.sax.InputSource
import org.xml.sax.helpers.DefaultHandler
import org.xml.sax.Attributes

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

  @Param(Array("1000", "10000", "100000"))
  var rows: Int = uninitialized

  var xlWorkbook: Workbook = uninitialized
  var poiWorkbook: PoiWorkbook = uninitialized
  var expectedSum: Double = uninitialized

  // Base files (single source of truth for copies)
  var xlBaseFile: Path = uninitialized
  var poiBaseFile: Path = uninitialized

  // In-memory read targets (copies to avoid shared OS cache)
  var xlReadFromXl: Path = uninitialized
  var poiReadFromXlFile: Path = uninitialized
  var xlReadFromPoiFile: Path = uninitialized
  var poiReadFromPoiFile: Path = uninitialized

  // Streaming read targets (copies to avoid shared OS cache)
  var xlStreamFromXl: Path = uninitialized
  var poiStreamFromXlFile: Path = uninitialized
  var xlStreamFromPoiFile: Path = uninitialized
  var poiStreamFromPoiFile: Path = uninitialized

  // Write targets
  var xlWriteFile: Path = uninitialized
  var poiWriteFile: Path = uninitialized

  @Setup(Level.Trial)
  def setup(): Unit = {
    // Generate verifiable workbook (cell A{i} = i for arithmetic series sum validation)
    xlWorkbook = BenchmarkUtils.createVerifiableWorkbook(rows)
    expectedSum = BenchmarkUtils.expectedSum(rows)
    // Validate in-memory data before writing
    requireCloseEnough(sumFirstColumn(xlWorkbook), expectedSum, "xlWorkbook construction failed")

    // Generate equivalent POI workbook with same verifiable data
    poiWorkbook = new XSSFWorkbook()
    val sheet = poiWorkbook.createSheet("Data")

    // Data rows (cell A{i} = i)
    for (i <- 1 to rows) {
      val row = sheet.createRow(i - 1) // 0-indexed in POI
      row.createCell(0, CellType.NUMERIC).setCellValue(i.toDouble)
    }

    // Base files
    xlBaseFile = Files.createTempFile("xl-base-", ".xlsx")
    ExcelIO.instance[IO].write(xlWorkbook, xlBaseFile).unsafeRunSync()
    val xlWritten = ExcelIO.instance[IO].read(xlBaseFile).unsafeRunSync()
    requireCloseEnough(sumFirstColumn(xlWritten), expectedSum, "xlBaseFile validation failed")

    poiBaseFile = Files.createTempFile("poi-base-", ".xlsx")
    val fos = new java.io.FileOutputStream(poiBaseFile.toFile)
    try {
      poiWorkbook.write(fos)
      fos.flush()
    } finally {
      fos.close()
    }
    // Validate POI base with POI
    val poiBack = {
      val fis = new java.io.FileInputStream(poiBaseFile.toFile)
      try WorkbookFactory.create(fis)
      finally fis.close()
    }
    try {
      val sheet = poiBack.getSheetAt(0)
      val sum = sheet.iterator().asScala.foldLeft(0.0) { (acc, row) =>
        val cell = row.getCell(0)
        if cell != null then acc + cell.getNumericCellValue else acc
      }
      requireCloseEnough(sum, expectedSum, "poiBaseFile validation failed")
    } finally poiBack.close()

    // Copies for in-memory read benchmarks (avoid cache contamination)
    xlReadFromXl = copyTemp(xlBaseFile, "xl-read-from-xl-")
    poiReadFromXlFile = copyTemp(xlBaseFile, "poi-read-from-xl-")
    xlReadFromPoiFile = copyTemp(poiBaseFile, "xl-read-from-poi-")
    poiReadFromPoiFile = copyTemp(poiBaseFile, "poi-read-from-poi-")

    // Copies for streaming read benchmarks (avoid cache contamination)
    xlStreamFromXl = copyTemp(xlBaseFile, "xl-stream-from-xl-")
    poiStreamFromXlFile = copyTemp(xlBaseFile, "poi-stream-from-xl-")
    xlStreamFromPoiFile = copyTemp(poiBaseFile, "xl-stream-from-poi-")
    poiStreamFromPoiFile = copyTemp(poiBaseFile, "poi-stream-from-poi-")

    // Dedicated targets to prevent cross-contamination between write benchmarks
    xlWriteFile = Files.createTempFile("xl-benchmark-write-", ".xlsx")
    poiWriteFile = Files.createTempFile("poi-benchmark-write-", ".xlsx")
  }

  @TearDown(Level.Trial)
  def teardown(): Unit = {
    Seq(
      xlBaseFile,
      poiBaseFile,
      xlReadFromXl,
      poiReadFromXlFile,
      xlReadFromPoiFile,
      poiReadFromPoiFile,
      xlStreamFromXl,
      poiStreamFromXlFile,
      xlStreamFromPoiFile,
      poiStreamFromPoiFile,
      xlWriteFile,
      poiWriteFile
    ).foreach { path =>
      if (path != null && Files.exists(path)) Files.delete(path)
    }
    if (poiWorkbook != null) poiWorkbook.close()
  }

  @Benchmark
  def xlWrite(): Unit = {
    // XL write performance (ScalaXml backend - default)
    ExcelIO.instance[IO].write(xlWorkbook, xlWriteFile).unsafeRunSync()
  }

  @Benchmark
  def xlWriteFast(): Unit = {
    // XL write performance (SaxStax backend - optimized)
    ExcelIO.instance[IO].writeFast(xlWorkbook, xlWriteFile).unsafeRunSync()
  }

  @Benchmark
  def poiWrite(): Unit = {
    // POI write performance
    val fos = new java.io.FileOutputStream(poiWriteFile.toFile)
    try {
      poiWorkbook.write(fos)
    } finally {
      fos.close()
    }
  }

  @Benchmark
  def xlRead(): Double = {
    // XL in-memory read performance (file pre-created in setup)
    val wb = ExcelIO.instance[IO].read(xlReadFromXl).unsafeRunSync()
    val sum = sumFirstColumn(wb)
    requireCloseEnough(sum, expectedSum, "xlRead sum validation failed")
    sum
  }

  @Benchmark
  def xlReadFromPoi(): Double = {
    val wb = ExcelIO.instance[IO].read(xlReadFromPoiFile).unsafeRunSync()
    val sum = sumFirstColumn(wb)
    requireCloseEnough(sum, expectedSum, "xlReadFromPoi sum validation failed")
    sum
  }

  @Benchmark
  def poiRead(): Double = {
    // POI in-memory read performance (file pre-created in setup)
    val fis = new java.io.FileInputStream(poiReadFromPoiFile.toFile)
    val wb =
      try WorkbookFactory.create(fis)
      finally fis.close()

    try {
      val sheet = wb.getSheetAt(0)
      val sum = sheet.iterator().asScala.foldLeft(0.0) { (acc, row) =>
        val cell = row.getCell(0)
        if cell != null then acc + cell.getNumericCellValue else acc
      }
      requireCloseEnough(sum, expectedSum, "poiRead sum validation failed")
      sum
    } finally wb.close()
  }

  @Benchmark
  def poiReadFromXl(): Double = {
    val fis = new java.io.FileInputStream(poiReadFromXlFile.toFile)
    val wb =
      try WorkbookFactory.create(fis)
      finally fis.close()

    try {
      val sheet = wb.getSheetAt(0)
      val sum = sheet.iterator().asScala.foldLeft(0.0) { (acc, row) =>
        val cell = row.getCell(0)
        if cell != null then acc + cell.getNumericCellValue else acc
      }
      requireCloseEnough(sum, expectedSum, "poiReadFromXl sum validation failed")
      sum
    } finally wb.close()
  }

  @Benchmark
  def xlReadStream(): Double = {
    // XL streaming read performance - constant memory with verifiable sum computation
    val sum = ExcelIO
      .instance[IO]
      .readStream(xlStreamFromXl)
      .evalMap { rowData =>
        IO.pure {
          rowData.cells.get(0) match {
            case Some(CellValue.Number(n)) => n.toDouble
            case _ => 0.0
          }
        }
      }
      .compile
      .fold(0.0)(_ + _)
      .unsafeRunSync()
    requireCloseEnough(sum, expectedSum, "xlReadStream sum validation failed")
    sum
  }

  @Benchmark
  def xlReadStreamFromPoi(): Double = {
    val sum = ExcelIO
      .instance[IO]
      .readStream(xlStreamFromPoiFile)
      .evalMap { rowData =>
        IO.pure {
          rowData.cells.get(0) match {
            case Some(CellValue.Number(n)) => n.toDouble
            case _ => 0.0
          }
        }
      }
      .compile
      .fold(0.0)(_ + _)
      .unsafeRunSync()
    requireCloseEnough(sum, expectedSum, "xlReadStreamFromPoi sum validation failed")
    sum
  }

  @Benchmark
  def poiReadStream(): Double = {
    val sum = poiStreamingSum(poiStreamFromXlFile)
    requireCloseEnough(sum, expectedSum, "poiReadStream sum validation failed")
    sum
  }

  @Benchmark
  def poiReadStreamFromXlMaterialized(): Double = {
    val sum = poiMaterializedStreamingSum(poiStreamFromXlFile)
    requireCloseEnough(sum, expectedSum, "poiReadStreamFromXlMaterialized sum validation failed")
    sum
  }

  @Benchmark
  def poiReadStreamFromPoiMaterialized(): Double = {
    val sum = poiMaterializedStreamingSum(poiStreamFromPoiFile)
    requireCloseEnough(sum, expectedSum, "poiReadStreamFromPoiMaterialized sum validation failed")
    sum
  }

  @Benchmark
  def poiReadStreamFromPoi(): Double = {
    val sum = poiStreamingSum(poiStreamFromPoiFile)
    requireCloseEnough(sum, expectedSum, "poiReadStreamFromPoi sum validation failed")
    sum
  }

  // ----- Helpers -----

  private def copyTemp(src: Path, prefix: String): Path = {
    val target = Files.createTempFile(prefix, ".xlsx")
    Files.copy(src, target, java.nio.file.StandardCopyOption.REPLACE_EXISTING)
    target
  }

  private def sumFirstColumn(wb: Workbook): Double = {
    val colA = Column.from0(0)
    // Find "Data" sheet specifically (Workbook.empty creates "Sheet1" first)
    wb.sheets
      .find(_.name.value == "Data")
      .map { dataSheet =>
        dataSheet.cells.valuesIterator.collect {
          case cell if cell.ref.col == colA =>
            cell.value match
              case CellValue.Number(n) => n.toDouble
              case _ => 0.0
        }.sum
      }
      .getOrElse(0.0)
  }

  private def requireCloseEnough(actual: Double, expected: Double, msg: String): Unit = {
    val tolerance = math.max(1e-6 * expected, 1e-3) // tolerate tiny floating error
    require(math.abs(actual - expected) <= tolerance, s"$msg: expected=$expected actual=$actual")
  }

  private def poiStreamingSum(path: Path): Double = {
    val pkg = org.apache.poi.openxml4j.opc.OPCPackage.open(path.toFile)
    try {
      val reader = new org.apache.poi.xssf.eventusermodel.XSSFReader(pkg)
      val parser = SAXParserFactory.newInstance().newSAXParser()
      val sheetsIter = reader.getSheetsData()
      sheetsIter.asScala.map { sheetStream =>
        try saxSumColumnA(parser, sheetStream)
        finally sheetStream.close()
      }.sum
    } finally pkg.close()
  }

  private def poiMaterializedStreamingSum(path: Path): Double = {
    val pkg = org.apache.poi.openxml4j.opc.OPCPackage.open(path.toFile)
    try {
      val reader = new org.apache.poi.xssf.eventusermodel.XSSFReader(pkg)
      val parser = SAXParserFactory.newInstance().newSAXParser()
      val sheetsIter = reader.getSheetsData()
      sheetsIter.asScala.map { sheetStream =>
        try saxSumColumnAMaterialized(parser, sheetStream)
        finally sheetStream.close()
      }.sum
    } finally pkg.close()
  }

  private def saxSumColumnA(
    parser: javax.xml.parsers.SAXParser,
    stream: java.io.InputStream
  ): Double = {
    var sum = 0.0
    var inValue = false
    var cellIsA = false
    val text = new StringBuilder

    val handler = new DefaultHandler {
      override def startElement(
        uri: String,
        localName: String,
        qName: String,
        attributes: Attributes
      ): Unit =
        qName match
          case "c" =>
            val ref = Option(attributes.getValue("r")).getOrElse("")
            cellIsA = ref.startsWith("A")
          case "v" =>
            inValue = cellIsA
            if inValue then text.setLength(0)
          case _ =>

      override def characters(ch: Array[Char], start: Int, length: Int): Unit =
        if inValue then text.appendAll(ch, start, length)

      override def endElement(uri: String, localName: String, qName: String): Unit =
        qName match
          case "v" if inValue =>
            val s = text.result().trim
            if s.nonEmpty then
              try sum += s.toDouble
              catch case _: NumberFormatException => ()
            inValue = false
          case "c" =>
            cellIsA = false
          case _ =>
    }

    parser.parse(InputSource(stream), handler)
    sum
  }

  private def saxSumColumnAMaterialized(
    parser: javax.xml.parsers.SAXParser,
    stream: java.io.InputStream
  ): Double = {
    var runningSum = 0.0
    var inValue = false
    var cellIsA = false
    val text = new StringBuilder
    var currentRow: mutable.Map[Int, Double] = mutable.Map.empty

    val handler = new DefaultHandler {
      override def startElement(
        uri: String,
        localName: String,
        qName: String,
        attributes: Attributes
      ): Unit =
        qName match
          case "row" =>
            currentRow = mutable.Map.empty
          case "c" =>
            val ref = Option(attributes.getValue("r")).getOrElse("")
            cellIsA = ref.startsWith("A")
          case "v" =>
            inValue = cellIsA
            if inValue then text.setLength(0)
          case _ =>

      override def characters(ch: Array[Char], start: Int, length: Int): Unit =
        if inValue then text.appendAll(ch, start, length)

      override def endElement(uri: String, localName: String, qName: String): Unit =
        qName match
          case "v" if inValue =>
            val s = text.result().trim
            if s.nonEmpty then
              try currentRow.update(0, s.toDouble)
              catch case _: NumberFormatException => ()
            inValue = false
          case "c" =>
            cellIsA = false
          case "row" =>
            runningSum += currentRow.values.sum
          case _ =>
    }

    parser.parse(InputSource(stream), handler)
    runningSum
  }
}
