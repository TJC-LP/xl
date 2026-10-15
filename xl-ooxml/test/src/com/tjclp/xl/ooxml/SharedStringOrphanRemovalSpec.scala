package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.util.zip.{ZipEntry, ZipInputStream, ZipOutputStream}

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.macros.ref
import com.tjclp.xl.ooxml.lint.{LintCategory, WorkbookLint}
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * GH-768: text an edit removes from every cell leaves xl/sharedStrings.xml on the next write.
 *
 *   - An entry above every position an untouched (verbatim-copied) sheet references is dropped; the
 *     regenerated sheets look their strings up again.
 *   - An entry at or below such a position becomes an empty `<si>`, so the untouched sheet keeps
 *     its bytes and its positions.
 *   - `WriterConfig.compactSharedStrings` drops every unreferenced entry, regenerating the
 *     untouched sheets that reference a later position.
 *
 * Each case checks the removed text is in no part of the package, the `shared-string-orphan` lint
 * is clean, and the cell values read back unchanged.
 */
class SharedStringOrphanRemovalSpec extends ScalaCheckSuite:

  // ========== fixtures ==========

  /** One sheet: A1 → OldLabelAlpha (0), A2 → KeepMe (1), A3 → Row3Text (2). count=3. */
  private val singleSheet: Array[Byte] = book(
    Vector("OldLabelAlpha", "KeepMe", "Row3Text"),
    Vector(Vector("A1" -> 0, "A2" -> 1, "A3" -> 2))
  )

  /**
   * Two sheets. Sheet1: A1 → OldLabelAlpha (0), A2 → KeepMe (1). Sheet2: A1 → KeepMe (1), A2 →
   * Sheet2Only (2), A3 → Tail (3). count=5, uniqueCount=4 (+ `extra` unreferenced entries).
   */
  private def twoSheets(extra: String*): Array[Byte] = book(
    Vector("OldLabelAlpha", "KeepMe", "Sheet2Only", "Tail") ++ extra,
    Vector(Vector("A1" -> 0, "A2" -> 1), Vector("A1" -> 1, "A2" -> 2, "A3" -> 3))
  )

  // ========== a sheet with no untouched neighbour: the table is rebuilt ==========

  test("GH-768: replacing a cell's text drops the old string (single sheet)") {
    val out = editSheet(singleSheet, "Sheet1")(_.put(ref"A1" -> "NewLabelBeta"))
    assertNoText(out, "OldLabelAlpha")
    assertEquals(sstTexts(out), Vector("KeepMe", "Row3Text", "NewLabelBeta"))
    assertSstCounts(out, count = 3, unique = 3)
    assertOrphanLintClean(out)
    assertEquals(
      texts(out, "Sheet1"),
      Map("A1" -> "NewLabelBeta", "A2" -> "KeepMe", "A3" -> "Row3Text")
    )
  }

  test("GH-768: clearing a cell drops its string (single sheet)") {
    val out = editSheet(singleSheet, "Sheet1")(_.remove(ref"A1"))
    assertNoText(out, "OldLabelAlpha")
    assertEquals(sstTexts(out), Vector("KeepMe", "Row3Text"))
    assertSstCounts(out, count = 2, unique = 2)
    assertOrphanLintClean(out)
    assertEquals(texts(out, "Sheet1"), Map("A2" -> "KeepMe", "A3" -> "Row3Text"))
  }

  test("GH-768: deleting a row drops the strings only it held (single sheet)") {
    val out = editSheet(singleSheet, "Sheet1")(_.deleteRows(0, 1))
    assertNoText(out, "OldLabelAlpha")
    assertOrphanLintClean(out)
    assertEquals(texts(out, "Sheet1"), Map("A1" -> "KeepMe", "A2" -> "Row3Text"))
  }

  test("GH-768: removing a sheet drops the strings only it referenced") {
    val out = write(
      read(twoSheets()).remove(SheetName.unsafe("Sheet2")).fold(e => fail(e.message), identity)
    )
    assertNoText(out, "Sheet2Only")
    assertNoText(out, "Tail")
    assertOrphanLintClean(out)
    assertEquals(texts(out, "Sheet1"), Map("A1" -> "OldLabelAlpha", "A2" -> "KeepMe"))
  }

  // ========== an untouched sheet keeps its bytes ==========

  test("GH-768: below an untouched sheet's positions the removed string becomes an empty entry") {
    val source = twoSheets()
    val out = editSheet(source, "Sheet1")(_.put(ref"A1" -> "NewLabelBeta"))
    assertNoText(out, "OldLabelAlpha")
    assertEquals(sstTexts(out), Vector("", "KeepMe", "Sheet2Only", "Tail", "NewLabelBeta"))
    assertSstCounts(out, count = 5, unique = 5)
    assertEquals(
      entry(out, "xl/worksheets/sheet2.xml").toSeq,
      entry(source, "xl/worksheets/sheet2.xml").toSeq,
      "the untouched sheet must ride verbatim"
    )
    assertOrphanLintClean(out)
    assertEquals(texts(out, "Sheet1"), Map("A1" -> "NewLabelBeta", "A2" -> "KeepMe"))
    assertEquals(texts(out, "Sheet2"), Map("A1" -> "KeepMe", "A2" -> "Sheet2Only", "A3" -> "Tail"))
  }

  test("GH-768: above every untouched position the removed string is dropped outright") {
    val source = twoSheets()
    val out = editSheet(source, "Sheet2")(_.put(ref"A3" -> "NewTail"))
    assertNoText(out, ">Tail<")
    assertEquals(sstTexts(out), Vector("OldLabelAlpha", "KeepMe", "Sheet2Only", "NewTail"))
    assertSstCounts(out, count = 5, unique = 4)
    assertEquals(
      entry(out, "xl/worksheets/sheet1.xml").toSeq,
      entry(source, "xl/worksheets/sheet1.xml").toSeq
    )
    assertOrphanLintClean(out)
    assertEquals(
      texts(out, "Sheet2"),
      Map("A1" -> "KeepMe", "A2" -> "Sheet2Only", "A3" -> "NewTail")
    )
  }

  test("GH-768: an edit that releases no string copies the table verbatim, old orphans included") {
    val source = twoSheets("Stale")
    val out = editSheet(source, "Sheet1")(_.put(ref"B1" -> 42))
    assertEquals(
      entry(out, "xl/sharedStrings.xml").toSeq,
      entry(source, "xl/sharedStrings.xml").toSeq,
      "a write must not touch strings its edit did not release"
    )
  }

  test("GH-768: a string still referenced by an untouched sheet is kept") {
    val out = editSheet(twoSheets(), "Sheet1")(_.remove(ref"A2"))
    assertEquals(sstTexts(out), Vector("OldLabelAlpha", "KeepMe", "Sheet2Only", "Tail"))
    assertSstCounts(out, count = 4, unique = 4)
    assertOrphanLintClean(out)
  }

  test("GH-768: several removed strings below an untouched sheet's positions keep every position") {
    // Sheet1 alone held 0, 2 and 4 (4 a rich-text entry); Sheet2 references 1, 3 and 5
    val source = book(
      Vector("Gone0", "Kept1", "Gone2", "Kept3", "RICHGONE", "Kept5"),
      Vector(Vector("A1" -> 0, "A2" -> 2, "A3" -> 4), Vector("A1" -> 1, "A2" -> 3, "A3" -> 5))
    )
    val withRich = rewriteEntry(
      source,
      "xl/sharedStrings.xml",
      _.replace(
        "<si><t>RICHGONE</t></si>",
        "<si><r><rPr><b/></rPr><t>RICH</t></r><r><t>GONE</t></r></si>"
      )
    )
    val out = editSheet(withRich, "Sheet1")(sheet =>
      sheet.remove(ref"A1").remove(ref"A2").remove(ref"A3").put(ref"B1" -> "Kept3")
    )
    Vector("Gone0", "Gone2", ">RICH<", ">GONE<").foreach(assertNoText(out, _))
    assertEquals(sstTexts(out), Vector("", "Kept1", "", "Kept3", "", "Kept5"))
    assertEquals(
      entry(out, "xl/worksheets/sheet2.xml").toSeq,
      entry(withRich, "xl/worksheets/sheet2.xml").toSeq
    )
    assertSstCounts(out, count = 4, unique = 6)
    assertOrphanLintClean(out)
    assertEquals(texts(out, "Sheet1"), Map("B1" -> "Kept3"))
    assertEquals(texts(out, "Sheet2"), Map("A1" -> "Kept1", "A2" -> "Kept3", "A3" -> "Kept5"))
  }

  test("GH-768: a duplicate entry the edit released becomes empty; its twin keeps the text") {
    // positions 0 and 2 both spell "Dup"; Sheet1 references 0, Sheet2 references 1 and 2
    val source = book(
      Vector("Dup", "Other", "Dup"),
      Vector(Vector("A1" -> 0), Vector("A1" -> 1, "A2" -> 2))
    )
    val out = editSheet(source, "Sheet1")(_.put(ref"B1" -> 1))
    assertEquals(
      entry(out, "xl/worksheets/sheet2.xml").toSeq,
      entry(source, "xl/worksheets/sheet2.xml").toSeq
    )
    assertEquals(texts(out, "Sheet1"), Map("A1" -> "Dup"))
    assertEquals(texts(out, "Sheet2"), Map("A1" -> "Other", "A2" -> "Dup"))
    assertOrphanLintClean(out)
  }

  test("GH-768: a revision log xl does not model pins the positions it references") {
    // Sheet1 alone held 0 and 2; the revision log references 3, Sheet2 references 1
    val source = book(
      Vector("Gone0", "Kept1", "Gone2", "RevisionText"),
      Vector(Vector("A1" -> 0, "A2" -> 2), Vector("A1" -> 1)),
      revisionLog = Some(Vector("A1" -> 3))
    )
    for config <- Vector(WriterConfig(), compact) do
      val out = editSheet(source, "Sheet1", config)(_.remove(ref"A1").remove(ref"A2"))
      assertNoText(out, "Gone0")
      assertNoText(out, "Gone2")
      assertEquals(
        entry(out, "xl/revisions/revisionLog1.xml").toSeq,
        entry(source, "xl/revisions/revisionLog1.xml").toSeq
      )
      // the log pins position 3, so even compacting leaves empty entries below it
      assertEquals(sstTexts(out), Vector("", "Kept1", "", "RevisionText"), config.toString)
      assertEquals(texts(out, "Sheet2"), Map("A1" -> "Kept1"))
  }

  test("GH-768: an untouched part that cannot be read keeps every entry (the old append plan)") {
    // the revision log is not well-formed, so it may reference any position
    val source = rewriteEntry(
      book(
        Vector("OldLabelAlpha", "KeepMe"),
        Vector(Vector("A1" -> 0, "A2" -> 1)),
        revisionLog = Some(Vector("A1" -> 1))
      ),
      "xl/revisions/revisionLog1.xml",
      _ => "<revisions><rcc><nc"
    )
    val out = editSheet(source, "Sheet1")(_.put(ref"A1" -> "NewLabelBeta"))
    assertEquals(sstTexts(out), Vector("OldLabelAlpha", "KeepMe", "NewLabelBeta"))
    assertEquals(
      entry(out, "xl/revisions/revisionLog1.xml").toSeq,
      entry(source, "xl/revisions/revisionLog1.xml").toSeq
    )
    assertEquals(texts(out, "Sheet1"), Map("A1" -> "NewLabelBeta", "A2" -> "KeepMe"))
  }

  // ========== compactSharedStrings ==========

  private val compact = WriterConfig(compactSharedStrings = true)

  test("GH-768: compactSharedStrings renumbers instead of leaving an empty entry") {
    val out = editSheet(twoSheets(), "Sheet1", compact)(_.put(ref"A1" -> "NewLabelBeta"))
    assertNoText(out, "OldLabelAlpha")
    assertEquals(sstTexts(out), Vector("KeepMe", "Sheet2Only", "Tail", "NewLabelBeta"))
    assertSstCounts(out, count = 5, unique = 4)
    assertOrphanLintClean(out)
    assertEquals(texts(out, "Sheet1"), Map("A1" -> "NewLabelBeta", "A2" -> "KeepMe"))
    assertEquals(texts(out, "Sheet2"), Map("A1" -> "KeepMe", "A2" -> "Sheet2Only", "A3" -> "Tail"))
  }

  test("GH-768: compactSharedStrings also drops orphans the source already carried") {
    val out = editSheet(twoSheets("Stale"), "Sheet1", compact)(_.put(ref"B1" -> 42))
    assertNoText(out, "Stale")
    assertEquals(sstTexts(out), Vector("OldLabelAlpha", "KeepMe", "Sheet2Only", "Tail"))
    assertOrphanLintClean(out)
  }

  test("GH-768: compactSharedStrings leaves a sheet referencing only earlier positions untouched") {
    // Sheet2 references positions up to 3; the orphan "Tail2" sits at 4, above every position
    val source = book(
      Vector("A", "B", "C", "D", "Tail2"),
      Vector(Vector("A1" -> 0, "A2" -> 4), Vector("A1" -> 1, "A2" -> 2, "A3" -> 3))
    )
    val out = editSheet(source, "Sheet1", compact)(_.remove(ref"A2"))
    assertEquals(sstTexts(out), Vector("A", "B", "C", "D"))
    assertEquals(
      entry(out, "xl/worksheets/sheet2.xml").toSeq,
      entry(source, "xl/worksheets/sheet2.xml").toSeq
    )
    assertOrphanLintClean(out)
  }

  test("GH-768: compactSharedStrings regenerates every untouched sheet a renumbering reaches") {
    // Sheet1 alone held 0; Sheet2 references 1..2, Sheet3 references 3: both must be renumbered
    val source = book(
      Vector("Gone0", "S2a", "S2b", "S3a"),
      Vector(Vector("A1" -> 0), Vector("A1" -> 1, "A2" -> 2), Vector("A1" -> 3))
    )
    val out = editSheet(source, "Sheet1", compact)(_.remove(ref"A1"))
    assertNoText(out, "Gone0")
    assertEquals(sstTexts(out), Vector("S2a", "S2b", "S3a"))
    assertSstCounts(out, count = 3, unique = 3)
    assertEquals(texts(out, "Sheet2"), Map("A1" -> "S2a", "A2" -> "S2b"))
    assertEquals(texts(out, "Sheet3"), Map("A1" -> "S3a"))
    assertOrphanLintClean(out)
  }

  test("GH-768: compactSharedStrings on a clean book with an orphan rewrites the table") {
    val wb = read(twoSheets("Stale"))
    val out = XlsxWriter.writeToBytes(wb, compact).fold(e => fail(e.message), identity)
    assertNoText(out, "Stale")
    assertOrphanLintClean(out)
    assertEquals(texts(out, "Sheet2"), Map("A1" -> "KeepMe", "A2" -> "Sheet2Only", "A3" -> "Tail"))
  }

  test("GH-768: compactSharedStrings on a clean book without orphans copies it verbatim") {
    val source = twoSheets()
    val out = XlsxWriter.writeToBytes(read(source), compact).fold(e => fail(e.message), identity)
    assertEquals(out.toSeq, source.toSeq)
  }

  // ========== no source table ==========

  test("GH-768: a fresh table on a surgical write holds only the regenerated sheets' strings") {
    val inline = WriterConfig(sstPolicy = SstPolicy.Never)
    val sheets = Vector(
      Sheet("Sheet1").put(ref"A1" -> "one"),
      Sheet("Sheet2").put(ref"A1" -> "untouched text")
    )
    val source =
      XlsxWriter.writeToBytes(Workbook(sheets), inline).fold(e => fail(e.message), identity)
    val wb = read(source)
    val edited = wb
      .update(SheetName.unsafe("Sheet1"), _.put(ref"A2" -> "two"))
      .fold(e => fail(e.message), identity)
    val out = XlsxWriter
      .writeToBytes(edited, WriterConfig(sstPolicy = SstPolicy.Always))
      .fold(e => fail(e.message), identity)
    assertEquals(sstTexts(out), Vector("one", "two"))
    assertOrphanLintClean(out)
    assertEquals(texts(out, "Sheet2"), Map("A1" -> "untouched text"))
  }

  // ========== a t="s" index that names no entry (ECMA-376 §18.4.9) ==========

  /** The indices past the end of a 3-entry table: one just past it, and Int.MaxValue. */
  private val pastEnd = Vector(3, Int.MaxValue)

  test("GH-768: an edited sheet's index past the table neither pins an entry nor fails the write") {
    for bad <- pastEnd; config <- Vector(WriterConfig(), compact) do
      val source = book(
        Vector("Gone0", "Kept1", "Kept2"),
        Vector(Vector("A1" -> 0, "A2" -> bad), Vector("A1" -> 1, "A2" -> 2))
      )
      val out = editSheet(source, "Sheet1", config)(_.remove(ref"A1").remove(ref"A2"))
      val clue = s"bad=$bad $config"
      assertNoText(out, "Gone0")
      assertEquals(texts(out, "Sheet2"), Map("A1" -> "Kept1", "A2" -> "Kept2"), clue)
      assertEquals(texts(out, "Sheet1"), Map.empty[String, String], clue)
      assertOrphanLintClean(out)
  }

  test("GH-768: an untouched part with an index past the table keeps the table unchanged") {
    for bad <- pastEnd; config <- Vector(WriterConfig(), compact) do
      val source = book(
        Vector("OldLabelAlpha", "KeepMe", "Sheet2Only"),
        Vector(Vector("A1" -> 0, "A2" -> 1), Vector("A1" -> 2, "A2" -> bad))
      )
      // the edit releases OldLabelAlpha and adds no string
      val out = editSheet(source, "Sheet1", config)(_.remove(ref"A1"))
      val clue = s"bad=$bad $config"
      // the untouched sheet is malformed, so the write keeps every entry, as for an unread part
      assertEquals(sstTexts(out), Vector("OldLabelAlpha", "KeepMe", "Sheet2Only"), clue)
      assertEquals(
        entry(out, "xl/worksheets/sheet2.xml").toSeq,
        entry(source, "xl/worksheets/sheet2.xml").toSeq,
        clue
      )
      assertEquals(texts(out, "Sheet1"), Map("A2" -> "KeepMe"), clue)
      assertEquals(texts(out, "Sheet2"), Map("A1" -> "Sheet2Only"), clue)
  }

  test("GH-768: a scan holds one bit per table entry at most, whatever index a cell names") {
    val scan = scanCells(Vector("0", "2", "3", Int.MaxValue.toString, "-1", "x"), tableSize = 3)
    assertEquals(scan.indices, scala.collection.immutable.BitSet(0, 2))
    assertEquals(scan.outOfRange, 4)
    assertEquals(scan.count, 6)
    assertEquals(scan.maxIndex, 2)
    assert(scan.indices.toBitMask.length <= 1, scan.indices.toBitMask.length.toString)
  }

  property("GH-768: a scan keeps exactly the indices below the table's size") {
    val genIndex = Gen.frequency(4 -> Gen.choose(0, 40), 1 -> Gen.choose(0, Int.MaxValue))
    Prop.forAll(Gen.listOf(genIndex), Gen.choose(0, 40)) { (indices, tableSize) =>
      val scan = scanCells(indices.toVector.map(_.toString), tableSize)
      val (in, out) = indices.partition(_ < tableSize)
      assertEquals(scan.indices, scala.collection.immutable.BitSet.fromSpecific(in))
      assertEquals(scan.outOfRange, out.size)
      assert(scan.indices.toBitMask.length <= tableSize / 64 + 1)
    }
  }

  private def scanCells(values: Vector[String], tableSize: Int): SharedStringCompaction.RefScan =
    val cells = values.zipWithIndex
      .map((v, i) => s"""<c r="A${i + 1}" t="s"><v>$v</v></c>""")
      .mkString
    val xml = s"<worksheet><sheetData><row r=\"1\">$cells</row></sheetData></worksheet>"
    SharedStringCompaction
      .scan(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)), tableSize)
      .fold(e => fail(e), identity)

  // ========== round-trip property ==========

  private val pool = Vector("alpha", "beta", "gamma", "delta", "epsilon", "zeta")
  private val refs = Vector("A1", "A2", "A3", "B1", "B2", "B3")

  private val genSheet: Gen[Map[String, String]] =
    Gen.listOf(Gen.zip(Gen.oneOf(refs), Gen.oneOf(pool))).map(_.toMap)

  /** An edit of one cell: Some(text) puts, None clears. */
  private val genEdits: Gen[List[(String, Option[String])]] =
    Gen.listOf(Gen.zip(Gen.oneOf(refs), Gen.option(Gen.oneOf(pool :+ "fresh")))).map(_.take(6))

  property("GH-768: after any edit, values read back and no entry with text is unreferenced") {
    Prop.forAll(genSheet, genSheet, genEdits, Gen.oneOf(false, true)) {
      (s1, s2, edits, compacting) =>
        val fresh = Workbook(Vector(sheetOf("Sheet1", s1), sheetOf("Sheet2", s2)))
        val source = XlsxWriter
          .writeToBytes(fresh, WriterConfig(sstPolicy = SstPolicy.Always))
          .fold(e => fail(e.message), identity)
        val expected1 = edits.foldLeft(s1) {
          case (m, (r, Some(v))) => m.updated(r, v)
          case (m, (r, None)) => m - r
        }
        val edited = read(source)
          .update(SheetName.unsafe("Sheet1"), _ => sheetOf("Sheet1", expected1))
          .fold(e => fail(e.message), identity)
        val config = WriterConfig(compactSharedStrings = compacting)
        val out = XlsxWriter.writeToBytes(edited, config).fold(e => fail(e.message), identity)
        assertEquals(texts(out, "Sheet1"), expected1)
        assertEquals(texts(out, "Sheet2"), s2)
        assertOrphanLintClean(out)
        if compacting then assert(!sstTexts(out).contains(""), sstTexts(out).toString)
        else
          assertEquals(
            entry(out, "xl/worksheets/sheet2.xml").toSeq,
            entry(source, "xl/worksheets/sheet2.xml").toSeq
          )
    }
  }

  // ========== helpers ==========

  private def sheetOf(name: String, cells: Map[String, String]): Sheet =
    cells.toVector.sorted.foldLeft(Sheet(SheetName.unsafe(name))) { case (s, (r, v)) =>
      s.put(ARef.parse(r).fold(e => fail(e), identity) -> v)
    }

  private def read(bytes: Array[Byte]): Workbook =
    XlsxReader.readFromBytes(bytes).fold(e => fail(s"read failed: $e"), identity)

  private def write(wb: Workbook, config: WriterConfig = WriterConfig()): Array[Byte] =
    XlsxWriter.writeToBytes(wb, config).fold(e => fail(s"write failed: $e"), identity)

  private def editSheet(source: Array[Byte], name: String, config: WriterConfig = WriterConfig())(
    f: Sheet => Sheet
  ): Array[Byte] =
    val wb = read(source)
    write(wb.update(SheetName.unsafe(name), f).fold(e => fail(e.message), identity), config)

  private def texts(bytes: Array[Byte], sheet: String): Map[String, String] =
    read(bytes)(SheetName.unsafe(sheet))
      .fold(e => fail(e.message), identity)
      .cells
      .toVector
      .flatMap { (r, cell) =>
        cell.value match
          case CellValue.Text(s) => Some(r.toA1 -> s)
          case _ => None
      }
      .toMap

  private def entries(bytes: Array[Byte]): Vector[(String, Array[Byte])] =
    val zip = new ZipInputStream(new ByteArrayInputStream(bytes))
    try
      Iterator
        .continually(zip.getNextEntry)
        .takeWhile(_ != null)
        .map(e => e.getName -> zip.readAllBytes())
        .toVector
    finally zip.close()

  /** The package with one part's text replaced, every other part as it was. */
  private def rewriteEntry(bytes: Array[Byte], name: String, f: String => String): Array[Byte] =
    val buffer = new ByteArrayOutputStream()
    val out = new ZipOutputStream(buffer)
    entries(bytes).foreach { (n, content) =>
      out.putNextEntry(new ZipEntry(n))
      out.write(
        if n == name then
          f(new String(content, StandardCharsets.UTF_8)).getBytes(StandardCharsets.UTF_8)
        else content
      )
      out.closeEntry()
    }
    out.close()
    buffer.toByteArray

  private def entry(bytes: Array[Byte], name: String): Array[Byte] =
    entries(bytes).find(_._1 == name).map(_._2).getOrElse(fail(s"missing $name"))

  private def sstXml(bytes: Array[Byte]): String =
    new String(entry(bytes, "xl/sharedStrings.xml"), StandardCharsets.UTF_8)

  private def sstTexts(bytes: Array[Byte]): Vector[String] =
    SharedStrings
      .fromXml(scala.xml.XML.loadString(sstXml(bytes)))
      .fold(e => fail(e), identity)
      .strings
      .map(_.fold(identity, _.toPlainText))

  private def assertSstCounts(bytes: Array[Byte], count: Int, unique: Int): Unit =
    val xml = sstXml(bytes)
    assert(xml.contains(s"""count="$count""""), xml)
    assert(xml.contains(s"""uniqueCount="$unique""""), xml)

  /** The removed text appears in no part of the package. */
  private def assertNoText(bytes: Array[Byte], text: String): Unit =
    entries(bytes).foreach { (name, content) =>
      assert(
        !new String(content, StandardCharsets.UTF_8).contains(text),
        s"'$text' still rides in $name"
      )
    }

  private def assertOrphanLintClean(bytes: Array[Byte]): Unit =
    val findings = WorkbookLint.lintBytes(bytes).fold(e => fail(e.message), identity)
    assertEquals(findings.filter(_.category == LintCategory.SharedStringOrphan), Vector.empty)

  /**
   * A raw package: the table holds `strings`; each sheet lists (ref, t="s" index) cells. A
   * `revisionLog`, when given, is a part xl does not model whose cell values (`<nc t="s">`)
   * reference the table too.
   */
  private def book(
    strings: Vector[String],
    sheets: Vector[Vector[(String, Int)]],
    revisionLog: Option[Vector[(String, Int)]] = None
  ): Array[Byte] =
    val n = sheets.size
    val bytes = new ByteArrayOutputStream()
    val out = new ZipOutputStream(bytes)
    def put(name: String, content: String): Unit =
      out.putNextEntry(new ZipEntry(name))
      out.write(content.getBytes(StandardCharsets.UTF_8))
      out.closeEntry()
    val ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    val rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    put(
      "[Content_Types].xml",
      """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">""" +
        """<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>""" +
        """<Default Extension="xml" ContentType="application/xml"/>""" +
        """<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>""" +
        (1 to n)
          .map(i =>
            s"""<Override PartName="/xl/worksheets/sheet$i.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"""
          )
          .mkString +
        revisionLog.fold("")(_ =>
          """<Override PartName="/xl/revisions/revisionLog1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml"/>"""
        ) +
        """<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>""" +
        "</Types>"
    )
    put(
      "_rels/.rels",
      """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""" +
        s"""<Relationship Id="rId1" Type="$rel/officeDocument" Target="xl/workbook.xml"/></Relationships>"""
    )
    put(
      "xl/workbook.xml",
      s"""<?xml version="1.0"?><workbook xmlns="$ns" xmlns:r="$rel"><sheets>""" +
        (1 to n).map(i => s"""<sheet name="Sheet$i" sheetId="$i" r:id="rId$i"/>""").mkString +
        "</sheets></workbook>"
    )
    put(
      "xl/_rels/workbook.xml.rels",
      """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""" +
        (1 to n)
          .map(i =>
            s"""<Relationship Id="rId$i" Type="$rel/worksheet" Target="worksheets/sheet$i.xml"/>"""
          )
          .mkString +
        s"""<Relationship Id="rId${n + 1}" Type="$rel/sharedStrings" Target="sharedStrings.xml"/>""" +
        "</Relationships>"
    )
    put(
      "xl/sharedStrings.xml",
      s"""<?xml version="1.0" encoding="UTF-8"?><sst xmlns="$ns" count="${sheets
          .map(_.size)
          .sum}" uniqueCount="${strings.size}">""" +
        strings.map(s => s"<si><t>$s</t></si>").mkString + "</sst>"
    )
    def rowsXml(cells: Vector[(String, Int)]): String = cells
      .groupBy((r, _) => r.dropWhile(_.isLetter).toInt)
      .toVector
      .sortBy(_._1)
      .map { (row, cs) =>
        s"""<row r="$row">""" +
          cs.sortBy(_._1).map((r, idx) => s"""<c r="$r" t="s"><v>$idx</v></c>""").mkString +
          "</row>"
      }
      .mkString
    sheets.zipWithIndex.foreach { (cells, i) =>
      put(
        s"xl/worksheets/sheet${i + 1}.xml",
        s"""<?xml version="1.0"?><worksheet xmlns="$ns"><sheetData>${rowsXml(
            cells
          )}</sheetData></worksheet>"""
      )
    }
    revisionLog.foreach { cells =>
      val changes = cells
        .map((r, idx) => s"""<rcc rId="1" sId="1"><nc r="$r" t="s"><v>$idx</v></nc></rcc>""")
        .mkString
      put(
        "xl/revisions/revisionLog1.xml",
        s"""<?xml version="1.0"?><revisions xmlns="$ns">$changes</revisions>"""
      )
    }
    out.close()
    bytes.toByteArray
