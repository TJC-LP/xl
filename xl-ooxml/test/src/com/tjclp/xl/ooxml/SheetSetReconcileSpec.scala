package com.tjclp.xl.ooxml

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.ZipFile
import scala.jdk.CollectionConverters.*

import munit.FunSuite

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.macros.ref

/**
 * GH-470: `copy(sheets = ...)` reductions were silently IGNORED by the preservation write — a
 * clean-looking SourceContext verbatim-copied (or surgically re-emitted) the SOURCE structure, so a
 * "redacted" workbook built by filtering `sheets` shipped the sheet the caller believes was removed
 * (a confidentiality hazard, not just a correctness bug). Same story for defined names dropped via
 * `copy(metadata = ...)` and for sheets APPENDED via `copy(sheets = ...)` (silently lost).
 *
 * The fix dirty-tracks the sheet set and the defined-name set at write time: the writer reconciles
 * the context against the model before strategy dispatch, routing untracked removals through the
 * same marking `Workbook.remove` uses (tracker deletion + identity-mapping drops, so GH-417 orphan
 * pruning engages) and forcing metadata regeneration for untracked additions / defined-name drift.
 */
class SheetSetReconcileSpec extends FunSuite:

  private def name(s: String): SheetName = SheetName.unsafe(s)

  private def write(wb: Workbook, label: String): Path =
    val out = Files.createTempFile(s"xl-470-$label-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter.write(wb, out).fold(err => fail(s"$label write failed: ${err.message}"), identity)
    out

  private def read(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"read failed: ${err.message}"), identity)

  private def entryNames(path: Path): Set[String] =
    val zip = new ZipFile(path.toFile)
    try zip.entries().asScala.map(_.getName).toSet
    finally zip.close()

  private def entryText(path: Path, entry: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(entry)) match
        case Some(e) =>
          new String(zip.getInputStream(e).readAllBytes(), StandardCharsets.UTF_8)
        case None => fail(s"zip entry $entry not found in $path (have ${entryNames(path)})")
    finally zip.close()

  /** Three-sheet fixture with distinct text markers, written then re-read (context present). */
  private def threeSheetSource(label: String): Path =
    val wb = Workbook(
      Vector(
        Sheet(name("One")).put(Cell(ref"A1", CellValue.Text("one-marker"))),
        Sheet(name("Two")).put(Cell(ref"A1", CellValue.Text("two-marker"))),
        Sheet(name("Three")).put(Cell(ref"A1", CellValue.Text("three-marker")))
      )
    )
    write(wb, s"$label-src")

  test("GH-470: clean context + copy(sheets = take(2)) must NOT ship the third sheet (c036)") {
    val loaded = read(threeSheetSource("clean"))
    assert(loaded.sourceContext.exists(_.isClean), "fixture must start clean")

    val reduced = loaded.copy(sheets = loaded.sheets.take(2))
    val out = write(reduced, "clean-reduced")

    assertEquals(read(out).sheetNames.map(_.value), Vector("One", "Two"))
    val wbXml = entryText(out, "xl/workbook.xml")
    assert(!wbXml.contains("name=\"Three\""), s"'removed' sheet leaked into workbook.xml: $wbXml")
    val worksheetParts =
      entryNames(out).filter(n => n.startsWith("xl/worksheets/") && !n.contains("_rels"))
    assertEquals(
      worksheetParts,
      Set("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"),
      "the removed sheet's part must not ride the preservation write"
    )
  }

  test("GH-470: dirty (surgical) write of a copy-reduced workbook drops the sheet too") {
    val loaded = read(threeSheetSource("dirty"))
    val first = loaded.sheets.headOption.getOrElse(fail("no sheet"))
    val edited = loaded.put(first.put(ref"B1", CellValue.Text("edited")))
    val reduced = edited.copy(sheets = edited.sheets.take(2))

    val out = write(reduced, "dirty-reduced")

    val result = read(out)
    assertEquals(result.sheetNames.map(_.value), Vector("One", "Two"))
    assertEquals(
      result.sheets.headOption.map(_(ref"B1").value),
      Some(CellValue.Text("edited")),
      "the tracked edit must survive reconciliation of the untracked reduction"
    )
    assert(!entryText(out, "xl/workbook.xml").contains("name=\"Three\""))
  }

  test("GH-470: copy-reduction writes the same package structure as remove()") {
    val src = threeSheetSource("parity")

    val viaRemove = read(src).remove(name("Three")).fold(e => fail(e.message), identity)
    val viaCopy = read(src).copy(sheets = read(src).sheets.take(2))

    val removeOut = write(viaRemove, "parity-remove")
    val copyOut = write(viaCopy, "parity-copy")

    assertEquals(entryNames(copyOut), entryNames(removeOut))
    assertEquals(
      entryText(copyOut, "xl/workbook.xml"),
      entryText(removeOut, "xl/workbook.xml"),
      "the sanctioned remove() and the reconciled copy-reduction must agree on workbook.xml"
    )
  }

  test("GH-470: an edit to a SURVIVING sheet made after the reduction is not lost") {
    // The untracked removal must not shift live modified-sheet marks: drop the FIRST sheet, then
    // edit the (new) first sheet — its strings must land in the output despite the ghost of "One"
    // reconciling at source index 0.
    val loaded = read(threeSheetSource("shift"))
    val reduced = loaded.copy(sheets = loaded.sheets.drop(1))
    val second = reduced.sheets.headOption.getOrElse(fail("no sheet"))
    val edited = reduced.put(second.put(ref"C1", CellValue.Text("fresh-string")))

    val out = write(edited, "shift-out")

    val result = read(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Two", "Three"))
    assertEquals(
      result.sheets.headOption.map(_(ref"C1").value),
      Some(CellValue.Text("fresh-string"))
    )
  }

  test("GH-470: defined names dropped via copy(metadata = ...) must not ship") {
    val src = write(
      Workbook(Sheet(name("One")).put(Cell(ref"A1", CellValue.Number(BigDecimal(1)))))
        .withDefinedName("SECRET_RATE", "One!$A$1"),
      "dn-src"
    )
    val loaded = read(src)
    assert(
      loaded.metadata.definedNames.exists(_.name == "SECRET_RATE"),
      s"fixture must carry the name: ${loaded.metadata.definedNames}"
    )
    assert(loaded.sourceContext.exists(_.isClean), "fixture must start clean")

    val redacted = loaded.copy(metadata = loaded.metadata.copy(definedNames = Vector.empty))
    val out = write(redacted, "dn-out")

    assert(
      !entryText(out, "xl/workbook.xml").contains("SECRET_RATE"),
      "the 'removed' defined name leaked into workbook.xml"
    )
    assertEquals(read(out).metadata.definedNames, Vector.empty)
  }

  test("GH-470: a sheet APPENDED via copy(sheets = ...) ships instead of silently vanishing") {
    val loaded = read(threeSheetSource("append"))
    val appended = loaded.copy(
      sheets = loaded.sheets :+ Sheet(name("Appended"))
        .put(Cell(ref"A1", CellValue.Text("appended-marker")))
    )

    val out = write(appended, "append-out")

    val result = read(out)
    assertEquals(result.sheetNames.map(_.value), Vector("One", "Two", "Three", "Appended"))
    assertEquals(
      result.sheets.lastOption.map(_(ref"A1").value),
      Some(CellValue.Text("appended-marker"))
    )
  }

  test("GH-470: untracked reduction prunes the removed sheet's chart/drawing chain (GH-417)") {
    val charts = name("Charts")
    val chartSheet = Sheet(charts)
      .put(Cell(ref"A2", CellValue.Text("a")))
      .put(Cell(ref"A3", CellValue.Text("b")))
      .put(Cell(ref"B2", CellValue.Number(BigDecimal(3))))
      .put(Cell(ref"B3", CellValue.Number(BigDecimal(4))))
    val chart = Chart
      .bar(
        Vector(Series(DataRef(charts, ref"B2:B3"), Some(DataRef(charts, ref"A2:A3")), None)),
        title = Some("Orphan Me")
      )
      .fold(e => fail(s"chart build failed: $e"), identity)
    val wb = Workbook(
      Vector(
        Sheet(name("Data")).put(Cell(ref"A1", CellValue.Text("keep"))),
        chartSheet.addChart(chart, ref"D2:K15")
      )
    )
    val src = write(wb, "chart-src")
    assert(entryNames(src).contains("xl/charts/chart1.xml"), s"fixture: ${entryNames(src)}")

    val loaded = read(src)
    val reduced = loaded.copy(sheets = loaded.sheets.filterNot(_.name == charts))
    val out = write(reduced, "chart-out")

    val names = entryNames(out)
    assertEquals(
      names.filter(n => n.startsWith("xl/charts/") || n.startsWith("xl/drawings/")),
      Set.empty[String],
      s"the chart chain must fall with its copy-removed sheet: ${names.toVector.sorted}"
    )
    assertEquals(read(out).sheetNames.map(_.value), Vector("Data"))
  }

  test("GH-470: a genuinely untouched workbook still takes the verbatim fast path") {
    val src = threeSheetSource("verbatim")
    val loaded = read(src)
    val out = write(loaded, "verbatim-out")
    assertEquals(
      Files.readAllBytes(out).toVector,
      Files.readAllBytes(src).toVector,
      "reconciliation must return the context unchanged when nothing diverged"
    )
  }
