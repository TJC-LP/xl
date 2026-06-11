package com.tjclp.xl.ooxml

import java.nio.file.{Files, Path, Paths}
import java.util.zip.ZipFile
import scala.jdk.CollectionConverters.*

import munit.FunSuite

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{CellValue, Comment}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.drawings.TestImages
import com.tjclp.xl.macros.ref
import com.tjclp.xl.styles.units.Emu

/**
 * GH-315: source-context part mappings are keyed by STABLE SHEET IDENTITY (the name as read), not
 * by sheet index. Deleting or reordering sheets in the same write as drawing/comment edits used to
 * leave the index-keyed mappings pointing at the wrong sheets — the writer skipped drawing
 * regeneration entirely for that combination, and the index-aligned copy/rels/SST paths silently
 * corrupted output. Identity keys make delete/reorder/rename safe in one write:
 *
 *   - deleted names drop out of the mappings (removeAt drops them; a deletion also marks workbook
 *     metadata modified, so every surviving sheet regenerates against ITS OWN source part)
 *   - reorders are identity-stable (names do not move)
 *   - renames re-key the mappings (`Workbook.rename` → identity follows the sheet)
 */
class SheetIdentityMappingSpec extends FunSuite:

  private val png = ImageData(TestImages.png2x3, ImageFormat.Png)
  private val gif = ImageData(TestImages.gif2x3, ImageFormat.Gif)
  private val extent = Extent(Emu(95250L), Emu(190500L))

  private def name(s: String): SheetName = SheetName.unsafe(s)

  private def writeRead(wb: Workbook, label: String): (Path, Workbook) =
    val path = Files.createTempFile(s"xl-315-$label-", ".xlsx")
    path.toFile.deleteOnExit()
    XlsxWriter.write(wb, path).fold(err => fail(s"$label write failed: ${err.message}"), identity)
    (path, XlsxReader.read(path).fold(err => fail(s"$label read failed: ${err.message}"), identity))

  private def write(wb: Workbook, label: String): Path =
    val out = Files.createTempFile(s"xl-315-$label-out-", ".xlsx")
    out.toFile.deleteOnExit()
    XlsxWriter.write(wb, out).fold(err => fail(s"$label write failed: ${err.message}"), identity)
    out

  private def reread(path: Path): Workbook =
    XlsxReader.read(path).fold(err => fail(s"re-read failed: ${err.message}"), identity)

  private def entryNames(path: Path): Set[String] =
    val zip = new ZipFile(path.toFile)
    try zip.entries().asScala.map(_.getName).toSet
    finally zip.close()

  private def entryText(path: Path, entry: String): String =
    val zip = new ZipFile(path.toFile)
    try
      Option(zip.getEntry(entry)) match
        case Some(e) => new String(zip.getInputStream(e).readAllBytes(), "UTF-8")
        case None => fail(s"zip entry $entry not found in $path (have ${entryNames(path)})")
    finally zip.close()

  /**
   * Structural validity invariant: every internal relationship in every worksheet rels part must
   * resolve to an entry that actually ships in the output zip. Index-drifted writes used to leave
   * dangling comment/VML/drawing relationships behind.
   */
  private def assertWorksheetRelsResolve(path: Path): Unit =
    val names = entryNames(path)
    names.filter(n => n.startsWith("xl/worksheets/_rels/") && n.endsWith(".rels")).foreach {
      relsName =>
        val xml = scala.xml.XML.loadString(entryText(path, relsName))
        (xml \ "Relationship").foreach { rel =>
          if (rel \ "@TargetMode").text != "External" then
            val target = (rel \ "@Target").text
            val cleaned = if target.startsWith("/") then target.drop(1) else target
            val resolved =
              if cleaned.startsWith("xl/") then cleaned
              else
                Paths.get("xl/worksheets").resolve(cleaned).normalize().toString.replace('\\', '/')
            assert(
              names.contains(resolved),
              s"$relsName references $target -> $resolved which is MISSING from the output. " +
                s"Entries: ${names.toVector.sorted.mkString(", ")}"
            )
        }
    }

  /**
   * OPC validity: every shipped comments part needs a [Content_Types].xml Override (the "xml"
   * Default is application/xml, which does not register a comments part); every shipped VML part
   * must be covered by an Override or the "vml" extension Default.
   */
  private def assertCommentPartsRegistered(path: Path): Unit =
    val names = entryNames(path)
    val ct = scala.xml.XML.loadString(entryText(path, "[Content_Types].xml"))
    val overridden = (ct \ "Override").map(o => (o \ "@PartName").text).toSet
    val defaults = (ct \ "Default").map(d => (d \ "@Extension").text.toLowerCase).toSet
    names.filter(n => n.startsWith("xl/comments") && n.endsWith(".xml")).foreach { n =>
      assert(
        overridden.contains(s"/$n"),
        s"shipped comment part $n must have a [Content_Types].xml Override, got $overridden"
      )
    }
    names.filter(_.endsWith(".vml")).foreach { n =>
      assert(
        overridden.contains(s"/$n") || defaults.contains("vml"),
        s"shipped VML part $n must be covered by an Override or the vml Default, got $overridden"
      )
    }

  private def textAt(wb: Workbook, sheet: String, r: com.tjclp.xl.addressing.ARef): Option[String] =
    wb(name(sheet)).toOption.flatMap(_.cells.get(r)).map(_.value).collect {
      case CellValue.Text(s) => s
    }

  /** Pictures in a drawing part, robust to default-ns ("<pic>") and xdr-prefixed serialization. */
  private def picCount(drawingXml: String): Int =
    "<(xdr:)?pic[ >]".r.findAllIn(drawingXml).size

  // ===== (e) deletion alone: the surviving sheet keeps ITS content =====

  test("GH-315(e): delete-only write keeps the surviving sheet's content") {
    val alpha = Sheet(name("Alpha")).put(ref"A1" -> "AlphaContent")
    val beta = Sheet(name("Beta")).put(ref"A1" -> "BetaContent")
    val (_, wb) = writeRead(Workbook(Vector(alpha, beta)), "delete-only")

    val deleted = wb.remove(name("Alpha")).fold(e => fail(e.message), identity)
    val out = write(deleted, "delete-only")

    val result = reread(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Beta"))
    assertEquals(
      textAt(result, "Beta", ref"A1"),
      Some("BetaContent"),
      "surviving sheet must keep its own content after a delete-only write"
    )
    assertWorksheetRelsResolve(out)
  }

  // ===== (a) delete + drawing edit in ONE write =====

  test("GH-315(a): delete sheet 1 + image edit on (formerly) sheet 2 regenerates drawings") {
    val alpha = Sheet(name("Alpha")).put(ref"A1" -> "AlphaContent")
    val beta = Sheet(name("Beta")).put(ref"A1" -> "BetaContent").addImage(png, ref"B2", extent)
    val (_, wb) = writeRead(Workbook(Vector(alpha, beta)), "delete-image")

    // Source mapping sanity: Beta's drawing part was recorded at read time
    val ctx = wb.sourceContext.getOrElse(fail("source context missing"))
    assert(ctx.drawingPathMapping.nonEmpty, "reader should record Beta's drawing part")

    val edited = wb
      .remove(name("Alpha"))
      .flatMap(_.update(name("Beta"), _.addImage(gif, ref"D8", extent)))
      .fold(e => fail(e.message), identity)
    val out = write(edited, "delete-image")

    val result = reread(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Beta"))
    assertEquals(textAt(result, "Beta", ref"A1"), Some("BetaContent"))

    val drawings = result(name("Beta")).fold(e => fail(e.message), identity).drawings
    assertEquals(
      drawings.size,
      2,
      s"Beta must carry BOTH images after delete+addImage in one write, got $drawings"
    )
    val formats = drawings.collect { case p: Drawing.Picture => p.image.format }.toSet
    assertEquals(formats, Set[ImageFormat](ImageFormat.Png, ImageFormat.Gif))

    // The drawing part regenerated at its SOURCE path with both pictures
    val drawingXml = entryText(out, "xl/drawings/drawing1.xml")
    assertEquals(picCount(drawingXml), 2, s"drawing1.xml should hold two pictures:\n$drawingXml")

    assertWorksheetRelsResolve(out)
  }

  // ===== (b) reorder + comment edit in ONE write =====

  test("GH-315(b): reorder + comment edit keeps comments with their sheets, rels resolve") {
    val alpha = Sheet(name("Alpha"))
      .put(ref"A1" -> "AlphaContent")
      .comment(ref"A1", Comment.plainText("alpha note", Some("Ann")))
    val beta = Sheet(name("Beta")).put(ref"A1" -> "BetaContent")
    val (_, wb) = writeRead(Workbook(Vector(alpha, beta)), "reorder-comment")

    val edited = wb
      .reorder(Vector(name("Beta"), name("Alpha")))
      .flatMap(
        _.update(
          name("Alpha"),
          _.comment(ref"A1", Comment.plainText("alpha note edited", Some("Ann")))
        )
      )
      .fold(e => fail(e.message), identity)
    val out = write(edited, "reorder-comment")

    val result = reread(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Beta", "Alpha"))
    assertEquals(textAt(result, "Beta", ref"A1"), Some("BetaContent"))
    assertEquals(textAt(result, "Alpha", ref"A1"), Some("AlphaContent"))

    val alphaComments = result(name("Alpha")).fold(e => fail(e.message), identity).comments
    assertEquals(
      alphaComments.get(ref"A1").map(_.text.toPlainText),
      Some("alpha note edited"),
      "Alpha's edited comment must stay with Alpha after the reorder"
    )
    val betaComments = result(name("Beta")).fold(e => fail(e.message), identity).comments
    assert(betaComments.isEmpty, s"Beta must not inherit Alpha's comment, got $betaComments")

    assertWorksheetRelsResolve(out)
  }

  // ===== (d) rename + drawing edit: identity follows the rename =====

  test("GH-315(d): rename + image edit regenerates the drawing at the source part path") {
    val alpha = Sheet(name("Alpha")).put(ref"A1" -> "AlphaContent").addImage(png, ref"B2", extent)
    val (_, wb) = writeRead(Workbook(Vector(alpha)), "rename-image")

    val edited = wb
      .rename(name("Alpha"), name("Omega"))
      .flatMap(_.update(name("Omega"), _.addImage(gif, ref"D8", extent)))
      .fold(e => fail(e.message), identity)
    val out = write(edited, "rename-image")

    val result = reread(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Omega"))
    val drawings = result(name("Omega")).fold(e => fail(e.message), identity).drawings
    assertEquals(
      drawings.size,
      2,
      s"renamed sheet must keep its source drawing identity: $drawings"
    )

    // Regenerated at the SOURCE part path (no churned drawing2.xml)
    val drawingXml = entryText(out, "xl/drawings/drawing1.xml")
    assertEquals(picCount(drawingXml), 2)
    assert(
      !entryNames(out).contains("xl/drawings/drawing2.xml"),
      "rename must not churn a fresh drawing part — identity follows the rename"
    )
    assertWorksheetRelsResolve(out)
  }

  // ===== guard: delete a commented sheet + fresh comments on the survivor =====

  test("GH-315: delete commented sheet + add comments to survivor — no duplicate VML entries") {
    val alpha = Sheet(name("Alpha"))
      .put(ref"A1" -> "AlphaContent")
      .comment(ref"A1", Comment.plainText("alpha note", Some("Ann")))
    val beta = Sheet(name("Beta")).put(ref"A1" -> "BetaContent")
    val (_, wb) = writeRead(Workbook(Vector(alpha, beta)), "delete-comment")

    val edited = wb
      .remove(name("Alpha"))
      .flatMap(
        _.update(name("Beta"), _.comment(ref"B2", Comment.plainText("beta note", Some("Bob"))))
      )
      .fold(e => fail(e.message), identity)
    // A duplicate zip entry (preserved VML + freshly emitted VML at the same path) fails the write
    val out = write(edited, "delete-comment")

    val result = reread(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Beta"))
    val betaComments = result(name("Beta")).fold(e => fail(e.message), identity).comments
    assertEquals(betaComments.get(ref"B2").map(_.text.toPlainText), Some("beta note"))
    assert(
      !betaComments.contains(ref"A1"),
      s"Beta must not inherit the deleted sheet's comment, got $betaComments"
    )
    assertWorksheetRelsResolve(out)
  }

  // ===== fresh-comment allocation must not collide with a surviving sheet's mapped part =====

  test("GH-315: reorder + fresh comment on the other sheet — comment paths do not collide") {
    val alpha = Sheet(name("Alpha"))
      .put(ref"A1" -> "AlphaContent")
      .comment(ref"A1", Comment.plainText("alpha note", Some("Ann")))
    val beta = Sheet(name("Beta")).put(ref"A1" -> "BetaContent")
    val (_, wb) = writeRead(Workbook(Vector(alpha, beta)), "reorder-fresh-comment")

    // Beta lands at index 0: the fresh-comment fallback used to allocate comments1.xml — the
    // very part Alpha's identity mapping still claims — and the duplicate zip entry killed the
    // ENTIRE write with an IOError.
    val edited = wb
      .reorder(Vector(name("Beta"), name("Alpha")))
      .flatMap(
        _.update(name("Beta"), _.comment(ref"B2", Comment.plainText("beta note", Some("Bob"))))
      )
      .fold(e => fail(e.message), identity)
    val out = write(edited, "reorder-fresh-comment")

    val result = reread(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Beta", "Alpha"))
    val alphaComments = result(name("Alpha")).fold(e => fail(e.message), identity).comments
    assertEquals(
      alphaComments.get(ref"A1").map(_.text.toPlainText),
      Some("alpha note"),
      "Alpha's comment must stay with Alpha"
    )
    assertEquals(alphaComments.size, 1)
    val betaComments = result(name("Beta")).fold(e => fail(e.message), identity).comments
    assertEquals(
      betaComments.get(ref"B2").map(_.text.toPlainText),
      Some("beta note"),
      "Beta's fresh comment must land on Beta"
    )
    assertEquals(betaComments.size, 1)

    // Alpha keeps its source part; Beta's fresh part allocates ABOVE every claimed number
    // (the drawing layer's maxDrawingNum+1 precedent)
    val entries = entryNames(out)
    assert(entries.contains("xl/comments1.xml"), s"Alpha's mapped part must survive: $entries")
    assert(entries.contains("xl/comments2.xml"), s"Beta's fresh part allocates above: $entries")
    assertWorksheetRelsResolve(out)
    assertCommentPartsRegistered(out)
  }

  test("GH-315: delete a commented sheet + fresh comment on another — fresh part allocates above") {
    val aye = Sheet(name("Aye"))
      .put(ref"A1" -> "AyeContent")
      .comment(ref"A1", Comment.plainText("aye note", Some("Ann")))
    val bee = Sheet(name("Bee"))
      .put(ref"A1" -> "BeeContent")
      .comment(ref"A1", Comment.plainText("bee note", Some("Bob")))
    val cee = Sheet(name("Cee")).put(ref"A1" -> "CeeContent")
    val (_, wb) = writeRead(Workbook(Vector(aye, bee, cee)), "delete-fresh-comment")

    // Cee lands at index 1: the fresh-comment fallback used to allocate comments2.xml — the part
    // surviving Bee still claims by identity — and the duplicate zip entry killed the write.
    val edited = wb
      .remove(name("Aye"))
      .flatMap(
        _.update(name("Cee"), _.comment(ref"C3", Comment.plainText("cee note", Some("Cal"))))
      )
      .fold(e => fail(e.message), identity)
    val out = write(edited, "delete-fresh-comment")

    val result = reread(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Bee", "Cee"))
    val beeComments = result(name("Bee")).fold(e => fail(e.message), identity).comments
    assertEquals(
      beeComments.get(ref"A1").map(_.text.toPlainText),
      Some("bee note"),
      "Bee must keep its own comment"
    )
    assertEquals(beeComments.size, 1)
    val ceeComments = result(name("Cee")).fold(e => fail(e.message), identity).comments
    assertEquals(
      ceeComments.get(ref"C3").map(_.text.toPlainText),
      Some("cee note"),
      "Cee's fresh comment must land on Cee"
    )
    assertEquals(ceeComments.size, 1)

    val entries = entryNames(out)
    assert(entries.contains("xl/comments2.xml"), s"Bee keeps its source part: $entries")
    assert(entries.contains("xl/comments3.xml"), s"Cee's fresh part allocates above: $entries")
    assertWorksheetRelsResolve(out)
    assertCommentPartsRegistered(out)
  }

  // ===== combined stress: every tracked structural+content operation in ONE write =====

  test("GH-315: delete middle + reorder + image edit + comment edit + fresh comment in ONE write") {
    val aye = Sheet(name("Aye"))
      .put(ref"A1" -> "AyeContent")
      .comment(ref"A1", Comment.plainText("aye note", Some("Ann")))
    val mid = Sheet(name("Mid")).put(ref"A1" -> "MidContent")
    val cee = Sheet(name("Cee")).put(ref"A1" -> "CeeContent").addImage(png, ref"B2", extent)
    val (_, wb) = writeRead(Workbook(Vector(aye, mid, cee)), "stress")

    val edited = wb
      .remove(name("Mid"))
      .flatMap(_.reorder(Vector(name("Cee"), name("Aye"))))
      .flatMap(_.update(name("Cee"), _.addImage(gif, ref"D8", extent)))
      .flatMap(
        _.update(
          name("Aye"),
          _.comment(ref"A1", Comment.plainText("aye note edited", Some("Ann")))
        )
      )
      .flatMap(
        _.update(name("Cee"), _.comment(ref"E5", Comment.plainText("cee note", Some("Cal"))))
      )
      .fold(e => fail(e.message), identity)
    val out = write(edited, "stress")

    val result = reread(out)
    assertEquals(result.sheetNames.map(_.value), Vector("Cee", "Aye"))
    assertEquals(textAt(result, "Cee", ref"A1"), Some("CeeContent"))
    assertEquals(textAt(result, "Aye", ref"A1"), Some("AyeContent"))

    val ayeComments = result(name("Aye")).fold(e => fail(e.message), identity).comments
    assertEquals(ayeComments.get(ref"A1").map(_.text.toPlainText), Some("aye note edited"))
    assertEquals(ayeComments.size, 1)
    val ceeComments = result(name("Cee")).fold(e => fail(e.message), identity).comments
    assertEquals(ceeComments.get(ref"E5").map(_.text.toPlainText), Some("cee note"))
    assertEquals(ceeComments.size, 1)

    val drawings = result(name("Cee")).fold(e => fail(e.message), identity).drawings
    assertEquals(drawings.size, 2, s"Cee must carry both images, got $drawings")
    val formats = drawings.collect { case p: Drawing.Picture => p.image.format }.toSet
    assertEquals(formats, Set[ImageFormat](ImageFormat.Png, ImageFormat.Gif))

    assertWorksheetRelsResolve(out)
    assertCommentPartsRegistered(out)
  }
