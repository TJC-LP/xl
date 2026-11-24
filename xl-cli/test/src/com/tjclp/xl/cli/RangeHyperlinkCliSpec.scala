package com.tjclp.xl.cli

import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipFile, ZipOutputStream}

import cats.effect.IO
import cats.syntax.traverse.*
import munit.CatsEffectSuite

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.macros.ref

/**
 * #931 / #932 through the CLI: cells under a range hyperlink report it (`cell`, `cell --json`,
 * `describe --full`), and batch `hyperlink` takes a range, to set or to clear.
 */
class RangeHyperlinkCliSpec extends CatsEffectSuite:

  private val excel = ExcelIO.instance[IO]
  private val url = "https://example.net/"
  private val sheetPart = "xl/worksheets/sheet1.xml"

  private def tmp(tag: String): Path =
    val p = Files.createTempFile(s"range-hl-$tag-", ".xlsx")
    p.toFile.deleteOnExit()
    p

  private def entryText(path: Path, name: String): String =
    val zip = new ZipFile(path.toFile)
    try
      val is = zip.getInputStream(zip.getEntry(name))
      try new String(is.readAllBytes(), "UTF-8")
      finally is.close()
    finally zip.close()

  /**
   * An Excel-shaped source: B7 linked on its own, and `D2:E3` linked as one range with a tooltip
   * (what Excel writes for a link applied to a selection). Only D2 and B7 hold values.
   */
  private def rangeSource: IO[Path] =
    val seed = tmp("seed")
    val src = tmp("src")
    val wb = Workbook(
      Sheet("Summary")
        .put(Cell(ref"B7", CellValue.Text("Example link")).withHyperlink("https://example.com/"))
        .put(Cell(ref"D2", CellValue.Text("Range link")).withHyperlink(url))
    )
    excel.write(wb, seed) *> IO.blocking {
      val in = new ZipFile(seed.toFile)
      val out = new ZipOutputStream(Files.newOutputStream(src))
      try
        in.entries().asIterator().forEachRemaining { e =>
          val bytes = in.getInputStream(e).readAllBytes()
          val body =
            if e.getName != sheetPart then bytes
            else
              new String(bytes, "UTF-8")
                .replace("""ref="D2"""", """ref="D2:E3" tooltip="Docs"""")
                .getBytes("UTF-8")
          out.putNextEntry(new ZipEntry(e.getName))
          out.write(body)
          out.closeEntry()
        }
      finally
        out.close()
        in.close()
      src
    }

  private def run(args: String*): IO[CliRun] = CliHarness.run(args*)

  private def batch(src: Path, out: Path, json: String, extra: String*): IO[CliRun] =
    CliHarness.run(
      List("-f", src.toString, "-s", "Summary", "-o", out.toString) ++ extra ++
        List("batch", "-"),
      json
    )

  private def linkLine(r: CliRun): Option[String] =
    r.stdout.linesIterator.find(_.startsWith("Hyperlink:"))

  // ===== #931: reads =====

  test("#931: cell reports the range link on every covered cell, and only there") {
    for
      src <- rangeSource
      covered <- List("D2", "E2", "D3", "E3").traverse(a1 => run("-f", src.toString, "cell", a1))
      outside <- run("-f", src.toString, "cell", "F2")
    yield
      covered.foreach { r =>
        assertEquals(r.exit, 0, r.stderr)
        assertEquals(linkLine(r), Some(s"Hyperlink: $url"), r.stdout)
      }
      assertEquals(linkLine(outside), None, outside.stdout)
  }

  test("#931: cell --json carries the range link instead of null") {
    for
      src <- rangeSource
      r <- run("-f", src.toString, "cell", "E3", "--json")
    yield
      assertEquals(r.exit, 0, r.stderr)
      val json = ujson.read(r.stdout)
      val hyperlink =
        json.obj.get("hyperlink").orElse(json.obj.get("data").flatMap(_.obj.get("hyperlink")))
      assertEquals(hyperlink.map(_.str), Some(url), r.stdout)
  }

  test("#931: describe --full counts the range link once, beside the cell link") {
    for
      src <- rangeSource
      r <- run("-f", src.toString, "-s", "Summary", "describe", "--full")
    yield
      assertEquals(r.exit, 0, r.stderr)
      assert(r.stdout.contains("hyperlinks 2"), r.stdout)
  }

  // ===== #932: batch hyperlink over a range =====

  test("#932: batch hyperlink over a range writes one range link that every cell reports") {
    val src = tmp("plain")
    val out = tmp("linked")
    for
      _ <- excel.write(Workbook(Sheet("Summary").put(ref"A1", CellValue.Text("Top"))), src)
      r <- batch(src, out, s"""[{"op":"hyperlink","ref":"A1:A5","target":"$url"}]""")
      cells <- List("A1", "A3", "A5").traverse(a1 => run("-f", out.toString, "cell", a1))
      after <- run("-f", out.toString, "cell", "A6")
    yield
      assertEquals(r.exit, 0, r.stderr)
      cells.foreach(c => assertEquals(linkLine(c), Some(s"Hyperlink: $url"), c.stdout))
      assertEquals(linkLine(after), None)
      val xml = entryText(out, sheetPart)
      assert(xml.contains("""ref="A1:A5""""), xml)
      assert(!xml.contains("""<c r="A3""""), s"a range link must not create cells:\n$xml")
  }

  test("#932: batch hyperlink with no target clears every cell of the range") {
    val src = tmp("plain2")
    val linked = tmp("linked2")
    val cleared = tmp("cleared2")
    for
      _ <- excel.write(Workbook(Sheet("Summary").put(ref"A1", CellValue.Text("Top"))), src)
      _ <- batch(src, linked, s"""[{"op":"hyperlink","ref":"A1:A5","target":"$url"}]""")
      r <- batch(linked, cleared, """[{"op":"hyperlink","ref":"A2:A3"}]""")
      wb <- excel.read(cleared)
    yield
      assertEquals(r.exit, 0, r.stderr)
      val s = wb.sheets.head
      assertEquals(s(ref"A2").hyperlink, None)
      assertEquals(s(ref"A3").hyperlink, None)
      assertEquals(s(ref"A1").hyperlink, Some(Hyperlink(url)))
      assertEquals(s(ref"A5").hyperlink, Some(Hyperlink(url)))
  }

  for (whole, written, probe, outside) <- Seq(
      ("A:A", "A1:A1048576", "A1048576", "B1"),
      ("Summary!B:C", "B1:C1048576", "C70000", "D1"),
      ("3:3", "A3:XFD3", "XFD3", "A4")
    )
  do
    test(s"#932: batch hyperlink over $whole writes one range link and creates no cells") {
      val src = tmp("whole")
      val out = tmp("whole-linked")
      for
        _ <- excel.write(Workbook(Sheet("Summary").put(ref"A1", CellValue.Text("Top"))), src)
        r <- batch(src, out, s"""[{"op":"hyperlink","ref":"$whole","target":"$url"}]""")
        in <- run("-f", out.toString, "cell", probe)
        off <- run("-f", out.toString, "cell", outside)
        wb <- excel.read(out)
      yield
        assertEquals(r.exit, 0, r.stderr)
        assertEquals(linkLine(in), Some(s"Hyperlink: $url"), in.stdout)
        assertEquals(linkLine(off), None, off.stdout)
        val xml = entryText(out, sheetPart)
        assert(xml.contains(s"""ref="$written""""), xml)
        val s = wb.sheets.head
        assertEquals(s.rangeHyperlinks.size, 1)
        assertEquals(s.cells.size, 1, "a whole-column link must not materialise cells")
    }

  test("#932: batch hyperlink with no target clears a whole column") {
    val src = tmp("whole-src")
    val linked = tmp("whole-linked2")
    val cleared = tmp("whole-cleared")
    for
      _ <- excel.write(Workbook(Sheet("Summary").put(ref"A1", CellValue.Text("Top"))), src)
      _ <- batch(src, linked, s"""[{"op":"hyperlink","ref":"A1:B9","target":"$url"}]""")
      r <- batch(linked, cleared, """[{"op":"hyperlink","ref":"A:A"}]""")
      wb <- excel.read(cleared)
    yield
      assertEquals(r.exit, 0, r.stderr)
      val s = wb.sheets.head
      assertEquals(s(ref"A5").hyperlink, None)
      assertEquals(s(ref"B5").hyperlink, Some(Hyperlink(url)))
  }

  test("#932: batch hyperlink is refused under --stream before any write") {
    val src = tmp("plain3")
    val out = tmp("streamed3")
    for
      _ <- excel.write(Workbook(Sheet("Summary").put(ref"A1", CellValue.Text("Top"))), src)
      _ <- IO.blocking(Files.deleteIfExists(out))
      r <- batch(src, out, s"""[{"op":"hyperlink","ref":"A1:A5","target":"$url"}]""", "--stream")
    yield
      assertNotEquals(r.exit, 0, r.stdout)
      assert(!Files.exists(out), "a refused stream batch must not write")
  }

  // ===== real edits over covered cells =====

  private def hyperlinksElem(path: Path): String =
    "(?s)<hyperlinks\\b.*?</hyperlinks>".r.findFirstIn(entryText(path, sheetPart)).getOrElse("")

  for (label, json) <- Seq(
      "batch style over the range" -> """[{"op":"style","range":"D2:E3","bold":true}]""",
      "batch put into a covered cell" -> """[{"op":"put","ref":"E2","value":1}]""",
      "batch styles-only clear" -> """[{"op":"clear","range":"E2","styles":true}]"""
    )
  do
    test(s"#931: $label keeps the range link as written") {
      val out = tmp("edit")
      for
        src <- rangeSource
        r <- batch(src, out, json)
        wb <- excel.read(out)
      yield
        assertEquals(r.exit, 0, r.stderr)
        assertEquals(hyperlinksElem(out), hyperlinksElem(src))
        val s = wb.sheets.head
        for a1 <- List(ref"D2", ref"E2", ref"D3", ref"E3") do
          assertEquals(s.cells.hyperlinkAt(a1), None, s"${a1.toA1} must not own the link")
          assertEquals(s(a1).hyperlink, Some(Hyperlink(url)), a1.toA1)
    }

  for (label, args) <- Seq(
      "the clear verb" -> List("clear", "E2"),
      "batch clear" -> List("batch", "-")
    )
  do
    test(s"#931: a contents clear through $label removes the link from the cleared cell") {
      val out = tmp("clear")
      for
        src <- rangeSource
        r <- CliHarness.run(
          List("-f", src.toString, "-s", "Summary", "-o", out.toString) ++ args,
          """[{"op":"clear","range":"E2"}]"""
        )
        wb <- excel.read(out)
      yield
        assertEquals(r.exit, 0, r.stderr)
        val s = wb.sheets.head
        assertEquals(s(ref"E2").hyperlink, None)
        for a1 <- List(ref"D2", ref"D3", ref"E3") do
          assertEquals(s(a1).hyperlink, Some(Hyperlink(url)), a1.toA1)
    }

  // ===== round trip: write -> read -> every covered cell -> edit one -> split =====

  test("#931: round trip keeps the range, and clearing one covered cell splits it") {
    val edited = tmp("edited")
    val split = tmp("split")
    for
      src <- rangeSource
      e <- batch(src, edited, """[{"op":"put","ref":"A1","value":"edit"}]""")
      covered <- List("D2", "E2", "D3", "E3").traverse(a1 => run("-f", edited.toString, "cell", a1))
      s <- batch(edited, split, """[{"op":"hyperlink","ref":"E2"}]""")
      wb <- excel.read(split)
    yield
      assertEquals(e.exit, 0, e.stderr)
      covered.foreach(c => assertEquals(linkLine(c), Some(s"Hyperlink: $url"), c.stdout))
      assertEquals(s.exit, 0, s.stderr)
      val sheet = wb.sheets.head
      assertEquals(sheet(ref"E2").hyperlink, None)
      for a1 <- List(ref"D2", ref"D3", ref"E3") do
        assertEquals(sheet(a1).hyperlink, Some(Hyperlink(url)), a1.toA1)
      assertEquals(sheet(ref"B7").hyperlink, Some(Hyperlink("https://example.com/")))
      // the remnants keep the source element's tooltip
      assert(
        sheet.rangeHyperlinks.nonEmpty && sheet.rangeHyperlinks.forall(_.tooltip.contains("Docs"))
      )
      assert(!sheet.contains(ref"E2"), "the split must not create E2")
  }
