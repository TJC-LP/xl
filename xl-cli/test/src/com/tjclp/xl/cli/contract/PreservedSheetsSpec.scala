package com.tjclp.xl.cli.contract

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}
import java.util.zip.{ZipEntry, ZipOutputStream}

import cats.effect.IO
import cats.syntax.all.*

import munit.CatsEffectSuite

private[contract] object PreservedSheetsFixture:
  private val relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  private val mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  val kinds = Vector(
    (
      "Macro",
      "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
      "macro",
      "hidden"
    ),
    (
      "Intl",
      "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
      "international-macro",
      "veryHidden"
    ),
    ("Chart", s"$relNs/chartsheet", "chart", "visible"),
    ("Dialog", s"$relNs/dialogsheet", "dialog", "visible"),
    (
      "Strict",
      "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet",
      "chart",
      "hidden"
    )
  )
  def write(path: Path, selected: Vector[(String, String, String, String)] = kinds): IO[Unit] =
    IO.blocking {
      val zip = new ZipOutputStream(Files.newOutputStream(path))
      def put(name: String, xml: String): Unit =
        zip.putNextEntry(new ZipEntry(name))
        zip.write(xml.getBytes(StandardCharsets.UTF_8))
        zip.closeEntry()
      try
        put(
          "_rels/.rels",
          s"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="$relNs/officeDocument" Target="xl/workbook.xml"/></Relationships>"""
        )
        val entries = selected.zipWithIndex.map { case ((name, _, _, state), i) =>
          s"""<sheet name="$name" sheetId="${i + 2}" r:id="r${i + 2}" state="$state"/>"""
        }.mkString
        put(
          "xl/workbook.xml",
          s"""<workbook xmlns="$mainNs" xmlns:r="$relNs"><sheets><sheet name="Data" sheetId="1" r:id="r1"/>$entries</sheets></workbook>"""
        )
        val rels = selected.zipWithIndex.map { case ((_, rel, _, _), i) =>
          s"""<Relationship Id="r${i + 2}" Type="$rel" Target="other${i + 2}.xml"/>"""
        }.mkString
        put(
          "xl/_rels/workbook.xml.rels",
          s"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="$relNs/worksheet" Target="worksheets/sheet1.xml"/>$rels</Relationships>"""
        )
        put(
          "xl/worksheets/sheet1.xml",
          s"""<worksheet xmlns="$mainNs"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>"""
        )
        selected.zipWithIndex.foreach { case ((_, _, kind, _), i) =>
          val root = kind match
            case "macro" | "international-macro" => "macrosheet"
            case "dialog" => "dialogsheet"
            case _ => "chartsheet"
          put(s"xl/other${i + 2}.xml", s"""<$root xmlns="$mainNs"><sheetData/></$root>""")
        }
        put(
          "[Content_Types].xml",
          """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>"""
        )
      finally zip.close()
    }.void

class PreservedSheetsSpec extends CatsEffectSuite:
  test(
    "GH-788: quick, streaming and stats listings expose preserved kinds without addressing them"
  ) {
    val kinds = PreservedSheetsFixture.kinds
    IO.blocking(Files.createTempFile("gh788", ".xlsm"))
      .flatTap(path => PreservedSheetsFixture.write(path))
      .bracket { path =>
        val base = List("-f", path.toString)
        List(List.empty[String], List("--stream"), List("--stats")).traverse_ { flags =>
          for
            text <- CliHarness.run(base ++ List("sheets") ++ flags, "")
            _ = assertEquals(text.exit, 0, text.stderr)
            _ = kinds.foreach { (name, _, kind, state) =>
              assert(
                text.stdout.contains(name) && text.stdout.contains(kind) && text.stdout
                  .contains(state)
              )
            }
            json <- CliHarness.run(base ++ List("sheets", "--json") ++ flags, "")
            _ = assertEquals(json.exit, 0, json.stderr)
            data = ujson.read(json.stdout)("data")
            _ = assertEquals(data("sheets").arr.map(_("name").str).toVector, Vector("Data"))
            _ = assertEquals(
              data("preservedSheets").arr.map(_("kind").str).toVector,
              kinds.map(_._3)
            )
            _ = assertEquals(
              data("preservedSheets").arr.map(_("sourceIndex").num.toInt).toVector,
              Vector(2, 3, 4, 5, 6)
            )
          yield ()
        } *> kinds.traverse_ { (name, _, _, _) =>
          List(List.empty[String], List("--stream")).traverse_ { flags =>
            CliHarness.run(base ++ flags ++ List("-s", name, "cell", "A1", "--json"), "").map {
              run =>
                assertEquals(run.exit, 3)
                assert(run.stdout.contains("SHEET_NOT_FOUND"), run.stdout)
            }
          }
        }
      }(path => IO.blocking(Files.deleteIfExists(path)).void)
  }

  test("GH-788: macro info is visible in text/JSON and gates only strict lint, in both modes") {
    IO.blocking(Files.createTempFile("gh788-macro", ".xlsm"))
      .flatTap(path => PreservedSheetsFixture.write(path, PreservedSheetsFixture.kinds.take(1)))
      .bracket { path =>
        List(List.empty[String], List("--stream")).traverse_ { flags =>
          val args = List("-f", path.toString, "lint") ++ flags
          for
            text <- CliHarness.run(args, "")
            _ = assertEquals(text.exit, 0, text.stderr)
            _ = assert(
              text.stdout.contains("[macro-sheet] (info)") && text.stdout.contains("hidden"),
              text.stdout
            )
            json <- CliHarness.run(args ++ List("--format", "json"), "")
            _ = assertEquals(json.exit, 0, json.stderr)
            data = ujson.read(json.stdout)
            _ = assertEquals(data("clean").bool, false)
            _ = assert(
              data("findings").arr
                .exists(f => f("category").str == "macro-sheet" && f("severity").str == "info")
            )
            strict <- CliHarness.run(args ++ List("--strict"), "")
            _ = assertEquals(strict.exit, 1, strict.stderr)
            _ = assert(strict.stdout.contains("1 info"), strict.stdout)
          yield ()
        }
      }(path => IO.blocking(Files.deleteIfExists(path)).void)
  }
