package com.tjclp.xl.ooxml

import java.io.{ByteArrayInputStream, ByteArrayOutputStream}
import java.nio.charset.StandardCharsets
import java.util.zip.{ZipEntry, ZipOutputStream}

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{CellError, CellValue}
import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * The one rule every reader and the writer's reference scan use to read a `t="s"` value as a table
 * position (SharedStringIndex): optional XML whitespace, an optional `+` or `-`, ASCII digits up to
 * Int.MaxValue, optional XML whitespace, and no negative value (`-0` is 0). The fixed cases are the
 * spellings desktop Excel was observed to open as an entry (no repair prompt) and to repair.
 */
class SharedStringIndexSpec extends ScalaCheckSuite:

  // ========== reference rule ==========

  private val Spec = "[ \t\r\n]*([+-]?)([0-9]+)[ \t\r\n]*".r

  /** The rule written as a regular expression and a BigInt bound. */
  private def reference(value: String): Option[Int] = value match
    case Spec(sign, digits) =>
      Some(BigInt(digits)).filter(n => n <= Int.MaxValue && (sign != "-" || n == 0)).map(_.toInt)
    case _ => None

  test("spellings Excel opens as an entry name the position") {
    val cases = Vector(
      "0" -> 0,
      "3" -> 3,
      "+3" -> 3,
      "+03" -> 3,
      "03" -> 3,
      " 3 " -> 3,
      " +3 " -> 3,
      "\t3" -> 3,
      "3\n" -> 3,
      "3\r" -> 3,
      "\t\r\n3\n" -> 3,
      "-0" -> 0,
      "-00" -> 0,
      "2147483647" -> Int.MaxValue,
      ("0" * 40 + "3") -> 3
    )
    cases.foreach((v, n) => assertEquals(SharedStringIndex.parse(v), Some(n), s"[$v]"))
  }

  test("other spellings name no position (Excel repairs all but an empty value)") {
    val cases = Vector(
      "",
      "  ",
      "+",
      "-",
      "-1",
      "++3",
      "+-3",
      "--0",
      "+ 3",
      "3.0",
      "3.",
      "3.000",
      "3.5",
      "3.01",
      ".0",
      "3e0",
      "3 3",
      "0x3",
      "٣", // ARABIC-INDIC DIGIT THREE
      "３", // FULLWIDTH DIGIT THREE
      " 3", // NO-BREAK SPACE is not XML whitespace
      "2147483648",
      "9" * 30
    )
    cases.foreach(v => assertEquals(SharedStringIndex.parse(v), None, s"[$v]"))
  }

  test("the accumulator reads a value split at any point the same as whole") {
    Vector("+03", " -0 ", " 3.0 ", "-1", "٣", ("0" * 25) + "17", "2147483648").foreach { v =>
      (0 to v.length).foreach { cut =>
        val acc = new SharedStringIndex.Accumulator
        val chars = v.toCharArray
        acc.append(chars, 0, cut)
        acc.append(chars, cut, chars.length - cut)
        assertEquals(acc.result, SharedStringIndex.parse(v), s"[$v] cut=$cut")
      }
    }
  }

  /** Strings over the characters that matter to the rule, plus a few that never belong. */
  private val genValue: Gen[String] =
    val char = Gen.frequency(
      12 -> Gen.choose('0', '9'),
      3 -> Gen.const('0'),
      2 -> Gen.oneOf(' ', '\t', '\n', '\r'),
      1 -> Gen.oneOf('+', '-', '.', 'e', 'x', '٣', '３', ' ')
    )
    Gen.frequency(
      8 -> Gen.choose(0, 14).flatMap(n => Gen.listOfN(n, char)).map(_.mkString),
      2 -> Gen.choose(0, 12).map(_.toString),
      1 -> Gen.zip(Gen.choose(0, 30), Gen.choose(0, 12)).map((z, n) => "0" * z + n)
    )

  property("parse agrees with the reference rule") {
    Prop.forAll(genValue)(v => assertEquals(SharedStringIndex.parse(v), reference(v), s"[$v]"))
  }

  // ========== readers and scanner agree ==========

  private val table = Vector("e0", "e1", "e2", "e3", "e4", "e5")

  property("the reference scan and XlsxReader read each value as the same position") {
    Prop.forAll(Gen.nonEmptyListOf(genValue).map(_.take(12).toVector), Gen.choose(0, table.size)) {
      (values, size) =>
        val strings = table.take(size)
        val scan = SharedStringCompaction
          .scan(new ByteArrayInputStream(sheetXml(values).getBytes(StandardCharsets.UTF_8)), size)
          .fold(e => fail(e), identity)
        val named = values.flatMap(SharedStringIndex.parse(_).filter(_ < size))
        assertEquals(scan.indices, scala.collection.immutable.BitSet.fromSpecific(named))
        assertEquals(scan.outOfRange, values.size - named.size)

        val sheet = XlsxReader
          .readFromBytes(book(strings, values))
          .fold(e => fail(s"read failed: $e"), identity)(SheetName.unsafe("Sheet1"))
          .fold(e => fail(e.message), identity)
        values.zipWithIndex.foreach { (v, i) =>
          val expected = SharedStringIndex.parse(v) match
            case Some(n) if n < size => CellValue.Text(strings(n))
            case Some(_) => CellValue.Error(CellError.Ref)
            case None => CellValue.Error(CellError.Value)
          val r = ARef.parse(s"A${i + 1}").fold(e => fail(e), identity)
          assertEquals(sheet.cells.get(r).map(_.value), Some(expected), s"[$v] size=$size")
        }
    }
  }

  // ========== fixture ==========

  private def xmlText(v: String): String =
    v.flatMap {
      case '\r' => "&#13;" // a literal CR would be normalized to LF by the XML parser
      case c => c.toString
    }

  private def sheetXml(values: Vector[String]): String =
    val rows = values.zipWithIndex
      .map((v, i) => s"""<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>${xmlText(v)}</v></c></row>""")
      .mkString
    s"""<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="$ns"><sheetData>$rows</sheetData></worksheet>"""

  private val ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

  /** One sheet whose A1, A2, ... are `t="s"` cells holding `values`, over a table of `strings`. */
  private def book(strings: Vector[String], values: Vector[String]): Array[Byte] =
    val rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    val bytes = new ByteArrayOutputStream()
    val out = new ZipOutputStream(bytes)
    def put(name: String, content: String): Unit =
      out.putNextEntry(new ZipEntry(name))
      out.write(content.getBytes(StandardCharsets.UTF_8))
      out.closeEntry()
    put(
      "[Content_Types].xml",
      """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">""" +
        """<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>""" +
        """<Default Extension="xml" ContentType="application/xml"/>""" +
        """<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>""" +
        """<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>""" +
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
        """<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"""
    )
    put(
      "xl/_rels/workbook.xml.rels",
      """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">""" +
        s"""<Relationship Id="rId1" Type="$rel/worksheet" Target="worksheets/sheet1.xml"/>""" +
        s"""<Relationship Id="rId2" Type="$rel/sharedStrings" Target="sharedStrings.xml"/>""" +
        "</Relationships>"
    )
    put(
      "xl/sharedStrings.xml",
      s"""<?xml version="1.0" encoding="UTF-8"?><sst xmlns="$ns" count="${values.size}" uniqueCount="${strings.size}">""" +
        strings.map(s => s"<si><t>$s</t></si>").mkString + "</sst>"
    )
    put("xl/worksheets/sheet1.xml", sheetXml(values))
    out.close()
    bytes.toByteArray
