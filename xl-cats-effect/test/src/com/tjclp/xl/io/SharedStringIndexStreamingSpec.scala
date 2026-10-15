package com.tjclp.xl.io

import java.io.ByteArrayInputStream
import java.nio.charset.StandardCharsets

import cats.effect.IO
import fs2.Stream
import munit.CatsEffectSuite
import org.scalacheck.Gen
import org.scalacheck.rng.Seed

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.ooxml.{SharedStringIndex, SharedStrings}

/**
 * Both streaming readers read a `t="s"` value with SharedStringIndex, the rule XlsxReader and the
 * writer's reference scan use: a value names entry n when the rule reads it as n and n is below the
 * table's size; any other value reads as empty (the streaming readers' blank for a reference that
 * resolves to nothing).
 */
class SharedStringIndexStreamingSpec extends CatsEffectSuite:

  private val table = SharedStrings.fromStrings(Vector("e0", "e1", "e2", "e3"))

  private val fixed = Vector(
    "0",
    "3",
    "+3",
    "+03",
    " 3 ",
    "3.0",
    "4",
    "-0",
    "-1",
    "3.5",
    "٣",
    "３",
    "",
    ("0" * 30) + "2",
    "2147483648"
  )

  private val sampled: Vector[String] =
    val char = Gen.frequency(
      12 -> Gen.choose('0', '5'),
      2 -> Gen.oneOf(' ', '\t', '\n'),
      1 -> Gen.oneOf('+', '-', '.', 'x', '٣', '３')
    )
    Gen
      .listOfN(300, Gen.choose(0, 6).flatMap(n => Gen.listOfN(n, char)).map(_.mkString))
      .apply(Gen.Parameters.default, Seed(768L))
      .getOrElse(Vector.empty)
      .toVector

  private val values = fixed ++ sampled

  private val sheetXml: String =
    val rows = values.zipWithIndex
      .map((v, i) => s"""<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>$v</v></c></row>""")
      .mkString
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
      s"<sheetData>$rows</sheetData></worksheet>"

  private def expected(v: String): CellValue =
    SharedStringIndex
      .parse(v)
      .flatMap(table.apply)
      .map(table.toCellValue)
      .getOrElse(CellValue.Empty)

  private def byRow(rows: List[RowData]): Map[Int, CellValue] =
    rows.flatMap(r => r.cells.get(0).map(r.rowIndex -> _)).toMap

  private def check(name: String, read: IO[List[RowData]]): IO[Unit] =
    read.map { rows =>
      val got = byRow(rows)
      values.zipWithIndex.foreach { (v, i) =>
        assertEquals(got.getOrElse(i + 1, CellValue.Empty), expected(v), s"$name [$v]")
      }
    }

  test("SaxStreamingReader reads each t=\"s\" value as the shared rule does") {
    check(
      "sax",
      SaxStreamingReader
        .parseWorksheetStream[IO](
          new ByteArrayInputStream(sheetXml.getBytes(StandardCharsets.UTF_8)),
          Some(table)
        )
        .compile
        .toList
    )
  }

  test("StreamingXmlReader reads each t=\"s\" value as the shared rule does") {
    check(
      "fs2-xml",
      StreamingXmlReader
        .parseWorksheetStream[IO](
          Stream.emits(sheetXml.getBytes(StandardCharsets.UTF_8)).covary[IO],
          Some(table)
        )
        .compile
        .toList
    )
  }
