package com.tjclp.xl.cli

import java.nio.file.{Files, Path}

import cats.effect.IO
import cats.effect.unsafe.implicits.global

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.contract.{CliHarness, CliRun}
import com.tjclp.xl.io.ExcelIO

import munit.FunSuite

/**
 * GH-936: a formula authored through `putf` (one cell, a drag, an explicit list; in memory and
 * under `--stream`) or batch `putf` stores each sheet qualifier as the tab's own name, as Excel
 * does on entry: `putf C1 "=sheet1!A1"` stores `=Sheet1!A1`. Unknown sheets and external references
 * stay as typed; a formula already in a file is never rewritten.
 */
class SheetQualifierCaseCliSpec extends FunSuite:

  private def withDir[A](f: Path => A): A =
    val dir = Files.createTempDirectory("sheet-qualifier-case")
    try f(dir)
    finally
      Files.list(dir).forEach(p => Files.deleteIfExists(p))
      Files.deleteIfExists(dir)

  private def num(n: Int): CellValue = CellValue.Number(BigDecimal(n))

  private val book: Workbook = Workbook(
    Vector(
      Sheet("Sheet1").put(ARef.from1(1, 1), num(10)).put(ARef.from1(1, 2), num(20)),
      Sheet("My Data")
        .put(ARef.from1(1, 1), num(5))
        // a formula already in the file, typed in another case: never rewritten
        .put(ARef.from1(2, 1), CellValue.Formula("sheet1!A1", Some(num(10))))
    )
  )

  private def write(wb: Workbook, path: Path): Unit =
    ExcelIO.instance[IO].write(wb, path).unsafeRunSync()

  private def readSheet(path: Path, name: String): Sheet =
    ExcelIO
      .instance[IO]
      .read(path)
      .unsafeRunSync()
      .sheets
      .find(_.name.value == name)
      .getOrElse(fail(s"no sheet $name"))

  private def run(args: String*): CliRun = CliHarness.run(args*).unsafeRunSync()

  private def runIn(stdin: String, args: String*): CliRun =
    CliHarness.run(args.toList, stdin).unsafeRunSync()

  private def formulaAt(s: Sheet, a1: String): (String, Option[CellValue]) =
    ARef.parse(a1).toOption.flatMap(s.cells.get).map(_.value) match
      case Some(CellValue.Formula(text, cache, _)) => (text, cache)
      case other => fail(s"$a1 is not a formula: $other")

  private def streamFlags(stream: Boolean): List[String] = if stream then List("--stream") else Nil

  test("GH-936: putf stores the qualifier as the sheet's own name; the value is unchanged") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      write(book, in)
      val r = run("-f", in.toString, "-o", out.toString, "-s", "Sheet1", "putf", "C1", "=sheet1!A1")
      assertEquals(r.exit, 0, r.stderr)
      val s = readSheet(out, "Sheet1")
      assertEquals(formulaAt(s, "C1"), ("Sheet1!A1", Some(num(10))))
    }
  }

  for stream <- List(false, true) do
    test(s"GH-936: putf drag, explicit list, quoted names (stream=$stream)") {
      withDir { dir =>
        val in = dir.resolve("in.xlsx")
        val out1 = dir.resolve("drag.xlsx")
        val out2 = dir.resolve("list.xlsx")
        write(book, in)
        val drag = run(
          (List("-f", in.toString, "-o", out1.toString, "-s", "Sheet1") ++ streamFlags(stream) ++
            List("putf", "C1:C2", "=SHEET1!A1+'my data'!$A$1"))*
        )
        assertEquals(drag.exit, 0, drag.stderr)
        val d = readSheet(out1, "Sheet1")
        assertEquals(formulaAt(d, "C1")._1, "Sheet1!A1+'My Data'!$A$1")
        assertEquals(formulaAt(d, "C2")._1, "Sheet1!A2+'My Data'!$A$1")
        val list = run(
          (List("-f", in.toString, "-o", out2.toString, "-s", "Sheet1") ++ streamFlags(stream) ++
            List("putf", "D1:D2", "=sheet1!A1", "=nosuch!A1"))*
        )
        assertEquals(list.exit, 0, list.stderr)
        val l = readSheet(out2, "Sheet1")
        assertEquals(formulaAt(l, "D1")._1, "Sheet1!A1")
        assertEquals(formulaAt(l, "D2")._1, "nosuch!A1")
        // the formula already in the file keeps its spelling
        assertEquals(formulaAt(readSheet(out2, "My Data"), "B1")._1, "sheet1!A1")
      }
    }

  for stream <- List(false, true) do
    test(s"GH-936: batch putf (one cell, drag, values) respells qualifiers (stream=$stream)") {
      withDir { dir =>
        val in = dir.resolve("in.xlsx")
        val out = dir.resolve("out.xlsx")
        write(book, in)
        val ops =
          """[{"op":"putf","ref":"E1","value":"=sheet1!A1*2"},""" +
            """{"op":"putf","ref":"F1:F2","value":"=SHEET1!A1","from":"F1"},""" +
            """{"op":"putf","ref":"G1:G2","values":["='MY DATA'!A1","=[1]sheet1!A1"]}]"""
        val r = runIn(
          ops,
          (List("-f", in.toString, "-o", out.toString, "-s", "Sheet1") ++ streamFlags(stream) ++
            List("batch", "-"))*
        )
        assertEquals(r.exit, 0, r.stderr)
        val s = readSheet(out, "Sheet1")
        assertEquals(formulaAt(s, "E1")._1, "Sheet1!A1*2")
        assertEquals(formulaAt(s, "F1")._1, "Sheet1!A1")
        assertEquals(formulaAt(s, "F2")._1, "Sheet1!A2")
        assertEquals(formulaAt(s, "G1")._1, "'My Data'!A1")
        assertEquals(formulaAt(s, "G2")._1, "[1]sheet1!A1")
        if !stream then assertEquals(formulaAt(s, "E1")._2, Some(num(20)))
      }
    }

  test("GH-936: a batch putf after add-sheet respells to the new tab's name") {
    withDir { dir =>
      val in = dir.resolve("in.xlsx")
      val out = dir.resolve("out.xlsx")
      write(book, in)
      val ops =
        """[{"op":"add-sheet","name":"Inputs"},""" +
          """{"op":"put","sheet":"Inputs","ref":"A1","value":3},""" +
          """{"op":"putf","sheet":"Sheet1","ref":"H1","value":"=inputs!A1+1"}]"""
      val r = runIn(ops, "-f", in.toString, "-o", out.toString, "batch", "-")
      assertEquals(r.exit, 0, r.stderr)
      assertEquals(formulaAt(readSheet(out, "Sheet1"), "H1"), ("Inputs!A1+1", Some(num(4))))
    }
  }
