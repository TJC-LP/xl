package com.tjclp.xl.cli.batch

import java.nio.file.Path

import cats.effect.{IO, Resource}
import munit.CatsEffectSuite

import com.tjclp.xl.cli.contract.{CliHarness, CliRun, TestFixtures}

/**
 * #752: every sheet name in batch JSON — an op's `sheet` key and a target ref's qualifier — and
 * every sheet verb's argument match a sheet as `-s` does: the exact spelling first, else the one
 * sheet equal to it ignoring case. One rule for every op, in memory and streamed.
 */
class BatchSheetCaseSpec extends CatsEffectSuite:

  private val fixtures = ResourceSuiteLocalFixture(
    "batch-sheet-case-fixtures",
    Resource.make(TestFixtures.materialize)(TestFixtures.delete)
  )

  override def munitFixtures = List(fixtures)

  /** `simple.xlsx`: sheets `Data` and `Summary`. */
  private def simple: String = fixtures().resolve("simple.xlsx").toString

  private def fresh(name: String): Path = fixtures().resolve(name)

  private def batch(out: Path, json: String, flags: String*): IO[CliRun] =
    CliHarness.run(
      List("-f", simple, "-o", out.toString, "--json") ++ flags.toList ++ List("batch", "-"),
      json
    )

  private def cell(path: Path, ref: String): IO[String] =
    CliHarness.run("-f", path.toString, "--json", "cell", ref).map { run =>
      assertEquals(run.exit, 0, run.stderr)
      run.stdout
    }

  private def sheetNames(path: Path): IO[String] =
    CliHarness.run("-f", path.toString, "sheets").map(_.stdout)

  test("put, putf, freeze and copy read a case-variant qualifier or sheet key as the tab") {
    val out = fresh("batch-case-mixed.xlsx")
    val json =
      """[
        |{"op":"put","ref":"data!F1","value":7},
        |{"op":"put","sheet":"summary","ref":"F2","value":3},
        |{"op":"put","sheet":"Data","ref":"DATA!F3","value":4},
        |{"op":"putf","ref":"data!G1:G3","value":"=F1*2","from":"G1"},
        |{"op":"freeze","ref":"data!B2"},
        |{"op":"copy","source":"data!F1","target":"SUMMARY!F9"}
        |]""".stripMargin
    for
      run <- batch(out, json, "-s", "Summary")
      f1 <- cell(out, "Data!F1")
      f2 <- cell(out, "Summary!F2")
      f3 <- cell(out, "Data!F3")
      g1 <- cell(out, "Data!G1")
      f9 <- cell(out, "Summary!F9")
      listed <- sheetNames(out)
    yield
      assertEquals(run.exit, 0, run.stderr + run.stdout)
      assert(f1.contains("7"), f1)
      assert(f2.contains("3"), f2)
      assert(f3.contains("4"), f3)
      assert(g1.contains("14"), g1)
      assert(f9.contains("7"), f9)
      assert(!listed.contains("data") && !listed.contains("DATA"), s"no new tab: $listed")
  }

  test("a qualifier and a sheet key naming different sheets are still a contradiction") {
    val out = fresh("batch-case-contradiction.xlsx")
    batch(out, """[{"op":"put","sheet":"Summary","ref":"data!A1","value":1}]""").map { run =>
      // a malformed batch document: exit 2, before the workbook is touched
      assertEquals(run.exit, 2, run.stderr)
      val error = ujson.read(run.stdout)("error")
      assertEquals(error("code"), ujson.Str("BATCH_OP_INVALID"))
      assert(error("message").str.contains("names sheet 'data'"), error("message").str)
    }
  }

  test("a streamed batch places case-variant qualifiers and sheet keys on the one tab") {
    val withS = fresh("batch-case-stream-s.xlsx")
    val withoutS = fresh("batch-case-stream.xlsx")
    for
      a <- batch(
        withS,
        """[{"op":"put","ref":"data!F1","value":5},{"op":"put","sheet":"DATA","ref":"F2","value":6}]""",
        "-s",
        "Data",
        "--stream"
      )
      b <- batch(
        withoutS,
        """[{"op":"put","ref":"data!F1","value":5},{"op":"put","ref":"Data!F2","value":6}]""",
        "--stream"
      )
      a1 <- cell(withS, "Data!F2")
      b1 <- cell(withoutS, "Data!F1")
    yield
      assertEquals(a.exit, 0, a.stderr + a.stdout)
      assertEquals(b.exit, 0, b.stderr + b.stdout)
      assert(a1.contains("6"), a1)
      assert(b1.contains("5"), b1)
  }

  test("sheet verbs take a sheet argument in any case") {
    val removed = fresh("verb-case-remove.xlsx")
    val moved = fresh("verb-case-move.xlsx")
    val copied = fresh("verb-case-copy.xlsx")
    val renamed = fresh("verb-case-rename.xlsx")
    def run(out: Path, verb: String*) =
      CliHarness.run(List("-f", simple, "-o", out.toString) ++ verb.toList, "")
    for
      r <- run(removed, "remove-sheet", "summary")
      m <- run(moved, "move-sheet", "DATA", "--after", "summary")
      c <- run(copied, "copy-sheet", "data", "Copy")
      n <- run(renamed, "rename-sheet", "data", "Inputs")
      afterRemove <- sheetNames(removed)
      afterMove <- sheetNames(moved)
      afterCopy <- sheetNames(copied)
      afterRename <- sheetNames(renamed)
    yield
      for v <- List(r, m, c, n) do assertEquals(v.exit, 0, v.stderr + v.stdout)
      assert(!afterRemove.contains("Summary"), afterRemove)
      assert(afterMove.indexOf("Summary") < afterMove.indexOf("Data"), afterMove)
      assert(afterCopy.contains("Copy"), afterCopy)
      assert(afterRename.contains("Inputs") && !afterRename.contains("Data"), afterRename)
  }

  test("a rename-sheet naming the default in another case moves the default with it") {
    val out = fresh("batch-case-rename-default.xlsx")
    val json =
      """[{"op":"rename-sheet","from":"data","to":"Inputs"},{"op":"put","ref":"F1","value":8}]"""
    for
      run <- batch(out, json, "-s", "Data")
      f1 <- cell(out, "Inputs!F1")
    yield
      assertEquals(run.exit, 0, run.stderr + run.stdout)
      assert(f1.contains("8"), f1)
  }
