package com.tjclp.xl.cli

import java.nio.file.Files

import cats.effect.{IO, unsafe}

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cli.commands.WriteCommands
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.writer.WriterConfig

import munit.FunSuite

/** Name edits must invalidate formula readers, even when no cell value or formula text changed. */
class BatchNameRecalcSpec extends FunSuite:
  given unsafe.IORuntime = unsafe.IORuntime.global

  private def formula(text: String, cache: Int): CellValue =
    CellValue.Formula(text, Some(CellValue.Number(cache)))

  private def run(
    wb: Workbook,
    ops: String,
    policy: WritePolicy = WritePolicy.default
  ): (Workbook, String) =
    val out = Files.createTempFile("batch-name-recalc", ".xlsx")
    try
      val summary = WriteCommands
        .batch(
          wb,
          wb.sheets.headOption,
          "-",
          out,
          WriterConfig.default,
          policy = policy,
          stdin = IO.pure(ops)
        )
        .unsafeRunSync()
      (ExcelIO.instance[IO].read(out).unsafeRunSync(), summary)
    finally Files.deleteIfExists(out)

  private val changeTax = """{"op":"define-name","name":"Tax","refersTo":"0.2"}"""

  for mixed <- List(false, true); noRecalc <- List(false, true) do
    test(
      s"a name replacement refreshes its readers unless --no-recalc; mixed=$mixed noRecalc=$noRecalc"
    ) {
      val wb = Workbook(Sheet("Data").put(ref"A1", formula("Tax*100", 10)))
        .withDefinedName("Tax", "0.1")
      val extra = if mixed then """,{"op":"put","ref":"C1","value":1}""" else ""
      val (result, summary) =
        run(wb, "[" + changeTax + extra + "]", WritePolicy(noRecalc = noRecalc))
      assertEquals(
        result.sheets(0)(ref"A1").effectiveValue,
        CellValue.Number(if noRecalc then 10 else 20),
        summary
      )
    }

  test(
    "case-insensitive aliases invalidate transitive cross-sheet readers and preserve unrelated caches"
  ) {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"A1", formula("EffectiveTax*100", 10))
        .put(ref"B1", formula("A1*2", 20))
        .put(ref"D1", formula("OtherRate*10", 999))
        .put(ref"E1", formula("NOSUCHFN(1)", 777)),
      Sheet("Summary").put(ref"A1", formula("Data!B1+1", 21))
    ).withDefinedName("Tax", "0.1")
      .withDefinedName("EffectiveTax", "tAx")
      .withDefinedName("OtherRate", "0.5")
    val (result, _) = run(wb, "[" + changeTax + "]")
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(20))
    assertEquals(result.sheets(0)(ref"B1").effectiveValue, CellValue.Number(40))
    assertEquals(result.sheets(1)(ref"A1").effectiveValue, CellValue.Number(41))
    assertEquals(result.sheets(0)(ref"D1").effectiveValue, CellValue.Number(999))
    assertEquals(result.sheets(0)(ref"E1").effectiveValue, CellValue.Number(777))
  }

  test("global name changes preserve local shadows; qualified readers follow their lookup scope") {
    val wb = Workbook(
      Sheet("Data").put(ref"A1", formula("Tax*100", 10)),
      Sheet("Local")
        .put(ref"A1", formula("Tax*100", 999))
        .put(ref"B1", formula("dAtA!Tax*100", 10))
    ).withDefinedName("Tax", "0.1")
      .withDefinedName("Tax", "0.5", SheetName.unsafe("Local"))
      .fold(e => fail(e.message), identity)
    val (result, _) = run(wb, "[" + changeTax + "]")
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(20))
    assertEquals(result.sheets(1)(ref"A1").effectiveValue, CellValue.Number(999))
    assertEquals(result.sheets(1)(ref"B1").effectiveValue, CellValue.Number(20))
  }

  test("removing a local name refreshes readers against the global fallback") {
    val wb = Workbook(Sheet("Data").put(ref"A1", formula("Tax*100", 20)))
      .withDefinedName("Tax", "0.1")
      .withDefinedName("Tax", "0.2", SheetName.unsafe("Data"))
      .fold(e => fail(e.message), identity)
    val (result, _) = run(wb, """[{"op":"remove-name","name":"tAX","scope":"data"}]""")
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(10))
  }

  test("removing a global name withdraws cached readers and their dependents") {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"A1", formula("Tax*100", 10))
        .put(ref"B1", formula("A1*2", 20))
    )
      .withDefinedName("Tax", "0.1")
    val (result, summary) = run(wb, """[{"op":"remove-name","name":"Tax"}]""")
    for ref <- List(ref"A1", ref"B1") do
      result.sheets(0)(ref).value match
        case CellValue.Formula(_, cache, _) => assertEquals(cache, None, summary)
        case other => fail(s"expected formula at $ref, got $other")
  }

  test("adding a previously missing name refreshes an existing cached reader") {
    val wb = Workbook(Sheet("Data").put(ref"A1", formula("Tax*100", 999)))
    val (result, _) = run(wb, "[" + changeTax + "]")
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(20))
  }

  test("name changes in aggregate range slots refresh their readers") {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"A1", formula("SUM(Revenue)", 3))
        .put(ref"D1", 1)
        .put(ref"D2", 2)
        .put(ref"E1", 3)
        .put(ref"E2", 4)
    )
      .withDefinedName("Revenue", "Data!$D$1:$D$2")
    val (result, _) =
      run(wb, """[{"op":"define-name","name":"Revenue","refersTo":"Data!$E$1:$E$2"}]""")
    assertEquals(result.sheets(0)(ref"A1").effectiveValue, CellValue.Number(7))
  }

  test(
    "unsupported readers of a changed name lose their cache; unrelated unsupported readers keep theirs"
  ) {
    val wb = Workbook(
      Sheet("Data")
        .put(ref"A1", formula("NOSUCHFN(Tax)", 10))
        .put(ref"B1", formula("NOSUCHFN(D1)", 999))
        .put(ref"C1", formula("NOSUCHFN(\"Tax\")", 777))
        .put(ref"D1", 1)
    )
      .withDefinedName("Tax", "0.1")
    val (result, _) = run(wb, "[" + changeTax + "]")
    assertEquals(result.sheets(0)(ref"A1").value, CellValue.Formula("NOSUCHFN(Tax)", None))
    assertEquals(result.sheets(0)(ref"B1").effectiveValue, CellValue.Number(999))
    assertEquals(result.sheets(0)(ref"C1").effectiveValue, CellValue.Number(777))
  }
