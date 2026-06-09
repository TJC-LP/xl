package com.tjclp.xl

import munit.FunSuite
import com.tjclp.xl.api.*
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.codec.syntax.*
import com.tjclp.xl.dsl.syntax.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.sheets.syntax.*

/** Workbook.upsert (total update-or-create) and readTypedOr/readTypedOpt (0.11.0). */
class UpsertSpec extends FunSuite:

  private val sales = SheetName.unsafe("Sales")
  private val absent = SheetName.unsafe("New")

  test("upsert updates an existing sheet in place"):
    val wb = Workbook(Sheet(sales).put(ref"A1" := 1))
    val updated = wb.upsert(sales, _.put(ref"B1" := 2))
    assertEquals(updated.sheets.size, 1)
    assertEquals(updated.sheets.headOption.map(_.cells.size), Some(2))

  test("upsert creates the sheet by applying f to a fresh empty sheet when absent"):
    val wb = Workbook(Sheet(sales))
    val updated = wb.upsert(absent, _.put(ref"A1" := "created"))
    assertEquals(updated.sheets.size, 2)
    assertEquals(
      updated.sheets.find(_.name == absent).map(_.cells.size),
      Some(1)
    )

  test("upsert with a literal string name is total and chains"):
    val wb = Workbook(Sheet(sales))
      .upsert("Summary", _.put(ref"A1" := "Total"))
      .upsert("Summary", _.put(ref"A2" := 100))
    assertEquals(wb.sheets.size, 2)
    assertEquals(
      wb.sheets.find(_.name.value == "Summary").map(_.cells.size),
      Some(2)
    )

  test("upsert with a runtime string returns XLResult and rejects invalid names"):
    val wb = Workbook(Sheet(sales))
    val bad = "Invalid:Name"
    wb.upsert(bad, identity) match
      case Left(err) => assert(err.message.nonEmpty)
      case Right(_) => fail("invalid runtime sheet name should be rejected")
    val good = "Runtime"
    assertEquals(wb.upsert(good, identity).map(_.sheets.size), Right(2))

  test("upsert preserves sheet order on update"):
    val wb = Workbook(Sheet(SheetName.unsafe("A")), Sheet(sales), Sheet(SheetName.unsafe("Z")))
    val updated = wb.upsert(sales, _.put(ref"A1" := 1))
    assertEquals(updated.sheets.map(_.name.value), Vector("A", "Sales", "Z"))

  // ========== readTypedOr / readTypedOpt ==========

  private val typedSheet = Sheet(sales)
    .put(ref"A1" := 42)
    .put(ref"B1" := "text")

  test("readTypedOr returns the value when present and decodable"):
    assertEquals(typedSheet.readTypedOr[Int](ref"A1", 0), 42)

  test("readTypedOr falls back on missing cell and on type mismatch"):
    assertEquals(typedSheet.readTypedOr[Int](ref"Z9", -1), -1)
    assertEquals(typedSheet.readTypedOr[Int](ref"B1", -1), -1)

  test("readTypedOpt flattens absence and mismatch to None"):
    assertEquals(typedSheet.readTypedOpt[Int](ref"A1"), Some(42))
    assertEquals(typedSheet.readTypedOpt[Int](ref"Z9"), None)
    assertEquals(typedSheet.readTypedOpt[Int](ref"B1"), None)
