package com.tjclp.xl.codec

import java.time.{LocalDate, LocalDateTime}

import munit.ScalaCheckSuite
import org.scalacheck.{Arbitrary, Gen}
import org.scalacheck.Prop.*

import com.tjclp.xl.api.*
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.codec.CellCodec.given
import com.tjclp.xl.codec.rowSyntax.*
import com.tjclp.xl.macros.ref
import com.tjclp.xl.richtext.RichText
import com.tjclp.xl.sheets.styleSyntax.*
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.numfmt.NumFmt

/** GH-590 (W2.9): records as rows — derived `RowCodec`, the Sheet entry points, and their laws. */
@SuppressWarnings(Array("org.wartremover.warts.OptionPartial"))
class RowCodecSpec extends ScalaCheckSuite:

  final case class Order(
    id: Int,
    customer: String,
    qty: Int,
    price: BigDecimal,
    note: Option[String]
  ) derives RowCodec

  /** Every primitive codec type, required and optional, in one record. */
  final case class Everything(
    s: String,
    i: Int,
    l: Long,
    d: Double,
    bd: BigDecimal,
    b: Boolean,
    date: LocalDate,
    dateTime: LocalDateTime,
    rich: RichText,
    os: Option[String],
    oi: Option[Int],
    ol: Option[Long],
    od: Option[Double],
    obd: Option[BigDecimal],
    ob: Option[Boolean],
    odate: Option[LocalDate],
    odateTime: Option[LocalDateTime],
    orich: Option[RichText]
  ) derives RowCodec

  /**
   * 23 fields: past `Tuple22`, so `Tuple.fromArray` yields a `TupleXXL` and `fromProduct` sees it.
   */
  final case class Wide(
    f01: Int,
    f02: String,
    f03: Long,
    f04: Double,
    f05: BigDecimal,
    f06: Boolean,
    f07: LocalDate,
    f08: Option[Int],
    f09: Option[String],
    f10: Int,
    f11: Int,
    f12: Int,
    f13: Int,
    f14: Int,
    f15: Int,
    f16: Int,
    f17: Int,
    f18: Int,
    f19: Int,
    f20: Int,
    f21: Int,
    f22: Int,
    f23: Option[Long]
  ) derives RowCodec

  private val row1 = Row.from1(1)

  private val genOrder: Gen[Order] =
    for
      id <- Gen.choose(1, 100000)
      customer <- Gen.alphaNumStr
      qty <- Gen.choose(-1000, 1000)
      price <- Gen.choose(0L, 10_000_000L).map(cents => BigDecimal(cents) / 100)
      note <- Gen.option(Gen.alphaNumStr)
    yield Order(id, customer, qty, price, note)

  private val genLocalDate: Gen[LocalDate] =
    Gen.choose(0L, 60000L).map(LocalDate.of(1905, 1, 1).plusDays)

  private val genLocalDateTime: Gen[LocalDateTime] =
    for
      date <- genLocalDate
      seconds <- Gen.choose(0, 86399)
    yield date.atStartOfDay.plusSeconds(seconds.toLong)

  private val genRichText: Gen[RichText] =
    Gen.alphaNumStr.map(RichText.plain)

  private val genEverything: Gen[Everything] =
    for
      s <- Gen.alphaNumStr
      i <- Arbitrary.arbitrary[Int]
      l <- Arbitrary.arbitrary[Long]
      d <- Gen.choose(-1.0e12, 1.0e12)
      bd <- Gen.choose(-1_000_000L, 1_000_000L).map(n => BigDecimal(n) / 1000)
      b <- Gen.oneOf(true, false)
      date <- genLocalDate
      dateTime <- genLocalDateTime
      rich <- genRichText
      os <- Gen.option(Gen.alphaNumStr)
      oi <- Gen.option(Arbitrary.arbitrary[Int])
      ol <- Gen.option(Arbitrary.arbitrary[Long])
      od <- Gen.option(Gen.choose(-1.0e12, 1.0e12))
      obd <- Gen.option(Gen.choose(-1_000_000L, 1_000_000L).map(n => BigDecimal(n) / 1000))
      ob <- Gen.option(Gen.oneOf(true, false))
      odate <- Gen.option(genLocalDate)
      odateTime <- Gen.option(genLocalDateTime)
      orich <- Gen.option(genRichText)
    yield Everything(
      s,
      i,
      l,
      d,
      bd,
      b,
      date,
      dateTime,
      rich,
      os,
      oi,
      ol,
      od,
      obd,
      ob,
      odate,
      odateTime,
      orich
    )

  /** Anchors far from the sheet edges so a generated block always fits. */
  private val genAnchor: Gen[ARef] =
    for
      col <- Gen.choose(0, 200)
      row <- Gen.choose(0, 5000)
    yield ARef.from0(col, row)

  private def right[A](result: Either[?, A]): A = result match
    case Right(a) => a
    case Left(err) => fail(s"expected Right, got Left($err)")

  // ========== Derivation ==========

  test("derived: fields are the case-class field names in declaration order; width matches") {
    assertEquals(RowCodec[Order].fields, Vector("id", "customer", "qty", "price", "note"))
    assertEquals(RowCodec[Order].width, 5)
    assertEquals(RowCodec[Everything].width, 18)
  }

  test("derived: write encodes each field through its CellCodec, None as an empty cell") {
    val written = RowCodec[Order].write(Order(7, "Acme", 2, BigDecimal("1.50"), None))
    assertEquals(
      written.map(_._1),
      Vector(
        CellValue.Number(BigDecimal(7)),
        CellValue.Text("Acme"),
        CellValue.Number(BigDecimal(2)),
        CellValue.Number(BigDecimal("1.50")),
        CellValue.Empty
      )
    )
    assertEquals(written(3)._2.map(_.numFmt), Some(NumFmt.Decimal))
    assertEquals(written(4)._2, None)
  }

  test("derived: read with the wrong number of cells is a Width error, not an exception") {
    val cells = Vector(Cell(ref"A1", CellValue.Number(BigDecimal(1))))
    assertEquals(RowCodec[Order].read(cells), Left(RowCodecError.Width(5, 1)))
  }

  test("derived: a record without fields does not compile") {
    val errors = compileErrors("final case class Unit0() derives RowCodec")
    assert(errors.contains("at least one field"), errors)
  }

  test("derived: a field type without a CellCodec does not compile") {
    val errors = compileErrors("final case class Bad(u: java.util.UUID) derives RowCodec")
    assert(errors.nonEmpty)
    assert(errors.contains("UUID"), errors)
  }

  // ========== Round-trip laws ==========

  property("law: readRows(putRows(at, rows).dataRange) == Right(rows) for Order") {
    forAll(genAnchor, Gen.nonEmptyListOf(genOrder)) { (at, rowsList) =>
      val rows = rowsList.toVector
      val placed = right(Sheet("Law").putRows(at, rows))
      val expectedRange =
        CellRange(at, ARef(at.col + RowCodec[Order].width - 1, at.row + rows.size - 1))
      (placed.dataRange ?= Some(expectedRange)) &&
      (placed.headerRange ?= None) &&
      (placed.dataRange.map(placed.sheet.readRows[Order]) ?= Some(Right(rows)))
    }
  }

  property("law: every primitive codec type round-trips, required and optional") {
    forAll(Gen.nonEmptyListOf(genEverything)) { rowsList =>
      val rows = rowsList.toVector
      val placed = right(Sheet("Law").putRows(ref"A1", rows))
      placed.dataRange.map(placed.sheet.readRows[Everything]) ?= Some(Right(rows))
    }
  }

  property("law: readRowsByHeader(putRowsWithHeader(at, rows).headerRow) == Right(rows)") {
    forAll(genAnchor, Gen.nonEmptyListOf(genOrder)) { (at, rowsList) =>
      val rows = rowsList.toVector
      val placed = right(Sheet("Law").putRowsWithHeader(at, rows))
      val width = RowCodec[Order].width
      (placed.headerRange ?= Some(CellRange(at, ARef(at.col + width - 1, at.row)))) &&
      (placed.dataRange ?= Some(
        CellRange(ARef(at.col, at.row + 1), ARef(at.col + width - 1, at.row + rows.size))
      )) &&
      (placed.range ?= Some(CellRange(at, ARef(at.col + width - 1, at.row + rows.size)))) &&
      (placed.sheet.readRowsByHeader[Order](at.row) ?= Right(rows))
    }
  }

  test("law: a 23-field record (past Tuple22) round-trips by position and by header") {
    val wide = Wide(
      1,
      "two",
      3L,
      4.5,
      BigDecimal("6.25"),
      true,
      LocalDate.of(2026, 9, 8),
      None,
      Some("nine"),
      10,
      11,
      12,
      13,
      14,
      15,
      16,
      17,
      18,
      19,
      20,
      21,
      22,
      Some(23L)
    )
    assertEquals(RowCodec[Wide].width, 23)
    assertEquals(RowCodec[Wide].fields.take(2), Vector("f01", "f02"))
    val rows = Vector(wide, wide.copy(f01 = 2, f08 = Some(8), f23 = None))
    val placed = right(Sheet("Wide").putRowsWithHeader(ref"B2", rows))
    assertEquals(placed.dataRange.map(_.toA1), Some("B3:X4"))
    assertEquals(placed.sheet.readRowsByHeader[Wide](Row.from1(2)), Right(rows))
    assertEquals(placed.dataRange.map(placed.sheet.readRows[Wide]), Some(Right(rows)))
  }

  test("putRows with no records writes nothing and reports no data range") {
    val sheet = Sheet("Empty").put(ref"A1", "keep")
    val placed = right(sheet.putRows(ref"A2", Vector.empty[Order]))
    assertEquals(placed.sheet, sheet)
    assertEquals(placed.dataRange, None)
    assertEquals(placed.range, None)
    assertEquals(placed.count, 0)
  }

  test("putRowsWithHeader with no records writes only the header") {
    val placed = right(Sheet("Empty").putRowsWithHeader(ref"A1", Vector.empty[Order]))
    assertEquals(placed.headerRange.map(_.toA1), Some("A1:E1"))
    assertEquals(placed.dataRange, None)
    assertEquals(placed.range.map(_.toA1), Some("A1:E1"))
    assertEquals(placed.sheet.cells.size, 5)
    assertEquals(placed.sheet.readRowsByHeader[Order](row1), Right(Vector.empty))
  }

  // ========== Writing semantics ==========

  test(
    "putRows: codec style hints register (Decimal for BigDecimal) and merge into existing styles"
  ) {
    val bold = CellStyle.default.withFont(CellStyle.default.font.withBold(true))
    val sheet = Sheet("Styled").withCellStyle(ref"D1", bold)
    val placed = right(sheet.putRows(ref"A1", Vector(Order(1, "a", 1, BigDecimal("2.5"), None))))
    val style =
      placed.sheet.cells.get(ref"D1").flatMap(_.styleId).flatMap(placed.sheet.styleRegistry.get)
    assertEquals(style.map(_.numFmt), Some(NumFmt.Decimal))
    assertEquals(style.map(_.font.bold), Some(true))
  }

  test("putRows: a codec NumFmt hint over a Currency-formatted cell keeps Currency, like put") {
    val currency = CellStyle.default.withNumFmt(NumFmt.Currency)
    val sheet = Sheet("Cur").withCellStyle(ref"D1", currency)
    def styleAt(s: Sheet) = s.cells.get(ref"D1").flatMap(_.styleId).flatMap(s.styleRegistry.get)
    val viaRows =
      right(sheet.putRows(ref"A1", Vector(Order(1, "a", 1, BigDecimal("2.5"), None)))).sheet
    val viaPut = sheet.put(ref"D1", BigDecimal("2.5"))
    assertEquals(styleAt(viaRows).map(_.numFmt), Some(NumFmt.Currency))
    assertEquals(styleAt(viaRows), styleAt(viaPut))
    assertEquals(viaRows.cells.get(ref"D1"), viaPut.cells.get(ref"D1"))
    // The same policy fills a General format in: bold-only D1 gains Decimal both ways
    val bold = CellStyle.default.withFont(CellStyle.default.font.withBold(true))
    val boldSheet = Sheet("Bold").withCellStyle(ref"D1", bold)
    val boldRows =
      right(boldSheet.putRows(ref"A1", Vector(Order(1, "a", 1, BigDecimal("2.5"), None)))).sheet
    assertEquals(styleAt(boldRows), styleAt(boldSheet.put(ref"D1", BigDecimal("2.5"))))
    assertEquals(styleAt(boldRows).map(s => (s.numFmt, s.font.bold)), Some((NumFmt.Decimal, true)))
  }

  test("putRows: only the records' cells are written — rows of a longer earlier block survive") {
    val three = Vector(
      Order(1, "a", 1, BigDecimal(1), None),
      Order(2, "b", 2, BigDecimal(2), None),
      Order(3, "c", 3, BigDecimal(3), None)
    )
    val first = right(Sheet("Regen").putRows(ref"A1", three)).sheet
    val shorter = right(first.putRows(ref"A1", three.take(1).map(_.copy(customer = "z")))).sheet
    assertEquals(
      shorter.readRows[Order](ref"A1:E3").map(_.map(_.customer)),
      Right(Vector("z", "b", "c")): Either[RowCodecError, Vector[String]]
    )
    // Regenerating in place means clearing the old block first
    val cleared = ref"A1:E3".cells.foldLeft(first)(_.remove(_))
    val regenerated = right(cleared.putRowsWithHeader(ref"A1", three.take(1)))
    assertEquals(regenerated.sheet.readRowsByHeader[Order](row1), Right(three.take(1)))
    assertEquals(regenerated.sheet.cells.size, 5 + 4)
  }

  test("putRows: a None field clears an existing cell's value but never creates a cell") {
    val sheet = Sheet("Clear").put(ref"E1", "stale")
    val placed = right(sheet.putRows(ref"A1", Vector(Order(1, "a", 1, BigDecimal(1), None))))
    assertEquals(placed.sheet.cells.get(ref"E1").map(_.value), Some(CellValue.Empty))
    val fresh =
      right(Sheet("Fresh").putRows(ref"A1", Vector(Order(1, "a", 1, BigDecimal(1), None))))
    assertEquals(fresh.sheet.cells.contains(ref"E1"), false)
    assertEquals(fresh.sheet.cells.size, 4)
  }

  test("putRows: a block that runs past column XFD or row 1048576 is OutOfBounds") {
    val order = Order(1, "a", 1, BigDecimal(1), None)
    val lastCol = ARef(Column.from0(Column.MaxIndex0 - 2), row1)
    Sheet("Edge").putRows(lastCol, Vector(order)) match
      case Left(XLError.OutOfBounds(_, _)) => ()
      case other => fail(s"expected OutOfBounds, got $other")
    val lastRow = ARef(Column.from0(0), Row.from0(Row.MaxIndex0))
    Sheet("Edge").putRows(lastRow, Vector(order, order)) match
      case Left(XLError.OutOfBounds(_, _)) => ()
      case other => fail(s"expected OutOfBounds, got $other")
    // The header counts as a row too
    Sheet("Edge").putRowsWithHeader(lastRow, Vector(order)) match
      case Left(XLError.OutOfBounds(_, _)) => ()
      case other => fail(s"expected OutOfBounds, got $other")
    // Exactly fitting is fine
    assert(Sheet("Edge").putRows(lastRow, Vector(order)).isRight)
  }

  test("putRows: a hand-written codec whose write width drifts from its fields is rejected") {
    given RowCodec[Int] = new RowCodec[Int]:
      def fields: Vector[String] = Vector("a", "b")
      def read(cells: Vector[Cell]): Either[RowCodecError, Int] = Right(0)
      def write(a: Int): Vector[(CellValue, Option[CellStyle])] =
        Vector((CellValue.Number(BigDecimal(a)), None))
    Sheet("Drift").putRows(ref"A1", Vector(1)) match
      case Left(XLError.ValueCountMismatch(2, 1, _)) => ()
      case other => fail(s"expected ValueCountMismatch(2, 1), got $other")
  }

  test("putRows: a width drift on a later record is all-or-nothing and names the record") {
    // Record n writes n cells: the first fits the two fields, the second does not
    given RowCodec[Int] = new RowCodec[Int]:
      def fields: Vector[String] = Vector("a", "b")
      def read(cells: Vector[Cell]): Either[RowCodecError, Int] = Right(0)
      def write(a: Int): Vector[(CellValue, Option[CellStyle])] =
        Vector.fill(a)((CellValue.Number(BigDecimal(a)), None))
    val sheet = Sheet("Drift").put(ref"A1", "keep")
    sheet.putRows(ref"A2", Vector(2, 1, 2)) match
      case Left(XLError.ValueCountMismatch(2, 1, context)) =>
        assert(context.contains("record 1"), context)
      case other => fail(s"expected ValueCountMismatch(2, 1), got $other")
    assertEquals(sheet.cells.size, 1)
    // With every record the right width the same rows go through
    assertEquals(right(sheet.putRows(ref"A2", Vector(2, 2))).dataRange.map(_.toA1), Some("A2:B3"))
  }

  test("putRows: a codec without fields cannot place anything") {
    given RowCodec[Int] = new RowCodec[Int]:
      def fields: Vector[String] = Vector.empty
      def read(cells: Vector[Cell]): Either[RowCodecError, Int] = Right(0)
      def write(a: Int): Vector[(CellValue, Option[CellStyle])] = Vector.empty
    assert(Sheet("Zero").putRows(ref"A1", Vector(1)).isLeft)
  }

  // ========== Reading semantics ==========

  test("readRows: reads through a formula's cached value (GH-477); uncached is a Field error") {
    val base =
      right(Sheet("Fx").putRows(ref"A1", Vector(Order(1, "a", 1, BigDecimal(1), None)))).sheet
    val cached = base.put(ref"C1", CellValue.Formula("1+1", Some(CellValue.Number(BigDecimal(2)))))
    assertEquals(
      cached.readRows[Order](ref"A1:E1").map(_.map(_.qty)),
      Right(Vector(2)): Either[RowCodecError, Vector[Int]]
    )
    val uncached = base.put(ref"C1", CellValue.Formula("1+1", None))
    uncached.readRows[Order](ref"A1:E1") match
      case Left(RowCodecError.Field(row, col, "qty", CodecError.TypeMismatch("Int", _))) =>
        assertEquals((col.toLetter, row.index1), ("C", 1))
      case other => fail(s"expected Field(qty, TypeMismatch), got $other")
  }

  test("readRows: a required field on an empty cell is Missing(row, column, field)") {
    val sheet = Sheet("Miss").put(ref"A3", 1).put(ref"C3", 2).put(ref"D3", BigDecimal(1))
    assertEquals(
      sheet.readRows[Order](ref"A3:E3"),
      Left(RowCodecError.Missing(Row.from1(3), Column.from0(1), "customer"))
    )
  }

  test("readRows: the first failing cell in row-major order is reported") {
    val sheet = right(
      Sheet("First").putRows(
        ref"A1",
        Vector(Order(1, "a", 1, BigDecimal(1), None), Order(2, "b", 2, BigDecimal(2), None))
      )
    ).sheet.put(ref"D1", "bad-price").put(ref"A2", "bad-id")
    sheet.readRows[Order](ref"A1:E2") match
      case Left(RowCodecError.Field(row, col, "price", _)) =>
        assertEquals((col.toLetter, row.index1), ("D", 1))
      case other => fail(s"expected the D1 price error first, got $other")
  }

  test("readRows: a range whose width differs from the record is a Width error") {
    assertEquals(
      Sheet("W").readRows[Order](ref"A1:C4"),
      Left(RowCodecError.Width(5, 3)): Either[RowCodecError, Vector[Order]]
    )
    assertEquals(
      Sheet("W").readRows[Order](ref"A1:F4"),
      Left(RowCodecError.Width(5, 6)): Either[RowCodecError, Vector[Order]]
    )
  }

  test("readRows: an error cell under an Option field is a Field error, not None") {
    final case class Sparse(a: Option[Int], b: Option[String]) derives RowCodec
    val sheet = Sheet("Err").put(ref"A1", CellValue.Error(CellError.NA))
    assertEquals(
      sheet.readRows[Sparse](ref"A1:B1"),
      Left(
        RowCodecError.Field(
          row1,
          Column.from0(0),
          "a",
          CodecError.TypeMismatch("Int", CellValue.Error(CellError.NA))
        )
      ): Either[RowCodecError, Vector[Sparse]]
    )
  }

  test("readRows: an all-optional record decodes blank rows as all-None records") {
    final case class Sparse(a: Option[Int], b: Option[String]) derives RowCodec
    assertEquals(
      Sheet("Sparse").readRows[Sparse](ref"A1:B3"),
      Right(Vector.fill(3)(Sparse(None, None))): Either[RowCodecError, Vector[Sparse]]
    )
  }

  // ========== Headers ==========

  test(
    "columnHeaders: verbatim non-blank header text left to right; numbers and rich text included"
  ) {
    val sheet = Sheet("H")
      .put(ref"B1", " Qty ")
      .put(ref"A1", "Id")
      .put(ref"D1", 2024)
      .put(ref"E1", RichText.plain("Rich"))
      .put(ref"F1", "")
      .put(ref"G1", "   ")
    assertEquals(
      sheet.columnHeaders(row1),
      Vector(
        Column.from0(0) -> "Id",
        Column.from0(1) -> " Qty ",
        Column.from0(3) -> "2024",
        Column.from0(4) -> "Rich"
      )
    )
    assertEquals(sheet.columnHeaders(Row.from1(2)), Vector.empty)
  }

  test("columnOf: exact match wins, then case/space/underscore-insensitive; leftmost on ties") {
    val sheet = Sheet("C")
      .put(ref"A1", "Order ID")
      .put(ref"B1", "order_id")
      .put(ref"C1", "orderId")
      .put(ref"D1", "Unit Price")
    assertEquals(sheet.columnOf("orderId", row1), Some(Column.from0(2)))
    assertEquals(sheet.columnOf("OrderID", row1), Some(Column.from0(0)))
    assertEquals(sheet.columnOf("unitPrice", row1), Some(Column.from0(3)))
    assertEquals(sheet.columnOf("unit_price", row1), Some(Column.from0(3)))
    assertEquals(sheet.columnOf("total", row1), None)
  }

  test("readRowsByHeader: matches fields to headers in any column order, ignores extra columns") {
    val sheet = Sheet("ByHeader")
      .put(ref"A1", "Total")
      .put(ref"B1", "Qty")
      .put(ref"C1", "Customer")
      .put(ref"D1", "Note")
      .put(ref"E1", "Price")
      .put(ref"F1", "Id")
      .put(ref"A2", 99)
      .put(ref"B2", 3)
      .put(ref"C2", "Acme")
      .put(ref"E2", BigDecimal("9.99"))
      .put(ref"F2", 1)
      .put(ref"B3", 1)
      .put(ref"C3", "Globex")
      .put(ref"D3", "rush")
      .put(ref"E3", BigDecimal("120"))
      .put(ref"F3", 2)
    assertEquals(
      sheet.readRowsByHeader[Order](row1),
      Right(
        Vector(
          Order(1, "Acme", 3, BigDecimal("9.99"), None),
          Order(2, "Globex", 1, BigDecimal("120"), Some("rush"))
        )
      ): Either[RowCodecError, Vector[Order]]
    )
  }

  test(
    "readRowsByHeader: stops at the first blank row (Excel's current region), ignores what follows"
  ) {
    val placed = right(
      Sheet("Region").putRowsWithHeader(
        ref"A1",
        Vector(Order(1, "a", 1, BigDecimal(1), None), Order(2, "b", 2, BigDecimal(2), None))
      )
    )
    val sheet = placed.sheet.put(ref"A5", "Total").put(ref"D5", BigDecimal(3))
    assertEquals(
      sheet.readRowsByHeader[Order](row1).map(_.map(_.id)),
      Right(Vector(1, 2)): Either[RowCodecError, Vector[Int]]
    )
  }

  test("readRowsByHeader: a missing header is HeaderNotFound with the headers that exist") {
    val sheet = Sheet("NoHeader").put(ref"A1", "id").put(ref"B1", "customer").put(ref"C1", "qty")
    assertEquals(
      sheet.readRowsByHeader[Order](row1),
      Left(RowCodecError.HeaderNotFound("price", row1, Vector("id", "customer", "qty")))
    )
  }

  test("readRowsByHeader: a header row with no data below decodes to no records") {
    val sheet = right(Sheet("Only").putRowsWithHeader(ref"C4", Vector.empty[Order])).sheet
    assertEquals(sheet.readRowsByHeader[Order](Row.from1(4)), Right(Vector.empty))
  }

  // ========== Tables ==========

  test("putTable: header + records + a TableSpec named and columned after the record") {
    val rows = Vector(Order(1, "a", 1, BigDecimal(1), None), Order(2, "b", 2, BigDecimal(2), None))
    val placed = right(Sheet("T").putTable(ref"B2", rows, "Orders"))
    assertEquals(placed.headerRange.map(_.toA1), Some("B2:F2"))
    assertEquals(placed.dataRange.map(_.toA1), Some("B3:F4"))
    val table = placed.sheet.getTable("Orders").get
    assertEquals(table.range.toA1, "B2:F4")
    assertEquals(table.displayName, "Orders")
    assertEquals(table.columns.map(_.name), RowCodec[Order].fields)
    assertEquals(table.autoFilter, None)
    assertEquals(placed.sheet.readRows[Order](table.dataRange), Right(rows))
  }

  test("putTable: no records keeps Excel's one blank data row under the header") {
    val placed = right(Sheet("T").putTable(ref"A1", Vector.empty[Order], "Empty"))
    assertEquals(placed.dataRange, None)
    assertEquals(placed.sheet.getTable("Empty").map(_.range.toA1), Some("A1:E2"))
  }

  test("putTable: invalid and duplicate table names are rejected, nothing is written") {
    val rows = Vector(Order(1, "a", 1, BigDecimal(1), None))
    Sheet("T").putTable(ref"A1", rows, "bad name") match
      case Left(XLError.InvalidTableName(_, _)) => ()
      case other => fail(s"expected InvalidTableName, got $other")
    val once = right(Sheet("T").putTable(ref"A1", rows, "Orders")).sheet
    once.putTable(ref"H1", rows, "Orders") match
      case Left(XLError.InvalidTableName("Orders", _)) => ()
      case other => fail(s"expected a duplicate-name InvalidTableName, got $other")
  }

  // ========== Errors ==========

  test("RowCodecError: message and toXLError carry the cell, the field and the cause") {
    val at = ref"C7"
    val field = RowCodecError.Field(
      at.row,
      at.col,
      "qty",
      CodecError.TypeMismatch("Int", CellValue.Text("three"))
    )
    assert(field.message.contains("C7"), field.message)
    assert(field.message.contains("qty"), field.message)
    field.toXLError match
      case XLError.TypeMismatch(expected, actual, "C7") =>
        assert(expected.contains("qty"))
        assert(actual.contains("three"))
      case other => fail(s"expected TypeMismatch at C7, got $other")

    val parse = RowCodecError.Field(
      at.row,
      at.col,
      "id",
      CodecError.ParseError("1.5", "Int", "decimal places")
    )
    parse.toXLError match
      case XLError.ParseError("C7", reason) => assert(reason.contains("id"), reason)
      case other => fail(s"expected ParseError at C7, got $other")

    val missing = RowCodecError.Missing(at.row, at.col, "customer")
    assert(missing.message.contains("customer"))
    missing.toXLError match
      case XLError.TypeMismatch(expected, _, "C7") => assert(expected.contains("customer"))
      case other => fail(s"expected TypeMismatch at C7, got $other")

    val header = RowCodecError.HeaderNotFound("price", row1, Vector("id", "qty"))
    assert(header.message.contains("price"))
    assert(header.message.contains("qty"))
    header.toXLError match
      case XLError.InvalidReference(reason) => assert(reason.contains("price"))
      case other => fail(s"expected InvalidReference, got $other")

    assertEquals(
      RowCodecError.Width(5, 3).toXLError,
      XLError.ValueCountMismatch(5, 3, "record columns")
    )
  }
