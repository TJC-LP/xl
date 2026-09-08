// Gate for GH-590 (W2.9): `derives RowCodec` and the record entry points must resolve through the
// scripting prelude alone, from OUTSIDE com.tjclp.xl — the inline/export landmine is what this
// tests. A record declared at the top level of a script is the shape scala-cli users write.
package xlprelude

import java.time.LocalDate

import munit.FunSuite

import com.tjclp.xl.scripting.{*, given}

final case class Order(
  id: Int,
  customer: String,
  qty: Int,
  price: BigDecimal,
  shipped: Option[LocalDate]
) derives RowCodec

/** The pure import path must derive too (no prelude, no `.unsafe`). */
object RowCodecBaseImportProbe:
  import com.tjclp.xl.{*, given}

  final case class Line(sku: String, units: Long, note: Option[String]) derives RowCodec

  val codec: RowCodec[Line] = RowCodec[Line]
  val fields: Vector[String] = codec.fields
  val placed: XLResult[RowsPlaced] =
    Sheet("Lines").putRows(ref"A1", Vector(Line("A-1", 2L, None)))
  val back: Option[Either[RowCodecError, Vector[Line]]] =
    placed.toOption.flatMap(p => p.dataRange.map(p.sheet.readRows[Line]))

class RowCodecPreludeTest extends FunSuite:

  private val orders = Vector(
    Order(1, "Acme", 3, BigDecimal("9.99"), Some(LocalDate.of(2026, 1, 15))),
    Order(2, "Globex", 1, BigDecimal("120.00"), None)
  )

  test("GH-590: derives RowCodec resolves through the prelude; fields follow declaration order"):
    assertEquals(RowCodec[Order].fields, Vector("id", "customer", "qty", "price", "shipped"))
    assertEquals(RowCodec[Order].width, 5)

  test("GH-590: putRows / readRows round-trip through the prelude"):
    val placed = Sheet("Orders").putRows(ref"A2", orders).unsafe
    assertEquals(placed.dataRange.map(_.toA1), Some("A2:E3"))
    assertEquals(placed.headerRange, None)
    assertEquals(
      placed.dataRange.map(placed.sheet.readRows[Order]),
      Some(Right(orders)): Option[Either[RowCodecError, Vector[Order]]]
    )
    // Option fields: None is an empty cell, so nothing is stored there
    assertEquals(placed.sheet.cells.get(ref"E3"), None)

  test(
    "GH-590: putRowsWithHeader writes field names; headers/column/readRowsByHeader resolve"
  ):
    val placed = Sheet("Orders").putRowsWithHeader(ref"B1", orders).unsafe
    assertEquals(placed.headerRange.map(_.toA1), Some("B1:F1"))
    assertEquals(placed.dataRange.map(_.toA1), Some("B2:F3"))
    assertEquals(placed.range.map(_.toA1), Some("B1:F3"))
    val sheet = placed.sheet
    assertEquals(sheet.headers(Row.from1(1)).map(_._2), RowCodec[Order].fields)
    assertEquals(sheet.column("qty", Row.from1(1)), Some(Column.from0(3)))
    assertEquals(sheet.column("Qty", Row.from1(1)), Some(Column.from0(3)))
    assertEquals(sheet.column("missing", Row.from1(1)), None)
    assertEquals(
      sheet.readRowsByHeader[Order](Row.from1(1)),
      Right(orders): Either[RowCodecError, Vector[Order]]
    )

  test("GH-590: putTable registers an Excel table over header + records"):
    val placed = Sheet("Orders").putTable(ref"A1", orders, "Orders").unsafe
    val table = placed.sheet.getTable("Orders")
    assertEquals(table.map(_.range.toA1), Some("A1:E3"))
    assertEquals(table.map(_.columns.map(_.name)), Some(RowCodec[Order].fields))

  test("GH-590: decode errors name the row, column, field and cause; toXLError/message resolve"):
    val sheet = Sheet("Bad").putRows(ref"A1", orders).unsafe.sheet.put(ref"C2", "three")
    val result = sheet.readRows[Order](ref"A1:E2")
    result match
      case Left(RowCodecError.Field(row, column, field, cause)) =>
        assertEquals(row.index1, 2)
        assertEquals(column.toLetter, "C")
        assertEquals(field, "qty")
        assertEquals(cause, CodecError.TypeMismatch("Int", CellValue.Text("three")))
        val err: RowCodecError = RowCodecError.Field(row, column, field, cause)
        assert(err.message.contains("C2"))
        assert(err.toXLError.message.contains("qty"))
      case other => fail(s"expected a Field error, got $other")
