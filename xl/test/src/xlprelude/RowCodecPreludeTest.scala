// Gate for GH-590 (W2.9): `derives RowCodec` and the record entry points must resolve through the
// scripting prelude alone, from OUTSIDE com.tjclp.xl — the inline/export landmine is what this
// tests. A record declared at the top level of a script is the shape scala-cli users write.
// GH-614 adds the annotation hop: `@header` reaches a script through the exported alias, and the
// derivation macro must still see it on the constructor parameters.
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

/** GH-614: headers an identifier cannot spell, through the prelude's `header` alias. */
final case class Deal(
  @header("Portfolio Co.") portfolioCo: String,
  @header("Rev ($M)") rev: BigDecimal,
  ebitda: Option[BigDecimal]
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

  /** GH-614 through `com.tjclp.xl.{*, given}`: the annotation and the runtime twin. */
  final case class Tranche(@header("Size ($M)") size: BigDecimal, note: Option[String])
      derives RowCodec

  val headers: Vector[String] = RowCodec[Tranche].headers
  val renamed: XLResult[RowCodec[Tranche]] =
    RowCodec.derived[Tranche].withHeaders(Map("note" -> "Note / Comment"))

class RowCodecPreludeTest extends FunSuite:

  private val orders = Vector(
    Order(1, "Acme", 3, BigDecimal("9.99"), Some(LocalDate.of(2026, 1, 15))),
    Order(2, "Globex", 1, BigDecimal("120.00"), None)
  )

  private val deals = Vector(
    Deal("Acme", BigDecimal("12.5"), Some(BigDecimal("3.25"))),
    Deal("Globex", BigDecimal("40"), None)
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
    "GH-590: putRowsWithHeader writes field names; columnHeaders/columnOf/readRowsByHeader resolve"
  ):
    val placed = Sheet("Orders").putRowsWithHeader(ref"B1", orders).unsafe
    assertEquals(placed.headerRange.map(_.toA1), Some("B1:F1"))
    assertEquals(placed.dataRange.map(_.toA1), Some("B2:F3"))
    assertEquals(placed.range.map(_.toA1), Some("B1:F3"))
    val sheet = placed.sheet
    assertEquals(sheet.columnHeaders(Row.from1(1)).map(_._2), RowCodec[Order].fields)
    assertEquals(sheet.columnOf("qty", Row.from1(1)), Some(Column.from0(3)))
    assertEquals(sheet.columnOf("Qty", Row.from1(1)), Some(Column.from0(3)))
    assertEquals(sheet.columnOf("missing", Row.from1(1)), None)
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

  test("GH-614: @header through the prelude alias renames the header, not the field"):
    assertEquals(RowCodec[Deal].fields, Vector("portfolioCo", "rev", "ebitda"))
    assertEquals(RowCodec[Deal].headers, Vector("Portfolio Co.", "Rev ($M)", "ebitda"))
    assertEquals(RowCodec[Order].headers, RowCodec[Order].fields)
    val placed = Sheet("Deals").putRowsWithHeader(ref"A1", deals).unsafe
    assertEquals(placed.sheet.columnHeaders(Row.from1(1)).map(_._2), RowCodec[Deal].headers)
    assertEquals(
      placed.sheet.readRowsByHeader[Deal](Row.from1(1)),
      Right(deals): Either[RowCodecError, Vector[Deal]]
    )
    val table = Sheet("Deals").putTable(ref"A1", deals, "Deals").unsafe.sheet.getTable("Deals")
    assertEquals(table.map(_.columns.map(_.name)), Some(RowCodec[Deal].headers))

  test("GH-614: RowCodec.derived[Deal].withHeaders layers on the annotation; orExit unwraps it"):
    // The issue's header row, verbatim, as SheetJS wrote it
    val tracker = Sheet("Tracker")
      .put(ref"A1", "Rev ($M)")
      .put(ref"B1", "EBITDA ($M)")
      .put(ref"C1", "Portfolio Co.")
      .put(ref"A2", BigDecimal("12.5"))
      .put(ref"B2", BigDecimal("3.25"))
      .put(ref"C2", "Acme")
      .put(ref"A3", BigDecimal("40"))
      .put(ref"C3", "Globex")
    assertEquals(tracker.columnOf("rev", Row.from1(1)), None)
    assertEquals(tracker.columnOf("portfolioCo", Row.from1(1)), None)
    tracker.readRowsByHeader[Deal](Row.from1(1)) match
      case Left(RowCodecError.HeaderNotFound("ebitda", _, available)) =>
        assertEquals(available, Vector("Rev ($M)", "EBITDA ($M)", "Portfolio Co."))
      case other => fail(s"expected HeaderNotFound(ebitda), got $other")
    // Spelled with `derived`, never `RowCodec[Deal]` — the latter would summon the given being
    // defined; nested because a local given is in scope for its whole block
    locally {
      given RowCodec[Deal] =
        orExit(RowCodec.derived[Deal].withHeaders(Map("ebitda" -> "EBITDA ($M)")))
      assertEquals(RowCodec[Deal].headers, Vector("Portfolio Co.", "Rev ($M)", "EBITDA ($M)"))
      assertEquals(
        tracker.readRowsByHeader[Deal](Row.from1(1)),
        Right(deals): Either[RowCodecError, Vector[Deal]]
      )
    }
    val refused: XLResult[RowCodec[Deal]] =
      RowCodec.derived[Deal].withHeaders(Map("nope" -> "x"))
    assertEquals(refused.left.map(_.code), Left("INVALID_ARGUMENT"))

  test("GH-614: the pure import path sees the annotation and the runtime twin too"):
    assertEquals(RowCodecBaseImportProbe.headers, Vector("Size ($M)", "note"))
    assertEquals(
      RowCodecBaseImportProbe.renamed.map(_.headers),
      Right(Vector("Size ($M)", "Note / Comment")): XLResult[Vector[String]]
    )
