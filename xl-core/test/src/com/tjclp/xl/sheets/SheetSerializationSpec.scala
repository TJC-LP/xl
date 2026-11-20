package com.tjclp.xl.sheets

import java.io.{ByteArrayInputStream, ByteArrayOutputStream, ObjectInputStream, ObjectOutputStream}

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.{Cell, CellError, CellValue}
import com.tjclp.xl.styles.units.StyleId
import munit.ScalaCheckSuite
import org.scalacheck.Prop
import org.scalacheck.Prop.*

/**
 * #891: a [[Sheet]] is Java-serializable, as it was while its cells were a `Map`. The store writes
 * its cells and is rebuilt through its builder on the way in, so what comes back is the canonical
 * store: equal to the one written, with every structural invariant of a built one.
 *
 * No promise across versions: the stream holds `Cell`s, and an `ARef`'s packing is this version's.
 */
class SheetSerializationSpec extends ScalaCheckSuite:
  import SheetSerializationSpec.*

  test("an empty sheet round-trips") {
    val sheet = Sheet(SheetName.unsafe("S"))
    val back = roundTrip(sheet)
    assertEquals(back, sheet)
    assertEquals(back.cells.checkInvariants, Right(()))
    assert(back.cells.isEmpty)
    // the canonical empty store, not a second one
    assert(back.cells eq CellStore.empty)
  }

  test("a populated sheet round-trips: styles, hyperlinks, a wide row, off-grid references") {
    val sheet = Sheet(SheetName.unsafe("Data")).copy(cells = CellStore.from(populated))
    val back = roundTrip(sheet)
    assertEquals(back, sheet)
    assertEquals(back.cells, sheet.cells)
    assertEquals(back.cells.hashCode, sheet.cells.hashCode)
    assertEquals(back.cells.checkInvariants, Right(()))
    assertEquals(back.cells.size, sheet.cells.size)
    assertEquals(back.cells.valuedCount, sheet.cells.valuedCount)
    assertEquals(back.cells.usedRange, sheet.cells.usedRange)
    assertEquals(back.cells.extent, sheet.cells.extent)
    assertEquals(back.cells.debugHeight, sheet.cells.debugHeight)
    assertEquals(back.cells.iterator.toList, sheet.cells.iterator.toList)
    // a style-only cell is still the Empty the store's tags and `match`es know
    back.cells.get(styleOnly).map(_.value) match
      case Some(CellValue.Empty) => ()
      case other => fail(s"expected Empty at the style-only cell, got $other")
    assertEquals(back.cells.styleAt(styleOnly), Some(StyleId(3)))
    assertEquals(back.cells.hyperlinkAt(linked), sheet.cells.hyperlinkAt(linked))
    assert(back.cells.hyperlinkAt(linked).isDefined)
    // and it is a working store: an edit after the round trip keeps the invariants
    val edited = back.cells.updated(Cell(ARef.from0(300, 5), CellValue.Text("after")))
    assertEquals(edited.checkInvariants, Right(()))
    assertEquals(edited, sheet.cells.updated(Cell(ARef.from0(300, 5), CellValue.Text("after"))))
  }

  test("a store round-trips on its own") {
    val store = CellStore.from(populated)
    val back = roundTrip(store)
    assertEquals(back, store)
    assertEquals(back.checkInvariants, Right(()))
  }

  property("any store round-trips to an equal store with the invariants of a built one") {
    Prop.forAll(CellStoreLawSpec.genStart) { cells =>
      val store = CellStore.from(cells)
      val back = roundTrip(store)
      (back == store) :| "equal" &&
      (back.hashCode == store.hashCode) :| "hash" &&
      (back.checkInvariants == Right(())) :| s"invariants: ${back.checkInvariants}" &&
      (back.debugHeight == store.debugHeight) :| "height" &&
      (back.iterator.toList == store.iterator.toList) :| "cells in order" &&
      (back.usedRange == store.usedRange && back.extent == store.extent) :| "bounds"
    }
  }

object SheetSerializationSpec:

  private val styleOnly = ARef.from0(2, 40)
  private val linked = ARef.from0(1, 0)

  /** Every row shape: a narrow block, a wide row, a banded row, style-only and off-grid cells. */
  private val populated: List[Cell] =
    val narrow = List.tabulate(200)(i => Cell(ARef.from0(i % 4, i / 4), number(i)))
    val wide = List.tabulate(70)(i => Cell(ARef.from0(i, 60), number(i), Some(StyleId(i % 5))))
    val banded = List.tabulate(600)(i => Cell(ARef.from0(i * 3, 61), CellValue.Text(s"t$i")))
    val mixed = List(
      Cell(linked, CellValue.Text("link")).withHyperlink("https://example.com/x"),
      Cell(styleOnly, CellValue.Empty, Some(StyleId(3))),
      Cell(ARef.from0(5, 41), CellValue.Bool(true)),
      Cell(ARef.from0(6, 41), CellValue.Error(CellError.Div0)),
      Cell(ARef.from0(7, 41), CellValue.Formula("A1*2", Some(number(4)))),
      Cell(ARef.from0(8, 41), CellValue.Formula("SUM(A1:A2)", None)),
      // off the grid: past the last row and column, and the negative ones only a shift can mint
      Cell(ARef.from0(16384, 1048576), number(1)),
      Cell(ARef.from0(-1, 3), number(2)),
      Cell(ARef.from0(3, -1), CellValue.Text("negative row")),
      Cell(ARef.from0(Int.MinValue, Int.MinValue), number(3), Some(StyleId(Int.MaxValue))),
      Cell(ARef.from0(Int.MaxValue, Int.MaxValue), number(4))
    )
    narrow ++ wide ++ banded ++ mixed

  private def number(i: Int): CellValue = CellValue.Number(BigDecimal(i))

  // A cast is the only way out of `readObject`; the test fails on a wrong class.
  @SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
  private def roundTrip[A <: AnyRef](a: A): A =
    val bytes = new ByteArrayOutputStream
    val out = new ObjectOutputStream(bytes)
    out.writeObject(a)
    out.close()
    val in = new ObjectInputStream(new ByteArrayInputStream(bytes.toByteArray))
    try in.readObject().asInstanceOf[A]
    finally in.close()
