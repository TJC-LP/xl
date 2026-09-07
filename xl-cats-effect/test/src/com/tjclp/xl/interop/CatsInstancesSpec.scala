package com.tjclp.xl.interop

import cats.syntax.all.*
import munit.FunSuite

import com.tjclp.xl.addressing.ARef
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.interop.CatsInstances.given
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.styles.font.Font
import com.tjclp.xl.styles.patch.StylePatch
import com.tjclp.xl.styles.units.StyleId

/**
 * The derived cats instances agree with xl's own Monoid, so cats syntax keeps working on patches.
 */
class CatsInstancesSpec extends FunSuite:

  private val ref = ARef.from1(1, 1)

  test("cats |+| on Patch is Patch.combine") {
    val put: Patch = Patch.Put(ref, CellValue.Text("x"))
    val style: Patch = Patch.SetStyle(ref, StyleId(3))

    assertEquals(put |+| style, Patch.combine(put, style))
    assertEquals(Patch.empty |+| put, Patch.combine(Patch.empty, put))
    assertEquals(cats.Monoid[Patch].empty, Patch.empty)
  }

  test("cats Monoid[StylePatch].combineAll folds like xl's Monoid") {
    val patches = Vector[StylePatch](
      StylePatch.SetFont(Font("Arial", 12.0)),
      StylePatch.SetFont(Font("Calibri", 14.0))
    )

    assertEquals(
      cats.Monoid[StylePatch].combineAll(patches),
      com.tjclp.xl.algebra.Monoid[StylePatch].combineAll(patches)
    )
    assertEquals(cats.Monoid[StylePatch].combineAll(Vector.empty), StylePatch.empty)
  }
