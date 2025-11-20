package com.tjclp.xl

import munit.FunSuite
import com.tjclp.xl.context.ModificationTracker

class ModificationTrackerSpec extends FunSuite:

  test("clean tracker reports no modifications") {
    assert(ModificationTracker.clean.isClean)
  }

  test("markSheet marks the provided index") {
    val tracker = ModificationTracker.clean.markSheet(2)
    assert(!tracker.isClean)
    assertEquals(tracker.modifiedSheets, Set(2))
  }

  test("delete removes sheet from modified set") {
    val tracker = ModificationTracker.clean.markSheet(1).delete(1)
    assert(tracker.deletedSheets.contains(1))
    assert(!tracker.modifiedSheets.contains(1))
  }

  test("merge combines changes") {
    val t1 = ModificationTracker.clean.markSheet(0)
    val t2 = ModificationTracker.clean.delete(2).markMetadata
    val merged = t1.merge(t2)
    assertEquals(merged.modifiedSheets, Set(0))
    assertEquals(merged.deletedSheets, Set(2))
    assert(merged.modifiedMetadata)
  }

  test("multiple deletions track correct indices") {
    // When deleting in ascending order, indices don't shift previous deletions
    // delete(2) then delete(5) means: delete sheet at index 2, then delete sheet at current index 5
    // After deleting 2, original sheet 5 is now at index 4, but we're deleting at index 5
    // (which is original sheet 6), so deletedSheets = {2, 5}
    val tracker = ModificationTracker.clean
      .delete(2) // deletedSheets = {2}
      .delete(5) // deletedSheets = {2, 5} (deleting current index 5, not original sheet 5)

    assertEquals(tracker.deletedSheets, Set(2, 5))
  }

  test("delete first, middle, last in sequence") {
    // Test edge cases: delete in various orders
    val tracker = ModificationTracker.clean
      .markSheet(0).markSheet(3).markSheet(6).markSheet(9)
      .delete(0) // Delete first (shifts 3→2, 6→5, 9→8), removes 0 from modified
      .delete(4) // Shifts 5→4, 8→7, then removes 4 (which was 5)
      .delete(7) // Removes 7 (which was 9)

    assertEquals(tracker.deletedSheets, Set(0, 4, 7))
    assertEquals(tracker.modifiedSheets, Set(2)) // Only original 3 remains (shifted twice: 3→2)
  }
