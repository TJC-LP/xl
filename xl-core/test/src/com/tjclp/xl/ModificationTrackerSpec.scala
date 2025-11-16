package com.tjclp.xl

import munit.FunSuite

class ModificationTrackerSpec extends FunSuite:

  test("clean tracker reports no modifications") {
    assert(ModificationTracker.clean.isClean)
  }

  test("markSheet marks the provided index") {
    val tracker = ModificationTracker.clean.markSheet(2)
    assert(!tracker.isClean)
    assertEquals(tracker.modifiedSheets, Set(2))
  }

  test("deleteSheet removes sheet from modified set") {
    val tracker = ModificationTracker.clean.markSheet(1).deleteSheet(1)
    assert(tracker.deletedSheets.contains(1))
    assert(!tracker.modifiedSheets.contains(1))
  }

  test("merge combines changes") {
    val t1 = ModificationTracker.clean.markSheet(0)
    val t2 = ModificationTracker.clean.deleteSheet(2).markMetadata
    val merged = t1.merge(t2)
    assertEquals(merged.modifiedSheets, Set(0))
    assertEquals(merged.deletedSheets, Set(2))
    assert(merged.modifiedMetadata)
  }
