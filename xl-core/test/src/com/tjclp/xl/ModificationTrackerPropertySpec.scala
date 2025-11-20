package com.tjclp.xl

import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import Generators.given
import com.tjclp.xl.context.ModificationTracker

/** Property-based tests for ModificationTracker Monoid laws and semantics */
class ModificationTrackerPropertySpec extends ScalaCheckSuite:

  // ========== Monoid Laws ==========

  property("merge is associative (Monoid law)") {
    forAll { (t1: ModificationTracker, t2: ModificationTracker, t3: ModificationTracker) =>
      val left = (t1.merge(t2)).merge(t3)
      val right = t1.merge(t2.merge(t3))
      assertEquals(left, right)
    }
  }

  property("merge left identity (clean merge a == a)") {
    forAll { (t: ModificationTracker) =>
      val result = ModificationTracker.clean.merge(t)
      assertEquals(result, t)
    }
  }

  property("merge right identity (a merge clean == a)") {
    forAll { (t: ModificationTracker) =>
      val result = t.merge(ModificationTracker.clean)
      assertEquals(result, t)
    }
  }

  property("merge is commutative for sets") {
    forAll { (t1: ModificationTracker, t2: ModificationTracker) =>
      val m1 = t1.merge(t2)
      val m2 = t2.merge(t1)

      // Set unions are commutative
      assertEquals(m1.modifiedSheets, m2.modifiedSheets)
      assertEquals(m1.deletedSheets, m2.deletedSheets)

      // Boolean OR is commutative
      assertEquals(m1.reorderedSheets, m2.reorderedSheets)
      assertEquals(m1.modifiedMetadata, m2.modifiedMetadata)
    }
  }

  // ========== Idempotence Properties ==========

  property("markSheet is idempotent") {
    forAll { (t: ModificationTracker) =>
      // Use small positive index for test efficiency
      val idx = 5
      val once = t.markSheet(idx)
      val twice = once.markSheet(idx)
      assertEquals(once, twice)
    }
  }

  property("markReordered is idempotent") {
    forAll { (t: ModificationTracker) =>
      val once = t.markReordered
      val twice = once.markReordered
      assertEquals(once, twice)
    }
  }

  property("markMetadata is idempotent") {
    forAll { (t: ModificationTracker) =>
      val once = t.markMetadata
      val twice = once.markMetadata
      assertEquals(once, twice)
    }
  }

  // ========== Semantic Properties ==========

  property("delete removes sheet from modified set") {
    // Use clean tracker for predictable behavior
    val idx = 7
    val marked = ModificationTracker.clean.markSheet(idx)
    val deleted = marked.delete(idx)

    // Should be in deletedSheets
    assert(deleted.deletedSheets.contains(idx))

    // Should NOT be in modifiedSheets
    assert(!deleted.modifiedSheets.contains(idx))
  }

  property("markSheets with empty set returns unchanged tracker") {
    forAll { (t: ModificationTracker) =>
      val result = t.markSheets(Set.empty)
      assertEquals(result, t)
    }
  }

  property("isClean detects any modification") {
    forAll { (t: ModificationTracker) =>
      val hasModifications =
        t.modifiedSheets.nonEmpty ||
        t.deletedSheets.nonEmpty ||
        t.reorderedSheets ||
        t.modifiedMetadata

      assertEquals(t.isClean, !hasModifications)
    }
  }

  property("delete shifts higher indices down") {
    // Use clean tracker to test index shifting behavior
    // Mark sheets 0, 5, 10 as modified
    val marked = ModificationTracker.clean.markSheet(0).markSheet(5).markSheet(10)

    // Delete sheet 5
    val deleted = marked.delete(5)

    // Verify indices adjusted correctly:
    // - Index 0 stays 0 (below deleted)
    // - Index 5 removed (was deleted)
    // - Index 10 becomes 9 (shifted down)
    assert(deleted.modifiedSheets.contains(0))
    assert(!deleted.modifiedSheets.contains(5))
    assert(deleted.modifiedSheets.contains(9))
    assert(!deleted.modifiedSheets.contains(10))
    assert(deleted.deletedSheets.contains(5))
  }

  property("multiple sequential deletions adjust indices correctly") {
    // Start with clean tracker, mark sheets at indices 2, 5, 10
    val marked = ModificationTracker.clean.markSheet(2).markSheet(5).markSheet(10)

    // Delete sheet 2 (shifts 5→4 and 10→9)
    val afterFirst = marked.delete(2)

    // Verify state after first deletion
    assert(afterFirst.deletedSheets.contains(2))
    assert(afterFirst.modifiedSheets.contains(4)) // Was 5
    assert(afterFirst.modifiedSheets.contains(9)) // Was 10
    assert(!afterFirst.modifiedSheets.contains(2)) // Deleted, removed from modified

    // Delete sheet 4 (which was originally sheet 5, shifts 9→8)
    val afterSecond = afterFirst.delete(4)

    // Verify final state
    assertEquals(afterSecond.deletedSheets, Set(2, 4))
    assertEquals(afterSecond.modifiedSheets, Set(8)) // Original sheet 10, shifted twice
  }

  property("deletedSheets indices shift on subsequent deletions") {
    // Test that deletedSheets itself gets shifted when subsequent deletions occur
    val tracker = ModificationTracker.clean
      .delete(5) // deletedSheets = {5}
      .delete(2) // Should shift 5→4, deletedSheets = {2, 4}

    assertEquals(tracker.deletedSheets, Set(2, 4))

    // Delete at index 4 (which is the shifted version of original index 5)
    val tracker2 = tracker.delete(4)
    assertEquals(tracker2.deletedSheets, Set(2, 4)) // Both deletions recorded
  }
