package com.tjclp.xl

import munit.ScalaCheckSuite
import org.scalacheck.Prop.*
import Generators.given

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
