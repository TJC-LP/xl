package com.tjclp.xl.formula.graph

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/**
 * #767: the static interval index answers every overlap query as a linear scan does, and a stab
 * costs O(log n) tree nodes per match (the cost law ported from GridCraft's
 * `interval_queries_cost_log_n_plus_matches`).
 */
class IntervalIndexSpec extends ScalaCheckSuite:

  private def found(index: IntervalIndex, a: Int, b: Int): Vector[Int] =
    val out = Vector.newBuilder[Int]
    index.foreachOverlap(a, b)(out += _)
    out.result().sorted

  test("20,000 random intervals: every query equals a linear scan") {
    val rng = new scala.util.Random(0x9e3779b97f4a7c15L)
    val n = 20000
    val starts = Array.fill(n)(rng.nextInt(500))
    val ends = starts.map(s => s + rng.nextInt(40))
    val ids = Array.tabulate(n)(identity)
    val index = IntervalIndex.build(starts, ends, ids)
    assertEquals(index.size, n)
    def linear(a: Int, b: Int): Vector[Int] =
      ids.iterator.filter(i => starts(i) <= b && ends(i) >= a).toVector.sorted
    (0 until 300).foreach { _ =>
      val a = rng.nextInt(560)
      val b = a + rng.nextInt(30)
      assertEquals(found(index, a, b), linear(a, b), s"rows $a..$b")
    }
    (0 until 560).foreach { row =>
      val stabbed = Vector.newBuilder[Int]
      index.foreachStab(row)(stabbed += _)
      assertEquals(stabbed.result().sorted, linear(row, row), s"row $row")
    }
  }

  property("small random sets: overlap equals a linear scan, duplicates kept") {
    val genIntervals = Gen.listOf(Gen.zip(Gen.choose(0, 30), Gen.choose(0, 8)))
    Prop.forAll(genIntervals, Gen.choose(-2, 40), Gen.choose(0, 10)) { (spans, a, width) =>
      val starts = spans.map(_._1).toArray
      val ends = spans.map((s, w) => s + w).toArray
      val ids = Array.tabulate(spans.size)(i => i % 5)
      val index = IntervalIndex.build(starts, ends, ids)
      val b = a + width
      val expected = spans.indices.filter(i => starts(i) <= b && ends(i) >= a).map(ids).toVector
      found(index, a, b) == expected.sorted
    }
  }

  test("empty and single-interval indexes") {
    assertEquals(found(IntervalIndex.empty, 0, 100), Vector.empty)
    val one = IntervalIndex.build(Array(5), Array(9), Array(42))
    assertEquals(found(one, 0, 4), Vector.empty)
    assertEquals(found(one, 9, 20), Vector(42))
    assertEquals(found(one, 6, 6), Vector(42))
  }

  test("a stab visits O(log n) nodes per match: 40k one-row ranges plus 1k covering ones") {
    val n = 40000
    val starts = Array.tabulate(n + 1000)(i => if i < n then i else 0)
    val ends = Array.tabulate(n + 1000)(i => if i < n then i else n)
    val ids = Array.tabulate(n + 1000)(identity)
    val index = IntervalIndex.build(starts, ends, ids)
    val worst = (0 until n by 7).map { row =>
      var k = 0
      index.foreachStab(row)(_ => k += 1)
      assertEquals(k, 1001)
      index.visitCount(row)
    }.max
    // 1,001 matches in a tree of depth 16: under four nodes per match plus the path
    assert(worst < 4 * 1001 + 64, s"visited $worst nodes")
    assert(index.visitCount(n + 1) < 64, "a query past every interval walks one path")
  }
