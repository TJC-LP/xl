package com.tjclp.xl.formula.graph

import munit.ScalaCheckSuite
import org.scalacheck.{Gen, Prop}

/** #767: the open-addressing `Long -> Int` table is a `Map` (last write wins). */
class LongIntMapSpec extends ScalaCheckSuite:

  private val genKey: Gen[Long] = Gen.frequency(
    3 -> Gen.choose(-50L, 50L),
    2 -> Gen.long,
    // packed cell keys: sheet << 34 | row << 14 | col, the shape NodeTable stores
    3 -> Gen
      .zip(Gen.choose(0, 3), Gen.choose(0, 2000), Gen.choose(0, 40))
      .map((s, r, c) => NodeTable.cellKey(s, r, c))
  )

  property("indexOf and get agree with Map on present and absent keys") {
    Prop.forAll(Gen.listOf(Gen.zip(genKey, Gen.choose(0, 1000))), Gen.listOf(genKey)) {
      (pairs, probes) =>
        val table = LongIntMap.from(pairs.map(_._1).toArray, pairs.map(_._2).toArray, pairs.size)
        val model = pairs.toMap
        Prop.all(
          Prop(table.size == model.size) :| "size",
          Prop((pairs.map(_._1) ++ probes).forall { k =>
            table.get(k) == model.get(k) && table.indexOf(k) == model.getOrElse(k, -1)
          }) :| "lookups"
        )
    }
  }

  test("100k packed cell keys of one column: every key found, absent neighbours not") {
    val n = 100000
    val keys = Array.tabulate(n)(r => NodeTable.cellKey(0, r, 2))
    val table = LongIntMap.from(keys, Array.tabulate(n)(identity), n)
    assertEquals(table.size, n)
    assert((0 until n).forall(r => table.indexOf(keys(r)) == r))
    assert((0 until 1000).forall(r => table.indexOf(NodeTable.cellKey(0, r, 3)) == -1))
  }

  test("cellKey: grid order and -1 off the grid") {
    assert(NodeTable.cellKey(0, 5, 9) < NodeTable.cellKey(0, 6, 0))
    assert(NodeTable.cellKey(0, 1048575, 16383) < NodeTable.cellKey(1, 0, 0))
    assertEquals(NodeTable.cellKey(0, 1048576, 0), -1L)
    assertEquals(NodeTable.cellKey(0, 0, 16384), -1L)
    assertEquals(NodeTable.cellKey(-1, 0, 0), -1L)
  }
