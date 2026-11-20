package com.tjclp.xl.formula.graph

import scala.collection.immutable.BitSet
import scala.collection.mutable

import com.tjclp.xl.formula.graph.CompactGraph.Components

/**
 * Tarjan's strongly connected components and the key-minimum condensation order over a graph held
 * as one CSR edge array (#767, #891).
 *
 * Nodes are the dense ids `[0, n)`; the formulas are `[0, f)` in grid order and the ids `[f, n)`
 * are virtual nodes (the compact graph's segments and ranges). Every edge points from a precedent
 * to a node that reads it: the successors of `v` are `succDst(succOff(v) until succOff(v + 1))`. A
 * formula that reads itself is marked in `selfLoops` instead of carrying an edge to itself.
 *
 * Both walks run on `Int` arrays sized once: no node is hashed, boxed or looked up in a map. The
 * compact recalculation graph and the workbook graph of `DependencyGraph.qualifiedSccOrder` share
 * them.
 */
private[graph] object SccKernel:

  /**
   * The strongly connected components, by an iterative Tarjan with a resumable cursor per frame
   * (stack-safe on a 100k-long chain). A component is cyclic when it has more than one node or is a
   * self-looping formula.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def tarjan(
    n: Int,
    f: Int,
    succOff: IArray[Int],
    succDst: IArray[Int],
    selfLoops: BitSet
  ): Components =
    val index = Array.fill(n)(-1)
    val low = new Array[Int](n)
    val onStack = new Array[Boolean](n)
    val stack = new Array[Int](n)
    val frameNode = new Array[Int](n)
    val frameCursor = new Array[Int](n)
    val compOf = new Array[Int](n)
    val compCyclic = new IntBuf()
    val cyclic = mutable.BitSet.empty
    var sp = 0
    var fp = 0
    var next = 0
    var comps = 0
    var work = 0L
    var root = 0
    while root < n do
      if index(root) < 0 then
        index(root) = next
        low(root) = next
        next += 1
        stack(sp) = root
        sp += 1
        onStack(root) = true
        frameNode(fp) = root
        frameCursor(fp) = succOff(root)
        fp += 1
        while fp > 0 do
          val v = frameNode(fp - 1)
          val c = frameCursor(fp - 1)
          if c < succOff(v + 1) then
            frameCursor(fp - 1) = c + 1
            work += 1
            val w = succDst(c)
            if index(w) < 0 then
              index(w) = next
              low(w) = next
              next += 1
              stack(sp) = w
              sp += 1
              onStack(w) = true
              frameNode(fp) = w
              frameCursor(fp) = succOff(w)
              fp += 1
            else if onStack(w) && index(w) < low(v) then low(v) = index(w)
          else
            fp -= 1
            if low(v) == index(v) then
              val top = sp
              var done = false
              while !done do
                sp -= 1
                val w = stack(sp)
                onStack(w) = false
                compOf(w) = comps
                done = w == v
              val size = top - sp
              val isCyclic = size > 1 || (v < f && selfLoops.contains(v))
              if isCyclic then
                var k = sp
                while k < top do
                  if stack(k) < f then cyclic += stack(k)
                  k += 1
              compCyclic += (if isCyclic then 1 else 0)
              comps += 1
            if fp > 0 then
              val parent = frameNode(fp - 1)
              if low(v) < low(parent) then low(parent) = low(v)
      root += 1
    Components(
      IArray.unsafeFromArray(compOf),
      comps,
      IArray.unsafeFromArray(compCyclic.toArray),
      cyclic.toImmutable,
      work
    )

  /**
   * The SCC condensation restricted to formulas, in the key-minimum dependency-first order: every
   * component after the components it reads, ties broken by the smallest member id (grid order).
   * Each entry is a component's formula ids ascending and whether it is cyclic. Components made
   * only of virtual nodes are expanded in passing.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def condensation(
    n: Int,
    f: Int,
    succOff: IArray[Int],
    succDst: IArray[Int],
    comps: Components
  ): Vector[(IArray[Int], Boolean)] =
    val c = comps.count
    // members per component (formulas only), ascending, by counting sort
    val memberCount = new Array[Int](c + 1)
    var v = 0
    while v < f do
      memberCount(comps.compOf(v) + 1) += 1
      v += 1
    var k = 0
    while k < c do
      memberCount(k + 1) += memberCount(k)
      k += 1
    val memberOff = memberCount.clone()
    val members = new Array[Int](f)
    val fill = memberCount.clone()
    v = 0
    while v < f do
      val cv = comps.compOf(v)
      members(fill(cv)) = v
      fill(cv) += 1
      v += 1
    // condensation edges (one per underlying edge; in-degrees count them the same way)
    val outCount = new Array[Int](c + 1)
    val indeg = new Array[Int](c)
    var u = 0
    while u < n do
      val cu = comps.compOf(u)
      var e = succOff(u)
      val end = succOff(u + 1)
      while e < end do
        val cs = comps.compOf(succDst(e))
        if cs != cu then
          outCount(cu + 1) += 1
          indeg(cs) += 1
        e += 1
      u += 1
    k = 0
    while k < c do
      outCount(k + 1) += outCount(k)
      k += 1
    val outDst = new Array[Int](outCount(c))
    val outFill = outCount.clone()
    u = 0
    while u < n do
      val cu = comps.compOf(u)
      var e = succOff(u)
      val end = succOff(u + 1)
      while e < end do
        val cs = comps.compOf(succDst(e))
        if cs != cu then
          outDst(outFill(cu)) = cs
          outFill(cu) += 1
        e += 1
      u += 1

    def key(comp: Int): Int =
      if memberOff(comp + 1) > memberOff(comp) then members(memberOff(comp)) else -1
    val heap = new IntMinHeap(f)
    val virtual = new Array[Int](math.max(1, c))
    var vsp = 0
    def release(comp: Int): Unit =
      val kk = key(comp)
      if kk >= 0 then heap.push(kk)
      else
        virtual(vsp) = comp
        vsp += 1
    def visit(comp: Int): Unit =
      var e = outCount(comp)
      val end = outCount(comp + 1)
      while e < end do
        val s = outDst(e)
        indeg(s) -= 1
        if indeg(s) == 0 then release(s)
        e += 1
    def drainVirtual(): Unit =
      while vsp > 0 do
        vsp -= 1
        visit(virtual(vsp))
    k = 0
    while k < c do
      if indeg(k) == 0 then release(k)
      k += 1
    drainVirtual()
    val out = Vector.newBuilder[(IArray[Int], Boolean)]
    while heap.nonEmpty do
      val comp = comps.compOf(heap.pop())
      val listed = java.util.Arrays.copyOfRange(members, memberOff(comp), memberOff(comp + 1))
      out += ((IArray.unsafeFromArray(listed), comps.compCyclic(comp) == 1))
      visit(comp)
      drainVirtual()
    out.result()
