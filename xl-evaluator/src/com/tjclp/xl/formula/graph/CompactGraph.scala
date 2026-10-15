package com.tjclp.xl.formula.graph

import scala.collection.immutable.{BitSet, LongMap}
import scala.collection.mutable

import com.tjclp.xl.CellRange
import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.RecalcCounters
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.workbooks.Workbook

/**
 * The recalculation graph with one node per formula, per distinct range read, and per segment of a
 * formula column (#767).
 *
 * Node ids: `[0, F)` formula cells in grid order ([[NodeTable]]), `[F, F + S)` segment nodes,
 * `[F + S, F + S + R)` range nodes. Every edge points from a precedent to a node that reads it, and
 * all edges live in one CSR array:
 *
 *   - formula `p` → formula `f` when `f` reads the single cell `p` (a point read of the reader's
 *     own cell is kept as [[readsItself]] instead of an edge);
 *   - formula → the parent segment of its leaf in its column's segment tree, segment → its parent;
 *   - segment (or a lone leaf formula) → every range node it is a canonical piece of;
 *   - range node → every formula that reads the range.
 *
 * A formula column that some range covers gets an implicit segment tree over its formula rows
 * (nodes `1..m-1` internal, `m..2m-1` the formulas themselves). A range over rows `a..b` of the
 * column takes edges from the O(log m) canonical nodes covering the formulas inside `a..b`, so a
 * formula reaches a range node exactly when it lies inside the range, and a running-total column
 * costs O(rows log rows) edges instead of O(rows²).
 *
 * Reachability between formulas is the legacy per-cell graph's (`fromWorkbookFormulaGraph`): a
 * legacy edge `c → f` (formula `c` inside a range `f` reads) is the path `c → … → canonical segment
 * → range → f`, and every virtual path ends at a reader of a range that contains its start. Hence
 * the same cycles, the same transitive dependents and, with the min-heap order of [[kahn]], the
 * same lexicographically smallest topological order of the formulas.
 *
 * Constant cells are not nodes. [[foreachReaderOfCell]] finds a cell's direct readers through the
 * point reads of constants and an [[IntervalIndex]] of the range nodes per column.
 *
 * Segment trees over formula columns follow GridCraft's #66 range-node design
 * (`crates/calc/src/recalc.rs`, MIT/Apache-2.0); the canonical decomposition is the bottom-up
 * segment-tree query.
 */
private[formula] final class CompactGraph private (
  val nodes: NodeTable,
  val segmentCount: Int,
  rangeSheet: IArray[Int],
  rangeRow0: IArray[Int],
  rangeCol0: IArray[Int],
  rangeRow1: IArray[Int],
  rangeCol1: IArray[Int],
  succOff: IArray[Int],
  succDst: IArray[Int],
  selfReads: BitSet,
  columnIndex: LongMap[IntervalIndex],
  wideIndex: Map[Int, IntervalIndex],
  constKeys: IArray[Long],
  constReaders: IArray[Int]
):
  import CompactGraph.*

  /** F: formula nodes. */
  def formulaCount: Int = nodes.formulaCount

  /** R: distinct (sheet, range) nodes. */
  def rangeCount: Int = rangeSheet.length

  /** F + S + R. */
  def nodeCount: Int = formulaCount + segmentCount + rangeCount

  /** Stored edges. */
  def edgeCount: Int = succDst.length

  def isFormula(n: Int): Boolean = n >= 0 && n < formulaCount

  /** The node id of range `i` (`0 <= i < rangeCount`). */
  def rangeNode(i: Int): Int = formulaCount + segmentCount + i

  /** Range `i` as (sheet name, range). */
  def range(i: Int): (SheetName, CellRange) =
    (
      nodes.sheetNames(rangeSheet(i)),
      CellRange(
        com.tjclp.xl.addressing.ARef.from0(rangeCol0(i), rangeRow0(i)),
        com.tjclp.xl.addressing.ARef.from0(rangeCol1(i), rangeRow1(i))
      )
    )

  /** Whether formula `f` reads its own cell through a point reference. */
  def readsItself(f: Int): Boolean = selfReads.contains(f)

  /** Calls `f` with every successor (reader-side neighbour) of node `n`, in insertion order. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def foreachSucc(n: Int)(f: Int => Unit): Unit =
    var e = succOff(n)
    val end = succOff(n + 1)
    while e < end do
      f(succDst(e))
      e += 1

  /**
   * Calls `f` with every formula that reads the cell (sheet id, 0-based row, 0-based column)
   * directly: by a point reference or through a range containing it. A reader of both is reported
   * twice.
   */
  def foreachReaderOfCell(sheet: Int, row: Int, col: Int)(f: Int => Unit): Unit =
    val self = nodes.indexOf(sheet, row, col)
    if self >= 0 then
      if readsItself(self) then f(self)
      foreachSucc(self)(n => if n < formulaCount then f(n))
    else
      val key = NodeTable.cellKey(sheet, row, col)
      if key >= 0L then constIndex.get(key).foreach(_.foreach(f))
    val readRange: Int => Unit = i => foreachSucc(rangeNode(i))(f)
    columnIndex.get(columnKey(sheet, col)).foreach(_.foreachStab(row)(readRange))
    wideIndex
      .get(sheet)
      .foreach(
        _.foreachStab(row)(i => if rangeCol0(i) <= col && col <= rangeCol1(i) then readRange(i))
      )

  /** Point reads of constant cells, grouped by cell key; built on first use. */
  private lazy val constIndex: LongMap[IArray[Int]] =
    val grouped = mutable.LongMap.empty[IntBuf]
    constKeys.indices.foreach { i =>
      grouped.getOrElseUpdate(constKeys(i), new IntBuf(4)) += constReaders(i)
    }
    LongMap.from(grouped.iterator.map((k, buf) => k -> IArray.unsafeFromArray(buf.toArray)))

  /**
   * Strongly connected components, by an iterative Tarjan over the CSR (a resumable cursor per
   * frame, so a 100k-deep chain needs no stack). A component is cyclic when it had more than one
   * node before its virtual nodes are discounted, or is one formula reading its own cell (design
   * §3.2): `A5 = SUM(A:A)` forms `{A5, segments, A:A}`, as the legacy self-edge did.
   */
  lazy val components: Components = tarjan()

  /** The formulas of every cyclic component. */
  def cyclicFormulas: BitSet = components.cyclic

  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def tarjan(): Components =
    val n = nodeCount
    val f = formulaCount
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
              val isCyclic = size > 1 || (v < f && selfReads.contains(v))
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
   * Every formula reachable from `seeds` (formula ids) along the edges, the seeds excluded, never
   * entering a formula of `skip`. The legacy `qualifiedTransitiveDependents` over formula edges.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def reach(seeds: BitSet, skip: BitSet): Reach =
    val n = nodeCount
    val f = formulaCount
    val visited = new Array[Boolean](n)
    val queue = new Array[Int](n)
    var head = 0
    var tail = 0
    seeds.foreach { s =>
      if s >= 0 && s < f && !visited(s) then
        visited(s) = true
        queue(tail) = s
        tail += 1
    }
    val found = mutable.BitSet.empty
    var work = 0L
    while head < tail do
      val u = queue(head)
      head += 1
      var e = succOff(u)
      val end = succOff(u + 1)
      while e < end do
        val s = succDst(e)
        work += 1
        if !visited(s) && !(s < f && skip.contains(s)) then
          visited(s) = true
          queue(tail) = s
          tail += 1
          if s < f then found += s
        e += 1
    Reach((found --= seeds).toImmutable, work)

  /**
   * The canonical evaluation order of every formula outside `removed`: Kahn's algorithm with a
   * min-heap on formula id, i.e. the lexicographically smallest topological order in grid order
   * (design §3.4).
   *
   * In-degrees count only edges from nodes that remain (a removed formula never holds a segment or
   * range node above zero). A segment or range node is expanded the moment its in-degree reaches
   * zero, never queued, so a formula becomes ready exactly when all its formula ancestors have been
   * emitted. `nodeOrder` lists every node processed, virtual ones included, in a topological order
   * ([[waveDepths]] walks it).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def kahn(removed: BitSet): KahnResult =
    val n = nodeCount
    val f = formulaCount
    def gone(v: Int): Boolean = v < f && removed.contains(v)
    val indeg = new Array[Int](n)
    var work = 0L
    var u = 0
    while u < n do
      if !gone(u) then
        var e = succOff(u)
        val end = succOff(u + 1)
        while e < end do
          indeg(succDst(e)) += 1
          work += 1
          e += 1
      u += 1
    val heap = new IntMinHeap(f)
    val virtual = new Array[Int](math.max(1, n - f))
    var vsp = 0
    val formulaOrder = new IntBuf(math.max(16, f))
    val nodeOrder = new IntBuf(math.max(16, n))

    def release(v: Int): Unit =
      if v < f then heap.push(v)
      else
        virtual(vsp) = v
        vsp += 1

    def visit(v: Int): Unit =
      var e = succOff(v)
      val end = succOff(v + 1)
      while e < end do
        val s = succDst(e)
        work += 1
        if !gone(s) then
          indeg(s) -= 1
          if indeg(s) == 0 then release(s)
        e += 1

    def drainVirtual(): Unit =
      while vsp > 0 do
        vsp -= 1
        val v = virtual(vsp)
        nodeOrder += v
        visit(v)

    u = 0
    while u < n do
      if indeg(u) == 0 && !gone(u) then release(u)
      u += 1
    drainVirtual()
    while heap.nonEmpty do
      val next = heap.pop()
      formulaOrder += next
      nodeOrder += next
      visit(next)
      drainVirtual()

    val order = formulaOrder.toArray
    val expected = f - removed.count(_ < f)
    val remaining =
      if order.length == expected then BitSet.empty
      else
        val emitted = mutable.BitSet.empty
        order.foreach(emitted += _)
        BitSet.fromSpecific((0 until f).iterator.filter(v => !removed.contains(v) && !emitted(v)))
    KahnResult(
      IArray.unsafeFromArray(order),
      IArray.unsafeFromArray(nodeOrder.toArray),
      remaining,
      work
    )

  /**
   * Longest-path depth of every formula of `members` over the formula reachability relation
   * restricted to `members`: 0 with no member ancestor, else one more than the deepest member
   * ancestor. A virtual node passes on the largest depth reaching it, and nothing when no member
   * reaches it, so a formula reading only constant ranges stays at depth 0. Formulas outside
   * `members` get -1. `kahn` must be this graph's [[kahn]] result for a removal set disjoint from
   * `members`.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def waveDepths(kahn: KahnResult, members: BitSet): IArray[Int] =
    val n = nodeCount
    val f = formulaCount
    val NoDepth = Int.MinValue
    val value = Array.fill(n)(NoDepth)
    val depth = Array.fill(f)(-1)
    val order = kahn.nodeOrder
    var i = 0
    while i < order.length do
      val v = order(i)
      val out =
        if v < f then
          if members.contains(v) then
            val d = math.max(0, value(v))
            depth(v) = d
            d + 1
          else NoDepth
        else value(v)
      if out != NoDepth then
        var e = succOff(v)
        val end = succOff(v + 1)
        while e < end do
          val s = succDst(e)
          if value(s) < out then value(s) = out
          e += 1
      i += 1
    IArray.unsafeFromArray(depth)

  /**
   * The SCC condensation restricted to formulas, in the key-minimum dependency-first order of
   * `DependencyGraph.qualifiedSccOrder`: every component after the components it reads, ties broken
   * by the smallest member id (grid order). Each entry is a component's formula ids ascending and
   * whether it is cyclic. Components made only of virtual nodes are expanded in passing, as in
   * [[kahn]].
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def condensation: Vector[(IArray[Int], Boolean)] =
    val comps = components
    val n = nodeCount
    val f = formulaCount
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

  /**
   * Member-to-member reads inside a set of formulas: for each member, the members it reads, found
   * by walking from every member through virtual nodes up to the next formula. Self-reads are left
   * out (they are back edges by definition in `withinComponentOrder`).
   */
  @SuppressWarnings(Array("org.wartremover.warts.While"))
  def memberPrecedents(members: BitSet): Map[Int, Set[Int]] =
    val f = formulaCount
    val acc = mutable.HashMap.empty[Int, Set[Int]]
    members.foreach { u =>
      val seen = mutable.HashSet.empty[Int]
      val pending = mutable.ArrayDeque.empty[Int]
      foreachSucc(u)(pending.append)
      while pending.nonEmpty do
        val s = pending.removeHead()
        if seen.add(s) then
          if s < f then
            if s != u && members.contains(s) then acc.update(s, acc.getOrElse(s, Set.empty) + u)
          else foreachSucc(s)(pending.append)
    }
    acc.toMap

private[formula] object CompactGraph:

  /** Tarjan's result: component per node, whether each is cyclic, and the cyclic formulas. */
  final case class Components(
    compOf: IArray[Int],
    count: Int,
    compCyclic: IArray[Int],
    cyclic: BitSet,
    work: Long
  )

  /** Formulas reached by [[CompactGraph.reach]] and the edges it walked. */
  final case class Reach(formulas: BitSet, work: Long)

  /**
   * [[CompactGraph.kahn]]'s result: the formula order, every processed node in order, the formulas
   * it could not order (empty unless a cycle was left in), and the edges it walked.
   */
  final case class KahnResult(
    formulaOrder: IArray[Int],
    nodeOrder: IArray[Int],
    remaining: BitSet,
    work: Long
  )

  /** Ranges wider than this many columns are indexed per sheet, filtered by column at stab time. */
  val Wide: Int = 64

  private def columnKey(sheet: Int, col: Int): Long =
    (sheet.toLong << 32) | (col.toLong & 0xffffffffL)

  /**
   * Build the graph of `wb`. Each formula is parsed once (timed as [[RecalcCounters.Phase.Parse]])
   * and its reads extracted with the legacy extractor in range-recording mode, with the precise
   * VLOOKUP/HLOOKUP strips of the ordering graph; each defined-name body is parsed once per build.
   * `counters` also receives the node and edge counts.
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  def build(wb: Workbook, counters: RecalcCounters = RecalcCounters.disabled): CompactGraph =
    val (nodes, texts) = NodeTable.build(wb)
    val f = nodes.formulaCount

    // ---- 1. reads of every formula -----------------------------------------------------------
    val canonicalSheet = DependencyGraph.sheetCanonicaliser(wb)
    val nameRanges = DependencyGraph.NameRanges.empty
    val nameBodies = mutable.HashMap.empty[String, Option[TExpr[?]]]
    val parseName: String => Option[TExpr[?]] =
      text => nameBodies.getOrElseUpdate(text, FormulaParser.parse(text).toOption)
    val rangeIds = mutable.HashMap.empty[(Int, Long, Long), Int]
    val rSheet, rRow0, rCol0, rRow1, rCol1 = new IntBuf()
    val pointSrc, pointDst = new IntBuf()
    val readerRange, readerDst = new IntBuf()
    val constKeys = new LongBuf()
    val constReaders = new IntBuf()
    val selfReads = mutable.BitSet.empty
    val scratch = new IntBuf()

    def intern(sheet: Int, range: CellRange): Int =
      val r0 = range.start.row.index0
      val c0 = range.start.col.index0
      val r1 = range.end.row.index0
      val c1 = range.end.col.index0
      rangeIds.getOrElseUpdate(
        (
          sheet,
          (r0.toLong << 32) | (c0.toLong & 0xffffffffL),
          (r1.toLong << 32) | (c1.toLong & 0xffffffffL)
        ), {
          rSheet += sheet
          rRow0 += r0
          rCol0 += c0
          rRow1 += r1
          rCol1 += c1
          rSheet.length - 1
        }
      )

    val record: (SheetName, CellRange) => Set[QualifiedRef] = (target, range) =>
      val s = nodes.sheetOf(target)
      // a range on a sheet the workbook lacks holds no cell, so it orders nothing
      if s >= 0 then scratch += intern(s, range)
      Set.empty

    var i = 0
    while i < f do
      val sheetName = nodes.sheetNames(nodes.sheet(i))
      scratch.clear()
      counters.time(RecalcCounters.Phase.Parse)(FormulaParser.parse(texts(i))) match
        case Right(expr) =>
          val points = DependencyGraph.extractQualifiedDependencies(
            expr,
            sheetName,
            record,
            Some(wb),
            preciseLookups = true,
            canonicalSheet = canonicalSheet,
            nameRanges = nameRanges,
            parseName = parseName
          )
          val reader = i
          points.foreach { p =>
            val s = nodes.sheetOf(p.sheet)
            if s >= 0 then
              val row = p.ref.row.index0
              val col = p.ref.col.index0
              val j = nodes.indexOf(s, row, col)
              if j == reader then selfReads += reader
              else if j >= 0 then
                pointSrc += j
                pointDst += reader
              else
                // a parsed reference always lies on Excel's grid, so its key exists
                val key = NodeTable.cellKey(s, row, col)
                if key >= 0L then
                  constKeys += key
                  constReaders += reader
          }
          scratch.sortDistinct()
          var k = 0
          while k < scratch.length do
            readerRange += scratch(k)
            readerDst += reader
            k += 1
        case Left(_) => () // an unparseable formula reads nothing (as in the legacy graph)
      i += 1

    val r = rSheet.length

    // ---- 2. formula columns: rows ascending, ids alongside -----------------------------------
    // ids are in (sheet, row, col) order, so appending by id keeps each column's rows ascending
    val colRows = mutable.LongMap.empty[IntBuf]
    val colIds = mutable.LongMap.empty[IntBuf]
    i = 0
    while i < f do
      val key = columnKey(nodes.sheet(i), nodes.col(i))
      colRows.getOrElseUpdate(key, new IntBuf(4)) += nodes.row(i)
      colIds.getOrElseUpdate(key, new IntBuf(4)) += i
      i += 1
    val rowsOf: LongMap[Array[Int]] = LongMap.from(colRows.iterator.map((k, b) => k -> b.toArray))
    val idsOf: LongMap[Array[Int]] = LongMap.from(colIds.iterator.map((k, b) => k -> b.toArray))
    // per sheet, its formula columns ascending
    val sheetCols: Map[Int, Array[Int]] =
      rowsOf.keysIterator.toVector
        .groupMap(k => (k >>> 32).toInt)(k => (k & 0xffffffffL).toInt)
        .view
        .mapValues(_.sorted.toArray)
        .toMap

    // ---- 3. each range's spans over formula columns -----------------------------------------
    val spanRange, spanStart, spanEnd = new IntBuf()
    val spanColumn = new LongBuf()
    val covered = mutable.LongMap.empty[Boolean]
    var ri = 0
    while ri < r do
      val s = rSheet(ri)
      val cols = sheetCols.getOrElse(s, Array.emptyIntArray)
      var ci = lowerBound(cols, rCol0(ri))
      while ci < cols.length && cols(ci) <= rCol1(ri) do
        val key = columnKey(s, cols(ci))
        val rows = rowsOf.getOrElse(key, Array.emptyIntArray)
        val a = lowerBound(rows, rRow0(ri))
        val b = lowerBound(rows, rRow1(ri) + 1) - 1
        if a <= b then
          spanRange += ri
          spanColumn += key
          spanStart += a
          spanEnd += b
          covered.update(key, true)
        ci += 1
      ri += 1

    // ---- 4. segment ids: internal nodes 1..m-1 of each covered column with m >= 2 ------------
    val coveredKeys = covered.keysIterator.toArray
    java.util.Arrays.sort(coveredKeys)
    val segBase = mutable.LongMap.empty[Int]
    var s = 0
    coveredKeys.foreach { key =>
      val m = rowsOf.getOrElse(key, Array.emptyIntArray).length
      segBase.update(key, s)
      if m >= 2 then s += m - 1
    }
    val segments = s
    def segNode(key: Long, k: Int): Int = f + segBase.getOrElse(key, 0) + (k - 1)
    def rangeNode(ri: Int): Int = f + segments + ri

    // ---- 5. edges ----------------------------------------------------------------------------
    val src = new IntBuf(math.max(16, pointSrc.length + readerDst.length + 2 * f))
    val dst = new IntBuf(math.max(16, pointSrc.length + readerDst.length + 2 * f))
    def edge(u: Int, v: Int): Unit =
      src += u
      dst += v
    i = 0
    while i < pointSrc.length do
      edge(pointSrc(i), pointDst(i))
      i += 1
    coveredKeys.foreach { key =>
      val ids = idsOf.getOrElse(key, Array.emptyIntArray)
      val m = ids.length
      if m >= 2 then
        var leaf = 0
        while leaf < m do
          edge(ids(leaf), segNode(key, (m + leaf) / 2))
          leaf += 1
        var k = 2
        while k < m do
          edge(segNode(key, k), segNode(key, k / 2))
          k += 1
    }
    i = 0
    while i < spanRange.length do
      val key = spanColumn(i)
      val ids = idsOf.getOrElse(key, Array.emptyIntArray)
      val m = ids.length
      val target = rangeNode(spanRange(i))
      // bottom-up canonical decomposition of leaves [spanStart, spanEnd]
      var lo = spanStart(i) + m
      var hi = spanEnd(i) + m + 1
      while lo < hi do
        if (lo & 1) == 1 then
          edge(if lo >= m then ids(lo - m) else segNode(key, lo), target)
          lo += 1
        if (hi & 1) == 1 then
          hi -= 1
          edge(if hi >= m then ids(hi - m) else segNode(key, hi), target)
        lo >>>= 1
        hi >>>= 1
      i += 1
    i = 0
    while i < readerDst.length do
      edge(rangeNode(readerRange(i)), readerDst(i))
      i += 1

    // ---- 6. CSR ------------------------------------------------------------------------------
    val n = f + segments + r
    val off = new Array[Int](n + 1)
    i = 0
    while i < src.length do
      off(src(i) + 1) += 1
      i += 1
    i = 0
    while i < n do
      off(i + 1) += off(i)
      i += 1
    val succ = new Array[Int](src.length)
    val cursor = java.util.Arrays.copyOf(off, n)
    i = 0
    while i < src.length do
      val u = src(i)
      succ(cursor(u)) = dst(i)
      cursor(u) += 1
      i += 1

    // ---- 7. interval indexes over range nodes (constant cells → covering ranges) -------------
    val narrow = mutable.LongMap.empty[(IntBuf, IntBuf, IntBuf)]
    val wide = mutable.HashMap.empty[Int, (IntBuf, IntBuf, IntBuf)]
    ri = 0
    while ri < r do
      val width = rCol1(ri) - rCol0(ri) + 1
      if width <= Wide then
        var c = rCol0(ri)
        while c <= rCol1(ri) do
          val (a, b, d) = narrow.getOrElseUpdate(
            columnKey(rSheet(ri), c),
            (new IntBuf(4), new IntBuf(4), new IntBuf(4))
          )
          a += rRow0(ri)
          b += rRow1(ri)
          d += ri
          c += 1
      else
        val (a, b, d) =
          wide.getOrElseUpdate(rSheet(ri), (new IntBuf(4), new IntBuf(4), new IntBuf(4)))
        a += rRow0(ri)
        b += rRow1(ri)
        d += ri
      ri += 1
    def index(t: (IntBuf, IntBuf, IntBuf)): IntervalIndex =
      IntervalIndex.build(t._1.toArray, t._2.toArray, t._3.toArray)

    if counters.enabled then
      counters.add(RecalcCounters.Counter.GraphNodes, f.toLong)
      counters.add(RecalcCounters.Counter.GraphEdges, succ.length.toLong)
      counters.add(RecalcCounters.Counter.VirtualNodes, segments.toLong)
      counters.add(RecalcCounters.Counter.RangeNodes, r.toLong)

    new CompactGraph(
      nodes,
      segments,
      IArray.unsafeFromArray(rSheet.toArray),
      IArray.unsafeFromArray(rRow0.toArray),
      IArray.unsafeFromArray(rCol0.toArray),
      IArray.unsafeFromArray(rRow1.toArray),
      IArray.unsafeFromArray(rCol1.toArray),
      IArray.unsafeFromArray(off),
      IArray.unsafeFromArray(succ),
      selfReads.toImmutable,
      LongMap.from(narrow.iterator.map((k, t) => k -> index(t))),
      wide.iterator.map((k, t) => k -> index(t)).toMap,
      IArray.unsafeFromArray(constKeys.toArray),
      IArray.unsafeFromArray(constReaders.toArray)
    )

  /** First index of `sorted` whose element is at least `x`. */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def lowerBound(sorted: Array[Int], x: Int): Int =
    var lo = 0
    var hi = sorted.length
    while lo < hi do
      val mid = (lo + hi) >>> 1
      if sorted(mid) < x then lo = mid + 1 else hi = mid
    lo
