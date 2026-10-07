package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.workbooks.{DefinedName, Workbook}

import scala.annotation.tailrec

/**
 * #691: the rules every defined-name walk shares — the evaluator's resolution, the dependency
 * graph's name edges, unresolved readers, the name-change readers, dynamic-reference classification
 * and the audit's volatility check.
 *
 *   - A reference resolves to a [[NameWalk.Node]]: the definition the scoped lookup finds and the
 *     sheet its body resolves from (its scope sheet, else the referencing formula's). Cycle guards
 *     and memos key on the node, never on the name's text: a sheet-local `X` defined as `T!X`
 *     reaches T's own `X`, a different node, not itself.
 *   - A chain resolves at most [[NameWalk.MaxDepth]] names deep. The evaluator refuses the next
 *     name on any chain it follows — a per-cell error, never a thrown `StackOverflowError` — so a
 *     static walk treats a name with a longer chain as unresolvable ([[NameWalk.Reach.TooDeep]]),
 *     and the dependency graph adds no edge past the cap.
 *   - [[NameWalk.Walk]] answers what a name reaches with a loop, never recursion, each node once
 *     however many names ask: linear on a diamond (each name referencing the next twice) and on a
 *     chain of any length.
 */
private[formula] object NameWalk:

  /**
   * The most names one resolution path follows; the 101st is unresolvable. `NameChanges` and
   * `ReferenceScan`, whose walks answer conservatively past it, share the bound.
   */
  val MaxDepth: Int = 100

  /**
   * A resolved defined name: its definition, and the sheet its body resolves from. The hash is
   * computed once: every walk keys several maps on nodes, and a definition hashes all its fields.
   */
  final case class Node(definition: DefinedName, sheet: SheetName) derives CanEqual:
    override val hashCode: Int = scala.util.hashing.MurmurHash3.caseClassHash(this)

  /**
   * `(qualifier, name, current)` → the node a reference resolves to, from a formula (or a
   * definition's body) on `current`; `None` for a name the lookup does not find.
   */
  type Resolver = (Option[SheetName], String, SheetName) => Option[Node]

  /**
   * The scoped lookup every walk uses ([[Evaluator.lookupDefinedNameAt]]): from the qualifier's
   * sheet, or `current`'s for an unqualified name, a sheet-scoped name shadowing a workbook-scoped
   * one. `canonical` folds a qualifier onto the sheet it names (the dependency graph's
   * case-insensitive rule); a qualifier naming no sheet sees only workbook-scoped names.
   */
  def resolver(wb: Workbook, canonical: SheetName => SheetName = identity): Resolver =
    // Reverse insertion so a duplicated sheet name keeps its FIRST position, like indexWhere.
    val positions: Map[SheetName, Int] =
      wb.sheets.zipWithIndex.reverseIterator.map((sheet, i) => sheet.name -> i).toMap
    (qualifier, name, current) =>
      Evaluator
        .lookupDefinedNameAt(wb, positions.get(qualifier.fold(current)(canonical)), name)
        .map(dn => Node(dn, Evaluator.definedNameScope(wb, dn).fold(current)(_.name)))

  /**
   * One node's own verdict and the nodes its body references. A walk's nodes are usually [[Node]]s;
   * one whose verdict depends on how a name was looked up (`NameChanges`, comparing two workbooks'
   * bindings) walks the lookups instead.
   *
   * @param hit
   *   the node's body satisfies the predicate itself, without following a name
   * @param next
   *   the nodes its name references resolve to
   */
  final case class Step[K](hit: Boolean, next: Vector[K])

  object Step:
    /** An unparseable definition: no verdict of its own, no references followed. */
    def empty[K]: Step[K] = Step(hit = false, next = Vector.empty)

    /**
     * The step of `expr`, a definition's body resolving from `sheet`. `direct(expr, viaName)` is
     * the walker's predicate over one expression, with `viaName(qualifier, name)` answering for a
     * name reference: here it records the reference and answers false, so `hit` is the body's own
     * verdict and `next` its references (incomplete only when `hit` already holds).
     */
    def of(
      expr: TExpr[?],
      sheet: SheetName,
      resolve: Resolver,
      direct: (TExpr[?], (Option[SheetName], String) => Boolean) => Boolean
    ): Step[Node] =
      val refs = Vector.newBuilder[Node]
      val hit = direct(
        expr,
        (qualifier, name) =>
          refs ++= resolve(qualifier, name, sheet)
          false
      )
      val next = refs.result()
      Step(hit, if next.sizeIs > 1 then next.distinct else next)

  /** What a walk asks of a node: its verdict over every chain of names from it. */
  enum Reach derives CanEqual:
    /** A node the root reaches satisfies the predicate. */
    case Found

    /** None does, and no chain from the root is too deep to resolve. */
    case Absent

    /** A chain from the root is more than [[MaxDepth]] names long: unresolvable, so unknown. */
    case TooDeep

  /**
   * What a node reaches over every chain of names from it.
   *
   * @param found
   *   the node or a node it reaches satisfies the walk's predicate (a [[Step]] hit)
   * @param cyclic
   *   the node reaches a cycle of names (a name reaching itself)
   * @param longest
   *   the most names on a chain from the node, itself included — exact where the names form no
   *   cycle, an upper bound through one (a cycle counts every member once); capped at MaxDepth + 1
   */
  final case class Verdict(found: Boolean, cyclic: Boolean, longest: Int):
    /**
     * A chain from the node is more than [[MaxDepth]] names long: the evaluator, which refuses a
     * name past the cap on any chain it follows, may not resolve it.
     */
    def tooDeep: Boolean = longest > MaxDepth

    /** The walk's answer: too deep before found, so a chain past the cap is never resolved. */
    def reach: Reach =
      if tooDeep then Reach.TooDeep else if found then Reach.Found else Reach.Absent

  /**
   * Every node's [[Verdict]], each computed once however many roots ask: an iterative Tarjan over
   * the strongly connected components of the nodes a root reaches, the components finishing
   * dependency-first so each takes its verdict from its members' steps and the components they
   * reach. Linear in the nodes and references stepped, on a constant stack, for a chain of any
   * length, a diamond of any depth and any cycle. Not thread-safe; local to one walk.
   */
  final class Walk[K](step: K => Step[K]):
    private val verdicts = scala.collection.mutable.HashMap.empty[K, Verdict]

    def apply(root: K): Verdict =
      verdicts.get(root) match
        case Some(known) => known
        case None =>
          explore(root)
          verdicts.getOrElse(root, Verdict(found = false, cyclic = false, longest = 1))

    /** One node on the depth-first path: its successors left, and what it has gathered so far. */
    @SuppressWarnings(Array("org.wartremover.warts.Var"))
    private final class Frame(val node: K, val index: Int, val children: Iterator[K], hit: Boolean):
      var low: Int = index
      var found: Boolean = hit
      var cyclic: Boolean = false
      var longest: Int = 0 // the longest verdict among successors in finished components

      def absorb(verdict: Verdict): Unit =
        found ||= verdict.found
        cyclic ||= verdict.cyclic
        longest = math.max(longest, verdict.longest)

    @SuppressWarnings(Array("org.wartremover.warts.Var"))
    private def explore(root: K): Unit =
      val path = scala.collection.mutable.ArrayBuffer.empty[Frame] // the depth-first call stack
      val component = scala.collection.mutable.ArrayBuffer.empty[Frame] // Tarjan's stack
      val open = scala.collection.mutable.HashMap.empty[K, Frame] // nodes on Tarjan's stack
      var counter = 0
      def enter(node: K): Unit =
        val own = step(node)
        val frame = Frame(node, counter, own.next.iterator, own.hit)
        counter += 1
        path += frame
        component += frame
        open(node) = frame
      // A frame with no successor left finishes: the root of a component pops it and gives every
      // member the component's verdict; any other hands its low-link to its parent.
      def finish(frame: Frame): Unit =
        path.dropRightInPlace(1)
        if frame.low == frame.index then
          val start = component.lastIndexWhere(_ eq frame)
          val size = component.length - start
          var found = false
          var cyclic = false
          var beyond = 0
          component.view.drop(start).foreach { member =>
            found ||= member.found
            cyclic ||= member.cyclic
            beyond = math.max(beyond, member.longest)
          }
          val verdict = Verdict(found, cyclic, math.min(MaxDepth + 1, size + beyond))
          component.view.drop(start).foreach { member =>
            open.remove(member.node)
            verdicts(member.node) = verdict
          }
          component.dropRightInPlace(size)
          path.lastOption.foreach(_.absorb(verdict))
        else path.lastOption.foreach(parent => parent.low = math.min(parent.low, frame.low))
      @tailrec
      def loop(): Unit =
        path.lastOption match
          case None => ()
          case Some(frame) =>
            if frame.children.hasNext then
              val child = frame.children.next()
              verdicts.get(child) match
                case Some(verdict) => frame.absorb(verdict)
                case None =>
                  open.get(child) match
                    // a reference back into the component being built: a cycle (itself included)
                    case Some(onStack) =>
                      frame.low = math.min(frame.low, onStack.index)
                      frame.cyclic = true
                    case None => enter(child)
            else finish(frame)
            loop()
      enter(root)
      loop()
