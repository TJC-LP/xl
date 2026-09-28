package com.tjclp.xl.formula.ast

import scala.annotation.tailrec

import TExpr.*

/**
 * GH-680: the left spine of a binary-operator chain, walked without recursing along it.
 *
 * The parser folds `a+b+c+…` (and every other binary operator) to the LEFT, so a flat chain of n
 * operators is a spine n nodes deep: `Add(Add(Add(a, b), c), …)`. A walker that recurses into the
 * left operand spends a stack frame per operator, and a chain as long as the parser admits would
 * exhaust a default-size thread stack under interpreted execution. The heavy walkers (evaluator,
 * printer, shifter, dependency extraction) instead take the whole spine in one loop: [[unwind]]
 * finds the innermost left operand that is not itself a binary node plus the spine's nodes
 * bottom-up, the walker handles that operand and every right operand (whose own depth is bounded by
 * the nesting budget), and [[mapOperands]] rebuilds bottom-up in a fold.
 */
@SuppressWarnings(Array("org.wartremover.warts.AsInstanceOf"))
private[formula] object BinarySpine:

  /** The operands of a binary-operator node: arithmetic, `&`, and the comparisons. */
  def operands(e: TExpr[?]): Option[(TExpr[?], TExpr[?])] = e match
    case Add(x, y) => Some((x, y))
    case Sub(x, y) => Some((x, y))
    case Mul(x, y) => Some((x, y))
    case Div(x, y) => Some((x, y))
    case Pow(x, y) => Some((x, y))
    case Concat(x, y) => Some((x, y))
    case Eq(x, y) => Some((x, y))
    case Neq(x, y) => Some((x, y))
    case Lt(x, y) => Some((x, y))
    case Lte(x, y) => Some((x, y))
    case Gt(x, y) => Some((x, y))
    case Gte(x, y) => Some((x, y))
    case _ => None

  def isBinary(e: TExpr[?]): Boolean = operands(e).isDefined

  /** True when `e` is a binary node whose left operand is one too — a spine worth unwinding. */
  def isChained(e: TExpr[?]): Boolean = operands(e).exists((x, _) => isBinary(x))

  /**
   * The same kind of node as `node` over new operands. `node` must be binary ([[isBinary]]); any
   * other node comes back unchanged.
   */
  def rebuild(node: TExpr[?], x: TExpr[?], y: TExpr[?]): TExpr[?] = node match
    case _: Add => Add(x.asInstanceOf[TExpr[BigDecimal]], y.asInstanceOf[TExpr[BigDecimal]])
    case _: Sub => Sub(x.asInstanceOf[TExpr[BigDecimal]], y.asInstanceOf[TExpr[BigDecimal]])
    case _: Mul => Mul(x.asInstanceOf[TExpr[BigDecimal]], y.asInstanceOf[TExpr[BigDecimal]])
    case _: Div => Div(x.asInstanceOf[TExpr[BigDecimal]], y.asInstanceOf[TExpr[BigDecimal]])
    case _: Pow => Pow(x.asInstanceOf[TExpr[BigDecimal]], y.asInstanceOf[TExpr[BigDecimal]])
    case _: Concat => Concat(x.asInstanceOf[TExpr[String]], y.asInstanceOf[TExpr[String]])
    case _: Eq[?] => Eq(x.asInstanceOf[TExpr[Any]], y.asInstanceOf[TExpr[Any]])
    case _: Neq[?] => Neq(x.asInstanceOf[TExpr[Any]], y.asInstanceOf[TExpr[Any]])
    case _: Lt[?] => Lt(x.asInstanceOf[TExpr[Any]], y.asInstanceOf[TExpr[Any]])
    case _: Lte[?] => Lte(x.asInstanceOf[TExpr[Any]], y.asInstanceOf[TExpr[Any]])
    case _: Gt[?] => Gt(x.asInstanceOf[TExpr[Any]], y.asInstanceOf[TExpr[Any]])
    case _: Gte[?] => Gte(x.asInstanceOf[TExpr[Any]], y.asInstanceOf[TExpr[Any]])
    case other => other

  /**
   * Unwind the left spine under `follow`: the innermost left operand the spine does not continue
   * through, and the spine's nodes bottom-up (the innermost first, `e` itself last). A node that is
   * not binary, or that `follow` refuses, is a spine of none.
   */
  def unwind(e: TExpr[?], follow: TExpr[?] => Boolean): (TExpr[?], List[TExpr[?]]) =
    @tailrec
    def loop(current: TExpr[?], above: List[TExpr[?]]): (TExpr[?], List[TExpr[?]]) =
      operands(current) match
        case Some((x, _)) if follow(current) => loop(x, current :: above)
        case _ => (current, above)
    loop(e, Nil)

  /** The right operand of a binary node (the node itself for any other). */
  def right(node: TExpr[?]): TExpr[?] = operands(node).fold(node)(_._2)

  /**
   * Rebuild `e` with `f` applied to every operand of its left spine — the innermost non-binary left
   * operand and each right operand — in one bottom-up fold. `f` never sees a spine node, so a
   * recursive walker calls this for a binary node and recurses only into operands.
   */
  def mapOperands(e: TExpr[?])(f: TExpr[?] => TExpr[?]): TExpr[?] =
    val (leftmost, spine) = unwind(e, isBinary)
    spine.foldLeft(f(leftmost))((acc, node) => rebuild(node, acc, f(right(node))))

  /**
   * True when `p` holds for some operand of `e`'s left spine (see [[mapOperands]]); `follow`
   * narrows which binary nodes the spine continues through (an operand it refuses is handed to `p`
   * whole).
   */
  def existsOperand(e: TExpr[?], follow: TExpr[?] => Boolean = isBinary)(
    p: TExpr[?] => Boolean
  ): Boolean =
    val (leftmost, spine) = unwind(e, follow)
    p(leftmost) || spine.exists(node => p(right(node)))

  /** Every operand of `e`'s left spine, leftmost first (see [[mapOperands]]). */
  def operandList(e: TExpr[?]): List[TExpr[?]] =
    val (leftmost, spine) = unwind(e, isBinary)
    leftmost :: spine.map(right)
