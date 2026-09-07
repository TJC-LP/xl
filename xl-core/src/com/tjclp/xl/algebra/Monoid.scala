package com.tjclp.xl.algebra

/**
 * An associative binary operation with an identity element.
 *
 * xl's own type class, so the pure core carries no third-party algebra dependency. Instances live
 * in the companions of the types they describe ([[com.tjclp.xl.patch.Patch]],
 * [[com.tjclp.xl.styles.patch.StylePatch]]) and are law-tested there:
 *
 *   - Associativity: `combine(combine(a, b), c) == combine(a, combine(b, c))`
 *   - Left identity: `combine(empty, a) == a`
 *   - Right identity: `combine(a, empty) == a`
 *
 * Cats users: `com.tjclp.xl.interop.CatsInstances` (module `xl-cats-effect`) derives a
 * `cats.Monoid[A]` from any `Monoid[A]` in scope.
 */
trait Monoid[A]:
  /** The identity element. */
  def empty: A

  /** Associative combination of two values. */
  def combine(x: A, y: A): A

  /** Fold every element with [[combine]], starting from [[empty]]. */
  def combineAll(as: IterableOnce[A]): A = as.iterator.foldLeft(empty)(combine)

object Monoid:
  /** Summon the instance for `A`. */
  def apply[A](using m: Monoid[A]): Monoid[A] = m

  /** Build an instance from an identity element and a combine function. */
  def instance[A](identity: A)(f: (A, A) => A): Monoid[A] = new Monoid[A]:
    def empty: A = identity
    def combine(x: A, y: A): A = f(x, y)

/** Operator syntax for [[Monoid]]. */
object syntax:
  extension [A](x: A)
    /**
     * Combine two values through the [[Monoid]] of their common supertype.
     *
     * The result type `B` is inferred as the least upper bound of both operands, so enum cases
     * compose without ascription: `Patch.Put(ref, v) |+| Patch.SetStyle(ref, id)` resolves
     * `Monoid[Patch]`. (`Patch` and `StylePatch` also offer `++`, fixed to their own type.)
     */
    infix def |+|[B >: A](y: B)(using m: Monoid[B]): B = m.combine(x, y)
