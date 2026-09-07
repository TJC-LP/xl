package com.tjclp.xl.interop

import com.tjclp.xl.algebra.Monoid

/**
 * Cats instances derived from xl's own algebra.
 *
 * The pure core defines [[com.tjclp.xl.algebra.Monoid]] itself and depends on no Cats artifact.
 * Code that composes patches with Cats syntax imports this object's givens:
 *
 * {{{
 * import cats.syntax.all.*
 * import com.tjclp.xl.interop.CatsInstances.given
 *
 * val patch = (Patch.Put(ref, v): Patch) |+| (Patch.SetStyle(ref, id): Patch)
 * Monoid[Patch].combineAll(patches)   // cats.Monoid
 * }}}
 */
object CatsInstances:

  /** Every xl `Monoid[A]` is a lawful `cats.Monoid[A]` with the same identity and combine. */
  given catsMonoidFromXl[A](using m: Monoid[A]): cats.Monoid[A] =
    cats.Monoid.instance(m.empty, m.combine)
