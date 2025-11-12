package com.tjclp.xl.optics

/** Lens: Total getter and setter for a field */
final case class Lens[S, A](get: S => A, set: (A, S) => S):
  /** Modify the focused field using a function */
  def modify(f: A => A)(s: S): S = set(f(get(s)), s)

  /** Update the focused field using a function */
  def update(f: A => A)(s: S): S = set(f(get(s)), s)

  /** Compose with another lens */
  def andThen[B](other: Lens[A, B]): Lens[S, B] =
    Lens(
      get = s => other.get(this.get(s)),
      set = (b, s) => this.set(other.set(b, this.get(s)), s)
    )
