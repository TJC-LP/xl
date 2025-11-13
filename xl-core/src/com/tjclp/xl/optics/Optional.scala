package com.tjclp.xl.optics

/** Optional: Partial getter and setter (may not have a value) */
final case class Optional[S, A](getOption: S => Option[A], set: (A, S) => S):
  /** Modify the focused field if it exists */
  def modify(f: A => A)(s: S): S =
    getOption(s) match
      case Some(a) => set(f(a), s)
      case None => s

  /** Get the value or use a default */
  def getOrElse(default: A)(s: S): A =
    getOption(s).getOrElse(default)
