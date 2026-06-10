package com.tjclp.xl.formula

/**
 * Rng trait for random number access in formula evaluation.
 *
 * This trait enables deterministic evaluation of RAND() and RANDBETWEEN() by explicitly passing the
 * randomness source as a parameter rather than accessing a global generator directly. This
 * preserves purity at the API boundary and enables testability — the Clock pattern, applied to
 * randomness.
 *
 * Design follows dependency injection: the evaluator takes an Rng parameter, tests provide a seeded
 * generator, production provides the system generator.
 *
 * Laws satisfied:
 *   1. Range: nextDouble() always returns a value in [0, 1)
 *   2. Determinism: Rng.seeded(s) produces the same sequence for the same seed
 *   3. Statefulness: a single Rng instance advances on each call (a sequence, not a constant)
 *
 * Example:
 * {{{
 * // Test with seeded rng (deterministic)
 * val testRng = Rng.seeded(42L)
 * sheet.evaluateFormula("=RAND()", Clock.system, Rng.seeded(42L)) // same value every run
 *
 * // Production with system rng
 * sheet.evaluateFormula("=RAND()") // varies per evaluation
 * }}}
 */
trait Rng:
  /**
   * Next uniformly distributed value in [0, 1).
   *
   * Used by RAND() directly and by RANDBETWEEN() to draw an integer in range.
   *
   * @return
   *   A double in [0, 1)
   */
  def nextDouble(): Double

object Rng:
  /**
   * System generator backed by ThreadLocalRandom — non-deterministic, thread-safe, no contention.
   *
   * Use in production where Excel-style volatile behavior (a fresh draw per evaluation) is wanted.
   */
  def system: Rng = new Rng:
    def nextDouble(): Double = java.util.concurrent.ThreadLocalRandom.current().nextDouble()

  /**
   * Seeded generator backed by scala.util.Random — deterministic for a given seed.
   *
   * Each returned instance carries its own generator state: the same seed always yields the same
   * sequence of values. Use in tests and reproducible pipelines.
   *
   * @param seed
   *   The seed determining the sequence
   */
  def seeded(seed: Long): Rng = new Rng:
    private val random = new scala.util.Random(seed)
    def nextDouble(): Double = random.nextDouble()
