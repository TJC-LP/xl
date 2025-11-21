package com.tjclp.xl.formula

import java.time.{LocalDate, LocalDateTime}

/**
 * Clock trait for pure date/time access in formula evaluation.
 *
 * This trait enables deterministic evaluation of TODAY() and NOW() functions by explicitly passing
 * time as a parameter rather than accessing system time directly. This preserves purity and enables
 * testability.
 *
 * Design follows dependency injection pattern: evaluator takes Clock parameter, tests provide fixed
 * clock, production provides system clock.
 *
 * Laws satisfied:
 *   1. Determinism: Fixed clock returns same values across multiple calls
 *   2. Monotonicity: System clock values advance (not guaranteed but expected)
 *   3. Consistency: today().atStartOfDay() <= now()
 *
 * Example:
 * {{{
 * // Test with fixed clock
 * val testClock = Clock.fixed(LocalDate.of(2025, 11, 21), LocalDateTime.of(2025, 11, 21, 18, 30))
 * evaluator.eval(TExpr.today(), sheet, testClock) == Right(LocalDate.of(2025, 11, 21))
 *
 * // Production with system clock
 * val sysClock = Clock.system
 * evaluator.eval(TExpr.now(), sheet, sysClock) // Current system time
 * }}}
 */
trait Clock:
  /**
   * Get current date without time component.
   *
   * Used by TODAY() function.
   *
   * @return
   *   Current date
   */
  def today(): LocalDate

  /**
   * Get current date and time.
   *
   * Used by NOW() function.
   *
   * @return
   *   Current date and time
   */
  def now(): LocalDateTime

object Clock:
  /**
   * System clock that reads actual system time.
   *
   * Use in production for real-time formula evaluation.
   *
   * Example:
   * {{{
   * val evaluator = Evaluator.instance
   * evaluator.eval(TExpr.today(), sheet, Clock.system)
   * }}}
   */
  def system: Clock = new Clock:
    def today(): LocalDate = LocalDate.now()
    def now(): LocalDateTime = LocalDateTime.now()

  /**
   * Fixed clock that returns constant values.
   *
   * Use in tests for deterministic evaluation.
   *
   * @param fixedDate
   *   The date to return from today()
   * @param fixedDateTime
   *   The date and time to return from now()
   *
   * Example:
   * {{{
   * val testClock = Clock.fixed(
   *   LocalDate.of(2025, 11, 21),
   *   LocalDateTime.of(2025, 11, 21, 18, 30, 0)
   * )
   * evaluator.eval(TExpr.today(), sheet, testClock) == Right(LocalDate.of(2025, 11, 21))
   * }}}
   */
  def fixed(fixedDate: LocalDate, fixedDateTime: LocalDateTime): Clock = new Clock:
    def today(): LocalDate = fixedDate
    def now(): LocalDateTime = fixedDateTime

  /**
   * Fixed clock with only date specified (time defaults to midnight).
   *
   * Convenience constructor for tests that only care about date.
   *
   * @param fixedDate
   *   The date to return from today() and use for now()
   *
   * Example:
   * {{{
   * val testClock = Clock.fixedDate(LocalDate.of(2025, 11, 21))
   * evaluator.eval(TExpr.today(), sheet, testClock) == Right(LocalDate.of(2025, 11, 21))
   * evaluator.eval(TExpr.now(), sheet, testClock) == Right(LocalDateTime.of(2025, 11, 21, 0, 0))
   * }}}
   */
  def fixedDate(fixedDate: LocalDate): Clock =
    fixed(fixedDate, fixedDate.atStartOfDay())
