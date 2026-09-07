package com.tjclp.xl.formula.eval

import com.tjclp.xl.formula.{Clock, Rng}

/**
 * How a recalculation treats circular references (ADR-017 §2.8).
 *
 *   - `Off`: cycles are isolated and reported as errors — the default `recalculate()` posture.
 *   - `FromCalcPr`: honour the workbook's own `<calcPr iterate="1"/>` (Excel's defaults 100 / 0.001
 *     for absent attributes, via [[IterativeCalc.fromCalcPr]]) and behave like `Off` when the book
 *     declares nothing — what `xl recalc` and every recalculating CLI verb do today.
 *   - `Force(calc)`: iterate with explicit settings regardless of what the file declares.
 */
enum IterativeMode derives CanEqual:
  case Off
  case FromCalcPr
  case Force(calc: IterativeCalc)

/**
 * The ONE options record behind `wb.recalculate(options)`, `wb.recalculateAfterEdit(sheet, refs,
 * options)` and `wb.recalculateUncached(options)` (ADR-017 §2.8). Every field has the value the
 * corresponding zero-argument overload uses, so `RecalcOptions()` reproduces `recalculate()` on an
 * acyclic book exactly; the record exists so that extension methods reached through the wildcard
 * export need no default arguments (the compiler landmine documented in exports.scala).
 *
 * @param clock
 *   Volatile time source (TODAY/NOW); pinned once per calculation generation
 * @param rng
 *   Randomness source (RAND/RANDBETWEEN); `Rng.seeded(n)` makes a sequential run reproducible.
 *   Ignored by the parallel path, which uses the thread-safe system generator (see
 *   `recalculateParallel`)
 * @param iterative
 *   Circular-reference posture, see [[IterativeMode]]; a declared or forced iteration wins over
 *   `parallelism` (cyclic components fixpoint sequentially), exactly as the CLI's calcPr rule does
 * @param parallelism
 *   Worker threads for independent formula regions when > 1 (GH-520); `1` is the sequential fold
 * @param seedTables
 *   Also seed every data-table interior after the recalculation (`seedDataTablesReport` with the
 *   same clock and iterative settings; GH-419/GH-453) — off by default because a seeding run is a
 *   what-if replay, not a recalculation. Per-table warnings are not carried by the
 *   [[RecalcResult]]; call `seedDataTablesReport` directly when they matter
 */
final case class RecalcOptions(
  clock: Clock = Clock.system,
  rng: Rng = Rng.system,
  iterative: IterativeMode = IterativeMode.FromCalcPr,
  parallelism: Int = 1,
  seedTables: Boolean = false
) derives CanEqual

object RecalcOptions:
  /** `RecalcOptions()` — the settings behind the zero-argument `recalculate()`. */
  val default: RecalcOptions = RecalcOptions()
