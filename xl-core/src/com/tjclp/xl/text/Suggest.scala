package com.tjclp.xl.text

import java.util.Locale

/**
 * "Did you mean" ranking for user-supplied names (sheet names, verbs, batch op names).
 *
 * Pure and deterministic: candidates are ranked by case-insensitive Levenshtein distance, ties
 * broken by the candidate's spelling, so the same inputs always yield the same suggestions.
 */
object Suggest:

  /**
   * The candidates within edit distance `max(2, input.length / 3)` of `input`, nearest first (ties
   * ordered by candidate), at most `limit` of them. Comparison ignores case; the result carries
   * each candidate as spelled. Duplicate candidates are reported once.
   */
  def closest(input: String, candidates: Iterable[String], limit: Int = 3): Vector[String] =
    val needle = input.toLowerCase(Locale.ROOT)
    val threshold = math.max(2, input.length / 3)
    candidates.toVector.distinct
      .map(candidate => (distance(needle, candidate.toLowerCase(Locale.ROOT)), candidate))
      .filter((d, _) => d <= threshold)
      .sorted
      .map((_, candidate) => candidate)
      .take(math.max(0, limit))

  /** Exact (case-sensitive) Levenshtein distance; two-row dynamic programme, no mutation. */
  private[xl] def distance(a: String, b: String): Int =
    if a.isEmpty then b.length
    else if b.isEmpty then a.length
    else
      val firstRow = Vector.tabulate(b.length + 1)(identity)
      val lastRow = a.zipWithIndex.foldLeft(firstRow) { case (previous, (ca, i)) =>
        b.zipWithIndex.foldLeft(Vector(i + 1)) { case (row, (cb, j)) =>
          val substitution = previous(j) + (if ca == cb then 0 else 1)
          val deletion = previous(j + 1) + 1
          val insertion = row(j) + 1
          row :+ math.min(substitution, math.min(deletion, insertion))
        }
      }
      lastRow(b.length)
