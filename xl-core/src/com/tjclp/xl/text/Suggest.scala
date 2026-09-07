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
   * The candidates within edit distance [[threshold]] of `input`, nearest first (ties ordered by
   * candidate), at most `limit` of them. Comparison ignores case; the result carries each candidate
   * as spelled. Duplicate candidates are reported once.
   */
  def closest(input: String, candidates: Iterable[String], limit: Int = 3): Vector[String] =
    val needle = input.toLowerCase(Locale.ROOT)
    val max = threshold(input)
    candidates.toVector.distinct
      .map(candidate => (distance(needle, candidate.toLowerCase(Locale.ROOT)), candidate))
      .filter((d, _) => d <= max)
      .sorted
      .map((_, candidate) => candidate)
      .take(math.max(0, limit))

  /**
   * The largest edit distance still read as a near miss: one edit for an input of three characters
   * or fewer (`pu` suggests `put`, not `cf` or `putf`), else `max(2, input.length / 3)` — and never
   * the input's own length, since rewriting every character is not a near miss.
   */
  private[xl] def threshold(input: String): Int =
    val byLength = if input.length <= 3 then 1 else math.max(2, input.length / 3)
    math.min(byLength, input.length - 1)

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
