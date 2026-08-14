package com.tjclp.xl.workbooks

/**
 * Case-insensitive lookup index over a workbook's defined-name table.
 *
 * Resolution returns exactly what a declaration-order linear scan of
 * [[WorkbookMetadata.definedNames]] with `String.equalsIgnoreCase` returns: the first declared
 * match, filtered by scope. The index exists because name resolution sits on the evaluator's
 * per-reference hot path and real bank-authored workbooks carry name tables of 10^5 entries, which
 * made recalculation O(sheets × names²) through linear scans (GH-536).
 *
 * Keys are the per-code-point `Character.toLowerCase(Character.toUpperCase(c))` form of the name.
 * Walking Unicode code points rather than UTF-16 code units is required for supplementary-plane
 * case pairs that `String.equalsIgnoreCase` treats as equal. Two names compare equal under
 * `equalsIgnoreCase` iff they have equal length and equal keys. Key equality is therefore the whole
 * matching rule; lookup performs no further string comparison. First-declared-wins is baked in at
 * build time, so lookup is one map probe per scope.
 *
 * Memory: one key string per name plus two maps — tens of MB at 10^5 names. The index is reached
 * only through the lazy [[WorkbookMetadata.definedNameIndex]], so paths that never resolve a name
 * (streaming reads in particular) do not pay it.
 *
 * `Serializable` because [[WorkbookMetadata]] is a serializable case class and a materialized lazy
 * val is a real field that rides along in its serialized form — without it, serializing metadata
 * whose index has been touched would throw `NotSerializableException`.
 *
 * see:
 * https://docs.oracle.com/en/java/javase/21/docs/api/java.base/java/lang/String.html#equalsIgnoreCase(java.lang.String)
 */
private[xl] final class DefinedNameIndex private (
  globalByKey: Map[String, DefinedName],
  sheetScopedByKey: Map[(String, Int), DefinedName]
) extends Serializable:

  /**
   * Excel/OOXML visibility from a sheet position: the sheet-scoped entry shadows a workbook-scoped
   * entry of the same identifier (OOXML §18.2.5); with no sheet position only workbook-scoped
   * entries are visible.
   */
  def resolve(name: String, sheetIdx: Option[Int]): Option[DefinedName] =
    val key = DefinedNameIndex.caseKey(name)
    val scoped =
      if sheetScopedByKey.isEmpty then None
      else sheetIdx.flatMap(idx => sheetScopedByKey.get((key, idx)))
    scoped.orElse(globalByKey.get(key))

private[xl] object DefinedNameIndex:

  /**
   * Build from a name table. `toMap` keeps the last entry per key, so inserting in reverse
   * declaration order makes the first declared entry win — the same answer `Vector.find` gave.
   */
  def apply(names: Vector[DefinedName]): DefinedNameIndex =
    new DefinedNameIndex(
      names.reverseIterator.collect {
        case dn if dn.localSheetId.isEmpty => caseKey(dn.name) -> dn
      }.toMap,
      names.reverseIterator
        .flatMap(dn => dn.localSheetId.map(idx => (caseKey(dn.name), idx) -> dn))
        .toMap
    )

  /**
   * The per-code-point `toLowerCase(toUpperCase(c))` mapping used by `equalsIgnoreCase`.
   * Already-canonical names (lower-case, digits, underscores) return unchanged without allocating a
   * replacement string; for conventionally capitalized names `allMatch` short-circuits at the first
   * upper-case code point and only the `map` pass runs.
   */
  private def caseKey(name: String): String =
    if name.codePoints().allMatch(c => caseFold(c) == c) then name
    else
      val folded = name.codePoints().map(c => caseFold(c)).toArray
      new String(folded, 0, folded.length)

  private def caseFold(codePoint: Int): Int =
    Character.toLowerCase(Character.toUpperCase(codePoint))
