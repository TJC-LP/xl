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
 * Keys are the per-character `Character.toLowerCase(Character.toUpperCase(c))` form of the name.
 * `String.equalsIgnoreCase` specifies per character: the characters are equal, or applying
 * `toUpperCase` then `toLowerCase` to each yields the same character — so two names compare equal
 * under `equalsIgnoreCase` iff they have equal length and equal keys. Key equality is therefore the
 * whole matching rule; lookup performs no further string comparison. First-declared-wins is baked
 * in at build time, so lookup is one map probe per scope.
 *
 * see:
 * https://docs.oracle.com/en/java/javase/21/docs/api/java.base/java/lang/String.html#equalsIgnoreCase(java.lang.String)
 */
private[xl] final class DefinedNameIndex private (
  globalByKey: Map[String, DefinedName],
  sheetScopedByKey: Map[(String, Int), DefinedName]
) extends Serializable:

  /**
   * The first declared name matching `name` case-insensitively with scope
   * `localSheetId == Some(sheetIdx)`.
   */
  def sheetScoped(name: String, sheetIdx: Int): Option[DefinedName] =
    sheetScopedByKey.get((DefinedNameIndex.caseKey(name), sheetIdx))

  /**
   * The first declared workbook-scoped (`localSheetId.isEmpty`) name matching `name`
   * case-insensitively.
   */
  def workbookScoped(name: String): Option[DefinedName] =
    globalByKey.get(DefinedNameIndex.caseKey(name))

  /**
   * Excel/OOXML visibility from a sheet position: the sheet-scoped entry shadows a workbook-scoped
   * entry of the same identifier (OOXML §18.2.5); with no sheet position only workbook-scoped
   * entries are visible.
   */
  def resolve(name: String, sheetIdx: Option[Int]): Option[DefinedName] =
    val key = DefinedNameIndex.caseKey(name)
    sheetIdx
      .flatMap(idx => sheetScopedByKey.get((key, idx)))
      .orElse(globalByKey.get(key))

private[xl] object DefinedNameIndex:

  /** Build from a name table, keeping the first declared entry per (key, scope). */
  def apply(names: Vector[DefinedName]): DefinedNameIndex =
    val global = Map.newBuilder[String, DefinedName]
    val scoped = Map.newBuilder[(String, Int), DefinedName]
    val seenGlobal = scala.collection.mutable.HashSet.empty[String]
    val seenScoped = scala.collection.mutable.HashSet.empty[(String, Int)]
    names.foreach { dn =>
      val key = caseKey(dn.name)
      dn.localSheetId match
        case None => if seenGlobal.add(key) then global += key -> dn
        case Some(idx) =>
          val scopedKey = (key, idx)
          if seenScoped.add(scopedKey) then scoped += scopedKey -> dn
    }
    new DefinedNameIndex(global.result(), scoped.result())

  /** The per-character `toLowerCase(toUpperCase(c))` mapping from the `equalsIgnoreCase` spec. */
  private def caseKey(name: String): String =
    name.map(c => Character.toLowerCase(Character.toUpperCase(c)))
