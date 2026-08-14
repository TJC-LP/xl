package com.tjclp.xl.workbooks

/**
 * Case-insensitive lookup index over a workbook's defined-name table.
 *
 * Resolution returns exactly what a declaration-order linear scan of
 * [[WorkbookMetadata.definedNames]] with `String.equalsIgnoreCase` returns: the first declared
 * match, filtered by scope. The index exists because name resolution sits on the evaluator's
 * per-reference hot path and real bank-authored workbooks carry name tables of 10^5 entries, which
 * made recalculation O(sheets × names²) through linear scans.
 *
 * Buckets key on the per-character `toLowerCase(toUpperCase(c))` form of the name — the same
 * two-step mapping `equalsIgnoreCase` applies per character — so two names equal under
 * `equalsIgnoreCase` always share a bucket, and candidates need only the final `equalsIgnoreCase`
 * confirmation.
 *
 * see:
 * https://docs.oracle.com/en/java/javase/21/docs/api/java.base/java/lang/String.html#equalsIgnoreCase(java.lang.String)
 */
final class DefinedNameIndex private (buckets: Map[String, Vector[DefinedName]]):

  /**
   * The first declared name matching `name` case-insensitively with scope
   * `localSheetId == Some(sheetIdx)`.
   */
  def sheetScoped(name: String, sheetIdx: Int): Option[DefinedName] =
    candidates(name).find(dn =>
      dn.name.equalsIgnoreCase(name) && dn.localSheetId.contains(sheetIdx)
    )

  /**
   * The first declared workbook-scoped (`localSheetId.isEmpty`) name matching `name`
   * case-insensitively.
   */
  def workbookScoped(name: String): Option[DefinedName] =
    candidates(name).find(dn => dn.name.equalsIgnoreCase(name) && dn.localSheetId.isEmpty)

  /**
   * Excel/OOXML visibility from a sheet position: the sheet-scoped entry shadows a workbook-scoped
   * entry of the same identifier (OOXML §18.2.5); with no sheet position only workbook-scoped
   * entries are visible.
   */
  def resolve(name: String, sheetIdx: Option[Int]): Option[DefinedName] =
    sheetIdx.flatMap(sheetScoped(name, _)).orElse(workbookScoped(name))

  private def candidates(name: String): Vector[DefinedName] =
    buckets.getOrElse(DefinedNameIndex.caseKey(name), Vector.empty)

object DefinedNameIndex:

  /** Build from a name table; `Vector.groupBy` keeps declaration order inside each bucket. */
  def apply(names: Vector[DefinedName]): DefinedNameIndex =
    new DefinedNameIndex(names.groupBy(dn => caseKey(dn.name)))

  /** The per-character two-step case mapping `String.equalsIgnoreCase` compares by. */
  private def caseKey(name: String): String =
    name.map(c => Character.toLowerCase(Character.toUpperCase(c)))
