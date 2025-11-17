package com.tjclp.xl.ooxml

/**
 * Dependency graph for OOXML relationships.
 *
 * Tracks which ZIP entries depend on which sheets, enabling surgical modification to determine
 * which parts can be preserved byte-for-byte vs. which must be regenerated.
 *
 * Example: If xl/drawings/drawing1.xml references Sheet1, and Sheet1 is modified, the drawing must
 * be regenerated or dropped (XL doesn't yet parse drawings, so we drop).
 *
 * PURITY: This is a pure data structure with no IO. All graph construction is deterministic.
 */
final case class RelationshipGraph(
  dependencies: Map[String, Set[Int]], // part path -> sheet indices it depends on
  sheetPaths: Map[Int, String] // sheet index -> original worksheet path
) derives CanEqual:

  /**
   * Get the set of sheet indices that a given part depends on.
   *
   * @param path
   *   ZIP entry path (e.g., "xl/drawings/drawing1.xml")
   * @return
   *   Set of sheet indices this part references (empty if no dependencies)
   */
  def dependenciesFor(path: String): Set[Int] =
    dependencies.getOrElse(path, Set.empty)

  /**
   * Get the original worksheet path for a sheet index.
   *
   * Used during hybrid writes to locate unmodified sheets in the source ZIP.
   *
   * @param idx
   *   Zero-based sheet index
   * @return
   *   Path to worksheet XML (e.g., "xl/worksheets/sheet1.xml")
   */
  def pathForSheet(idx: Int): String =
    sheetPaths.getOrElse(idx, s"xl/worksheets/sheet${idx + 1}.xml")

object RelationshipGraph:

  /** Empty graph (no dependencies) */
  val empty: RelationshipGraph = RelationshipGraph(Map.empty, Map.empty)

  /**
   * Build dependency graph from part manifest.
   *
   * Extracts:
   *   1. Sheet paths from manifest entries with sheetIndex
   *   1. Dependencies from relationship metadata (which parts reference which sheets)
   *
   * SIMPLIFICATION (Phase 4): We use manifest's sheetIndex hints (populated by reader). In future
   * phases, we'll parse _rels files to extract precise rId → target mappings.
   *
   * @param manifest
   *   Complete manifest of all ZIP entries
   * @return
   *   Dependency graph mapping parts to sheets
   */
  def fromManifest(manifest: PartManifest): RelationshipGraph =
    // Extract sheet paths: sheet index -> worksheet path
    val sheetPaths = manifest.entries.flatMap { case (path, entry) =>
      entry.sheetIndex.map(idx => idx -> path)
    }

    // Extract dependencies: each part maps to the sheets it references
    // Currently, we use the sheetIndex hint from the manifest (if present)
    // Future: Parse _rels files to get precise rId -> target mappings
    val dependencies = manifest.entries.map { case (path, entry) =>
      val deps = entry.sheetIndex.map(Set(_)).getOrElse(Set.empty)
      path -> deps
    }

    RelationshipGraph(dependencies, sheetPaths)
