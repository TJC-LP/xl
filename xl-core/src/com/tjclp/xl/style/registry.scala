package com.tjclp.xl.style

/**
 * Registry for tracking CellStyle → StyleId mappings within a Sheet.
 *
 * Coordinates style application by deduplicating styles and assigning stable indices.
 */

// ========== StyleRegistry ==========

/**
 * Registry for tracking CellStyle → StyleId mappings within a Sheet.
 *
 * Coordinates style application by deduplicating styles and assigning stable indices. Styles are
 * identified by their canonical key for structural equality.
 *
 * Invariants:
 *   - Index 0 is always CellStyle.default
 *   - Styles are deduplicated by canonicalKey
 *   - Indices are stable (same style always gets same index within a registry)
 */
case class StyleRegistry(
  styles: Vector[CellStyle] = Vector(CellStyle.default),
  index: Map[String, StyleId] = Map(CellStyle.canonicalKey(CellStyle.default) -> StyleId(0))
):
  /**
   * Register a style and get its index.
   *
   * If the style already exists (by canonical key), returns existing index. Otherwise, appends to
   * styles vector and returns new index.
   *
   * @param style
   *   The CellStyle to register
   * @return
   *   (updated registry, style index)
   */
  def register(style: CellStyle): (StyleRegistry, StyleId) =
    val key = CellStyle.canonicalKey(style)
    index.get(key) match
      case Some(idx) =>
        // Already registered
        (this, idx)
      case None =>
        // New style - append and index
        val idx = StyleId(styles.size)
        val updated = copy(
          styles = styles :+ style,
          index = index + (key -> idx)
        )
        (updated, idx)

  /**
   * Look up index for a style by canonical key.
   *
   * @param style
   *   The CellStyle to look up
   * @return
   *   Some(index) if registered, None otherwise
   */
  def indexOf(style: CellStyle): Option[StyleId] =
    index.get(CellStyle.canonicalKey(style))

  /**
   * Retrieve style by index.
   *
   * @param idx
   *   The style index
   * @return
   *   Some(style) if index is valid, None otherwise
   */
  def get(idx: StyleId): Option[CellStyle] =
    styles.lift(idx.value)

  /** Number of registered styles (including default) */
  def size: Int = styles.size

  /** Check if registry contains only default style */
  def isEmpty: Boolean = styles.size == 1 && styles.head == CellStyle.default

object StyleRegistry:
  /** Default registry with only CellStyle.default at index 0 */
  def default: StyleRegistry = StyleRegistry()
