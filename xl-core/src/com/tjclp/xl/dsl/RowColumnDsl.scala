package com.tjclp.xl.dsl

import com.tjclp.xl.addressing.{Column, Row}
import com.tjclp.xl.patch.Patch
import com.tjclp.xl.sheets.{ColumnProperties, RowProperties}

/**
 * Builder DSL for row and column properties.
 *
 * Provides a fluent API for creating row/column property patches with compile-time validation.
 *
 * Usage:
 * {{{
 *   import com.tjclp.xl.{*, given}
 *   import com.tjclp.xl.dsl.*
 *
 *   // Using builder DSL with patches
 *   val patch =
 *     (ref"A1" := "Revenue") ++
 *     row(0).height(30.0).toPatch ++        // Header row height
 *     col"A".width(25.0).toPatch ++          // First column width
 *     col"B".width(15.0).hidden.toPatch      // Hidden column
 *
 *   // Using extension methods on Row/Column
 *   val patch2 =
 *     Row.from0(0).height(30.0).toPatch ++
 *     Column.from0(0).width(25.0).toPatch
 *
 *   // Grouping/outlining
 *   val groupedPatch =
 *     rows(1 to 5).map(_.outlineLevel(1).toPatch).reduce(_ ++ _) ++
 *     row(0).collapsed.toPatch
 * }}}
 */
object RowColumnDsl:

  // ========== Row Builder ==========

  /**
   * Immutable builder for RowProperties.
   *
   * Each method returns a new builder with the updated property, enabling method chaining.
   */
  final class RowBuilder private[dsl] (
    private val targetRow: Row,
    private val props: RowProperties = RowProperties()
  ):
    /** Set row height in points */
    def height(h: Double): RowBuilder =
      RowBuilder(targetRow, props.copy(height = Some(h)))

    /** Mark row as hidden */
    def hidden: RowBuilder =
      RowBuilder(targetRow, props.copy(hidden = true))

    /** Mark row as visible (default) */
    def visible: RowBuilder =
      RowBuilder(targetRow, props.copy(hidden = false))

    /** Set outline/grouping level (0-7) */
    def outlineLevel(level: Int): RowBuilder =
      RowBuilder(targetRow, props.copy(outlineLevel = Some(level)))

    /** Mark outline group as collapsed */
    def collapsed: RowBuilder =
      RowBuilder(targetRow, props.copy(collapsed = true))

    /** Mark outline group as expanded (default) */
    def expanded: RowBuilder =
      RowBuilder(targetRow, props.copy(collapsed = false))

    /** Convert builder to a Patch for application to a Sheet */
    def toPatch: Patch =
      Patch.SetRowProperties(targetRow, props)

    /** Get the underlying Row */
    def row: Row = targetRow

    /** Get the current properties */
    def properties: RowProperties = props

  // ========== Column Builder ==========

  /**
   * Immutable builder for ColumnProperties.
   *
   * Each method returns a new builder with the updated property, enabling method chaining.
   */
  final class ColumnBuilder private[dsl] (
    private val targetCol: Column,
    private val props: ColumnProperties = ColumnProperties()
  ):
    /** Set column width in character units */
    def width(w: Double): ColumnBuilder =
      ColumnBuilder(targetCol, props.copy(width = Some(w)))

    /** Mark column as hidden */
    def hidden: ColumnBuilder =
      ColumnBuilder(targetCol, props.copy(hidden = true))

    /** Mark column as visible (default) */
    def visible: ColumnBuilder =
      ColumnBuilder(targetCol, props.copy(hidden = false))

    /** Set outline/grouping level (0-7) */
    def outlineLevel(level: Int): ColumnBuilder =
      ColumnBuilder(targetCol, props.copy(outlineLevel = Some(level)))

    /** Mark outline group as collapsed */
    def collapsed: ColumnBuilder =
      ColumnBuilder(targetCol, props.copy(collapsed = true))

    /** Mark outline group as expanded (default) */
    def expanded: ColumnBuilder =
      ColumnBuilder(targetCol, props.copy(collapsed = false))

    /** Convert builder to a Patch for application to a Sheet */
    def toPatch: Patch =
      Patch.SetColumnProperties(targetCol, props)

    /** Get the underlying Column */
    def column: Column = targetCol

    /** Get the current properties */
    def properties: ColumnProperties = props

  // ========== Entry Points ==========

  /** Create a RowBuilder for the given 0-based row index */
  def row(index: Int): RowBuilder =
    RowBuilder(Row.from0(index))

  /** Create RowBuilders for a range of 0-based row indices */
  def rows(range: Range): Seq[RowBuilder] =
    range.map(i => RowBuilder(Row.from0(i)))

  // ========== Row Extensions ==========

  extension (r: Row)
    /** Create a RowBuilder for this row */
    @annotation.targetName("rowBuilder")
    def builder: RowBuilder = RowBuilder(r)

    /** Create a RowBuilder with height set */
    @annotation.targetName("rowHeight")
    def height(h: Double): RowBuilder = RowBuilder(r).height(h)

    /** Create a RowBuilder with hidden flag set */
    @annotation.targetName("rowHidden")
    def hidden: RowBuilder = RowBuilder(r).hidden

    /** Create a RowBuilder with outline level set */
    @annotation.targetName("rowOutlineLevel")
    def outlineLevel(level: Int): RowBuilder = RowBuilder(r).outlineLevel(level)

    /** Create a RowBuilder with collapsed flag set */
    @annotation.targetName("rowCollapsed")
    def collapsed: RowBuilder = RowBuilder(r).collapsed

  // ========== Column Extensions ==========

  extension (c: Column)
    /** Create a ColumnBuilder for this column */
    @annotation.targetName("columnBuilder")
    def builder: ColumnBuilder = ColumnBuilder(c)

    /** Create a ColumnBuilder with width set */
    @annotation.targetName("columnWidth")
    def width(w: Double): ColumnBuilder = ColumnBuilder(c).width(w)

    /** Create a ColumnBuilder with hidden flag set */
    @annotation.targetName("columnHidden")
    def hidden: ColumnBuilder = ColumnBuilder(c).hidden

    /** Create a ColumnBuilder with outline level set */
    @annotation.targetName("columnOutlineLevel")
    def outlineLevel(level: Int): ColumnBuilder = ColumnBuilder(c).outlineLevel(level)

    /** Create a ColumnBuilder with collapsed flag set */
    @annotation.targetName("columnCollapsed")
    def collapsed: ColumnBuilder = ColumnBuilder(c).collapsed

end RowColumnDsl
