package com.tjclp.xl

/**
 * Immutable tracker for workbook modifications. Tracks all structural
 * changes that impact whether a sheet needs to be rewritten during a hybrid
 * surgical write.
 */
final case class ModificationTracker(
  modifiedSheets: Set[Int] = Set.empty,
  deletedSheets: Set[Int] = Set.empty,
  reorderedSheets: Boolean = false,
  modifiedMetadata: Boolean = false
):

  /** True when no modifications have been recorded. */
  def isClean: Boolean =
    modifiedSheets.isEmpty &&
      deletedSheets.isEmpty &&
      !reorderedSheets &&
      !modifiedMetadata

  /** Mark a sheet as modified. */
  def markSheet(index: Int): ModificationTracker =
    copy(modifiedSheets = modifiedSheets + index)

  /** Mark several sheets as modified. */
  def markSheets(indices: Set[Int]): ModificationTracker =
    if indices.isEmpty then this else copy(modifiedSheets = modifiedSheets ++ indices)

  /** Mark a sheet as deleted. */
  def deleteSheet(index: Int): ModificationTracker =
    copy(
      deletedSheets = deletedSheets + index,
      modifiedSheets = modifiedSheets - index
    )

  /** Indicate that sheet order changed. */
  def markReordered: ModificationTracker =
    if reorderedSheets then this else copy(reorderedSheets = true)

  /** Indicate workbook metadata changed. */
  def markMetadata: ModificationTracker =
    if modifiedMetadata then this else copy(modifiedMetadata = true)

  /** Merge with another tracker (union of changes). */
  def merge(other: ModificationTracker): ModificationTracker =
    ModificationTracker(
      modifiedSheets = modifiedSheets ++ other.modifiedSheets,
      deletedSheets = deletedSheets ++ other.deletedSheets,
      reorderedSheets = reorderedSheets || other.reorderedSheets,
      modifiedMetadata = modifiedMetadata || other.modifiedMetadata
    )

object ModificationTracker:
  /** Clean tracker (no modifications). */
  val clean: ModificationTracker = ModificationTracker()

  /** Tracker with all sheets marked as modified. */
  def allModified(sheetCount: Int): ModificationTracker =
    ModificationTracker(modifiedSheets = (0 until sheetCount).toSet)
