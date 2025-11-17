package com.tjclp.xl

/**
 * Immutable tracker for workbook modifications. Tracks all structural changes that impact whether a
 * sheet needs to be rewritten during a hybrid surgical write.
 */
final case class ModificationTracker(
  modifiedSheets: Set[Int] = Set.empty,
  deletedSheets: Set[Int] = Set.empty,
  reorderedSheets: Boolean = false,
  modifiedMetadata: Boolean = false
) derives CanEqual:

  /** True when no modifications have been recorded. */
  def isClean: Boolean =
    modifiedSheets.isEmpty &&
      deletedSheets.isEmpty &&
      !reorderedSheets &&
      !modifiedMetadata

  /** Mark a sheet as modified. */
  def markSheet(index: Int): ModificationTracker =
    copy(modifiedSheets = modifiedSheets + index)

  /**
   * Mark multiple sheets as modified in a single operation.
   *
   * @param indices
   *   Set of zero-based sheet indices to mark as modified
   * @return
   *   New tracker with specified sheets marked modified (returns this if indices empty)
   */
  def markSheets(indices: Set[Int]): ModificationTracker =
    if indices.isEmpty then this else copy(modifiedSheets = modifiedSheets ++ indices)

  /**
   * Mark a sheet as deleted and adjust all higher indices.
   *
   * When a sheet is deleted, all sheets at higher indices shift down by 1. This method adjusts the
   * tracked indices accordingly to maintain correctness for surgical writes. This includes shifting
   * both modifiedSheets and deletedSheets to reflect the new positions after deletion.
   *
   * Example: If sheets [0,1,2,3,4] exist and you delete(1) then delete(3), the deletedSheets will
   * contain {1, 2} (not {1, 3}) because after deleting sheet 1, what was sheet 3 becomes sheet 2.
   *
   * @param index
   *   Zero-based sheet index to mark as deleted
   * @return
   *   New tracker with deletion recorded and all higher indices shifted down
   */
  def delete(index: Int): ModificationTracker =
    copy(
      // Shift all existing deleted indices > index, then add the newly deleted index
      deletedSheets = deletedSheets.flatMap { i =>
        if i > index then Some(i - 1) // Shift down existing deletions
        else Some(i) // Keep unchanged
      } + index, // Add newly deleted index AFTER shifting
      // First shift all higher indices down, then remove the deleted index
      modifiedSheets = modifiedSheets
        .map { i =>
          if i > index then i - 1 // Shift down
          else i // Keep unchanged
        }
        .filterNot(_ == index) // Remove the deleted index after shifting
    )

  /** Indicate that sheet order changed. */
  def markReordered: ModificationTracker =
    if reorderedSheets then this else copy(reorderedSheets = true)

  /** Indicate workbook metadata changed. */
  def markMetadata: ModificationTracker =
    if modifiedMetadata then this else copy(modifiedMetadata = true)

  /**
   * Merge with another tracker, combining all modifications.
   *
   * Combines changes using set union for sheet indices and logical OR for boolean flags. Forms a
   * Monoid with `clean` as identity: associative and commutative for sets.
   *
   * @param other
   *   Tracker to merge with this one
   * @return
   *   New tracker containing union of all modifications from both trackers
   */
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
