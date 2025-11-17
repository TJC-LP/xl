package com.tjclp.xl.ooxml

import java.util.zip.ZipEntry

/** Metadata about every part inside the originating XLSX ZIP. */
final case class PartManifestEntry(
  path: String,
  parsed: Boolean,
  sheetIndex: Option[Int],
  relationships: Set[String],
  size: Option[Long],
  compressedSize: Option[Long],
  crc: Option[Long],
  method: Option[Int]
) derives CanEqual

object PartManifestEntry:
  def unparsed(path: String): PartManifestEntry =
    PartManifestEntry(path, parsed = false, None, Set.empty, None, None, None, None)

/** Complete manifest for all ZIP entries. */
final case class PartManifest(entries: Map[String, PartManifestEntry]) derives CanEqual:

  /**
   * Returns set of all ZIP entry paths that XL parsed (e.g., worksheets, styles.xml, SST). These
   * parts will be regenerated during write operations.
   */
  def parsedParts: Set[String] = entries.collect {
    case (path, entry) if entry.parsed => path
  }.toSet

  /**
   * Returns set of all ZIP entry paths that XL did not parse (e.g., charts, drawings). These parts
   * will be preserved byte-for-byte during surgical write operations.
   */
  def unparsedParts: Set[String] = entries.collect {
    case (path, entry) if !entry.parsed => path
  }.toSet

  /**
   * Returns the sheet indices that depend on the given ZIP entry path.
   *
   * @param path
   *   ZIP entry path (e.g., "xl/worksheets/sheet1.xml")
   * @return
   *   Set containing the sheet index if this entry is a worksheet, empty set otherwise
   */
  def dependentSheets(path: String): Set[Int] =
    entries.get(path).flatMap(_.sheetIndex).map(Set(_)).getOrElse(Set.empty)

  /**
   * Returns the set of relationship IDs that the given entry depends on.
   *
   * @param path
   *   ZIP entry path
   * @return
   *   Set of relationship IDs referenced by this entry, empty set if entry not found
   */
  def relationshipsFor(path: String): Set[String] =
    entries.get(path).map(_.relationships).getOrElse(Set.empty)

  def contains(path: String): Boolean = entries.contains(path)

object PartManifest:
  val empty: PartManifest = PartManifest(Map.empty)

/** Builder used by the reader to accumulate manifest entries. */
final class PartManifestBuilder:
  private val entries = scala.collection.mutable.Map.empty[String, PartManifestEntry]

  def recordParsed(
    path: String,
    sheetIndex: Option[Int] = None,
    relationships: Set[String] = Set.empty
  ): PartManifestBuilder =
    updateEntry(path) { entry =>
      entry.copy(
        parsed = true,
        sheetIndex = sheetIndex.orElse(entry.sheetIndex),
        relationships = relationships
      )
    }

  def recordUnparsed(
    path: String,
    sheetIndex: Option[Int] = None,
    relationships: Set[String] = Set.empty
  ): PartManifestBuilder =
    updateEntry(path) { entry =>
      entry.copy(
        parsed = false,
        sheetIndex = sheetIndex.orElse(entry.sheetIndex),
        relationships = relationships
      )
    }

  def recordRelationships(path: String, relationships: Set[String]): PartManifestBuilder =
    updateEntry(path)(entry => entry.copy(relationships = relationships))

  def withSheetIndex(path: String, sheetIndex: Int): PartManifestBuilder =
    updateEntry(path)(entry => entry.copy(sheetIndex = Some(sheetIndex)))

  def +=(entry: ZipEntry): PartManifestBuilder =
    updateEntry(entry.getName) { current =>
      current.copy(
        size = sizeOf(entry.getSize),
        compressedSize = sizeOf(entry.getCompressedSize),
        crc = sizeOf(entry.getCrc),
        method = Some(entry.getMethod)
      )
    }

  def build(): PartManifest = PartManifest(entries.toMap)

  private def updateEntry(path: String)(
    f: PartManifestEntry => PartManifestEntry
  ): PartManifestBuilder =
    val updated = f(entries.getOrElse(path, PartManifestEntry.unparsed(path)))
    entries.update(path, updated)
    this

  private def sizeOf(value: Long): Option[Long] =
    Option.when(value >= 0L)(value)
