package com.tjclp.xl.ooxml

import scala.xml.*
import XmlUtil.*
import java.text.Normalizer

/**
 * Shared Strings Table (SST) for xl/sharedStrings.xml
 *
 * Deduplicates string values across the workbook. Cells reference strings by index rather than
 * embedding them inline.
 *
 * All strings are normalized to NFC form for consistent deduplication.
 *
 * @param strings
 *   Ordered list of unique strings
 * @param indexMap
 *   String → index lookup
 * @param totalCount
 *   Total number of string cell instances in workbook (including duplicates)
 */
case class SharedStrings(
  strings: Vector[String],
  indexMap: Map[String, Int],
  totalCount: Int // Total instances (>= strings.size)
) extends XmlWritable:

  /** Get index for string (returns None if not found) */
  def indexOf(s: String): Option[Int] =
    indexMap.get(SharedStrings.normalize(s))

  /** Get string at index (returns None if out of bounds) */
  def apply(index: Int): Option[String] =
    if index >= 0 && index < strings.size then Some(strings(index))
    else None

  /** Count of unique strings */
  def uniqueCount: Int = strings.size

  /**
   * Serialize SharedStrings to XML (ECMA-376 Part 1, §18.4.9)
   *
   * REQUIRES: strings is a valid Vector of normalized strings ENSURES:
   *   - Emits <sst> root element with xmlns namespace
   *   - count attribute = totalCount (total string cell instances)
   *   - uniqueCount attribute = strings.size (unique strings)
   *   - Each string emitted as <si><t>text</t></si>
   *   - Adds xml:space="preserve" for strings with leading/trailing/double spaces
   *   - Uses PrefixedAttribute("xml", "space", "preserve") for proper namespace
   * DETERMINISTIC: Yes (Vector iteration order is stable) ERROR CASES: None (total function)
   *
   * @return
   *   XML element representing xl/sharedStrings.xml
   */
  def toXml: Elem =
    val siElems = strings.map { s =>
      // Check if string needs xml:space="preserve"
      val needsPreserve = s.startsWith(" ") || s.endsWith(" ") || s.contains("  ")
      val tElem =
        if needsPreserve then
          Elem(
            null,
            "t",
            PrefixedAttribute("xml", "space", "preserve", Null),
            TopScope,
            false,
            Text(s)
          )
        else elem("t")(Text(s))

      elem("si")(tElem)
    }

    elem(
      "sst",
      "xmlns" -> nsSpreadsheetML,
      "count" -> totalCount.toString, // Total instances
      "uniqueCount" -> strings.size.toString // Unique strings
    )(siElems*)

object SharedStrings extends XmlReadable[SharedStrings]:
  val empty: SharedStrings = SharedStrings(Vector.empty, Map.empty, 0)

  /**
   * Create SharedStrings from a collection of strings with deduplication
   * @param strings
   *   Collection of strings (may contain duplicates)
   * @param totalCount
   *   Total number of string instances (defaults to strings.size for backward compatibility)
   */
  def fromStrings(strings: Iterable[String], totalCount: Option[Int] = None): SharedStrings =
    val stringVec = strings.toVector
    val normalized = stringVec.map(normalize).distinct
    val indexMap = normalized.zipWithIndex.toMap
    val count = totalCount.getOrElse(stringVec.size) // Default to input size if not specified
    SharedStrings(normalized, indexMap, count)

  /** Normalize string to NFC form for consistent comparison */
  def normalize(s: String): String =
    Normalizer.normalize(s, Normalizer.Form.NFC)

  /**
   * Build SharedStrings table from workbook text cells
   *
   * Extracts all CellValue.Text instances, deduplicates them, and tracks total count.
   *
   * REQUIRES: wb contains only valid CellValue.Text cells ENSURES:
   *   - strings contains unique text values (NFC-normalized and deduplicated)
   *   - indexMap maps each string to its SST index (0-based)
   *   - totalCount = number of CellValue.Text instances in workbook
   *   - totalCount >= strings.size (equality only when no duplicates)
   *   - Iteration order is stable (sheets processed in Vector order)
   * DETERMINISTIC: Yes (stable Vector traversal, deterministic deduplication) ERROR CASES: None
   * (total function, handles empty workbook → empty SST)
   *
   * @param wb
   *   Workbook to extract strings from
   * @return
   *   SharedStrings with deduplicated strings and total count
   */
  def fromWorkbook(wb: com.tjclp.xl.Workbook): SharedStrings =
    val allStrings = wb.sheets.flatMap { sheet =>
      sheet.cells.values.collect {
        case cell if cell.value match {
              case com.tjclp.xl.CellValue.Text(_) => true
              case _ => false
            } =>
          cell.value match {
            case com.tjclp.xl.CellValue.Text(s) => s
            case _ => "" // Should never happen due to filter
          }
      }
    }
    // Count total instances before deduplication
    val totalCount = allStrings.size
    fromStrings(allStrings, Some(totalCount))

  /**
   * Determine if using SST saves space vs inline strings
   *
   * Heuristic: Use SST if total string bytes saved > SST overhead. SST overhead ≈ 200 bytes + 50
   * bytes per unique string. Inline overhead ≈ 30 bytes per cell + string length.
   */
  def shouldUseSST(wb: com.tjclp.xl.Workbook): Boolean =
    val textCells = wb.sheets.flatMap { sheet =>
      sheet.cells.values.collect {
        case cell if cell.value match {
              case com.tjclp.xl.CellValue.Text(s) => true
              case _ => false
            } =>
          cell.value match {
            case com.tjclp.xl.CellValue.Text(s) => s
            case _ => ""
          }
      }
    }

    if textCells.isEmpty then false
    else
      val uniqueStrings = textCells.toSet
      val totalCells = textCells.size
      val uniqueCount = uniqueStrings.size

      // Simple heuristic: use SST if we have duplicates and >10 total cells
      totalCells > uniqueCount && totalCells > 10

  def fromXml(elem: Elem): Either[String, SharedStrings] =
    val siElems = getChildren(elem, "si")

    val strings = siElems.map { si =>
      // Get text from <t> element
      (si \ "t").headOption.map(_.text) match
        case Some(text) => Right(normalize(text))
        case None => Left("SharedString <si> missing <t> element")
    }

    val errors = strings.collect { case Left(err) => err }

    if errors.nonEmpty then Left(s"SharedStrings parse errors: ${errors.mkString(", ")}")
    else
      val stringVec = strings.collect { case Right(s) => s }.toVector
      val indexMap = stringVec.zipWithIndex.toMap

      // Try to read totalCount from count attribute, fall back to uniqueCount
      val totalCount = elem
        .attribute("count")
        .flatMap(_.text.toIntOption)
        .getOrElse(stringVec.size) // Default to unique count if missing

      Right(SharedStrings(stringVec, indexMap, totalCount))
