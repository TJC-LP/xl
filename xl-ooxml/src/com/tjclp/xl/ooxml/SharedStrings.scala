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
 */
case class SharedStrings(
  strings: Vector[String], // Ordered list of unique strings
  indexMap: Map[String, Int] // String → index lookup
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

  def toXml: Elem =
    val siElems = strings.map { s =>
      // Check if string needs xml:space="preserve"
      val needsPreserve = s.startsWith(" ") || s.endsWith(" ") || s.contains("  ")
      val tElem =
        if needsPreserve then
          Elem(
            null,
            "t",
            new UnprefixedAttribute("xml:space", "preserve", Null),
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
      "count" -> strings.size.toString,
      "uniqueCount" -> strings.size.toString
    )(siElems*)

object SharedStrings extends XmlReadable[SharedStrings]:
  val empty: SharedStrings = SharedStrings(Vector.empty, Map.empty)

  /** Create SharedStrings from a collection of strings with deduplication */
  def fromStrings(strings: Iterable[String]): SharedStrings =
    val normalized = strings.map(normalize).toVector.distinct
    val indexMap = normalized.zipWithIndex.toMap
    SharedStrings(normalized, indexMap)

  /** Normalize string to NFC form for consistent comparison */
  def normalize(s: String): String =
    Normalizer.normalize(s, Normalizer.Form.NFC)

  /** Build SST from all text cells in a workbook */
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
    fromStrings(allStrings)

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
      Right(SharedStrings(stringVec, indexMap))
