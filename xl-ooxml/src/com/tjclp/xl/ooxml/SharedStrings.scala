package com.tjclp.xl.ooxml

import com.tjclp.xl.api.Workbook
import com.tjclp.xl.cell.CellValue
import com.tjclp.xl.richtext.RichText
import scala.xml.*
import XmlUtil.*
import java.text.Normalizer

/** SharedString entry: either plain text or rich text with formatting */
type SSTEntry = Either[String, RichText]

/**
 * Shared Strings Table (SST) for xl/sharedStrings.xml
 *
 * Deduplicates string values across the workbook. Cells reference strings by index rather than
 * embedding them inline.
 *
 * Supports both plain text and rich text (formatted runs). All plain text is normalized to NFC form
 * for consistent deduplication. RichText entries are deduplicated by plain text content.
 *
 * @param strings
 *   Ordered list of unique entries (plain text or rich text)
 * @param indexMap
 *   Plain text → index lookup (used for deduplication)
 * @param totalCount
 *   Total number of string cell instances in workbook (including duplicates)
 */
case class SharedStrings(
  strings: Vector[SSTEntry],
  indexMap: Map[String, Int],
  totalCount: Int // Total instances (>= strings.size)
) extends XmlWritable:

  /** Get index for plain text string (returns None if not found) */
  def indexOf(s: String): Option[Int] =
    indexMap.get(SharedStrings.normalize(s))

  /** Get index for RichText by plain text content (returns None if not found) */
  def indexOf(richText: RichText): Option[Int] =
    indexMap.get(SharedStrings.normalize(richText.toPlainText))

  /** Get entry at index (returns None if out of bounds) */
  def apply(index: Int): Option[SSTEntry] =
    if index >= 0 && index < strings.size then Some(strings(index))
    else None

  /**
   * Convert SST entry to CellValue.
   *
   * Used by worksheet reader to convert SST index references to cell values.
   */
  def toCellValue(entry: SSTEntry): CellValue = entry match
    case Left(text) => CellValue.Text(text)
    case Right(richText) => CellValue.RichText(richText)

  /** Count of unique entries */
  def uniqueCount: Int = strings.size

  /**
   * Serialize SharedStrings to XML (ECMA-376 Part 1, §18.4.9)
   *
   * REQUIRES: strings is a valid Vector of SST entries (plain text or RichText) ENSURES:
   *   - Emits <sst> root element with xmlns namespace
   *   - count attribute = totalCount (total string cell instances)
   *   - uniqueCount attribute = strings.size (unique entries)
   *   - Plain text: <si><t>text</t></si>
   *   - RichText: <si><r><rPr>...</rPr><t>text</t></r>...</si>
   *   - Adds xml:space="preserve" for text with leading/trailing/double spaces
   * DETERMINISTIC: Yes (Vector iteration order is stable) ERROR CASES: None (total function)
   *
   * @return
   *   XML element representing xl/sharedStrings.xml
   */
  def toXml: Elem =
    val siElems = strings.map {
      case Left(s) =>
        // Plain text: <si><t>text</t></si>
        val needsPreserve = needsXmlSpacePreserve(s)
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

      case Right(richText) =>
        // RichText: <si><r><rPr>...</rPr><t>text</t></r>...</si>
        val runElems = richText.runs.map { run =>
          // Use preserved raw <rPr> if available (byte-perfect), otherwise build from Font
          val rPrElems = run.rawRPrXml.flatMap { xmlString =>
            // Parse preserved XML string back to Elem for byte-perfect preservation
            try Some(scala.xml.XML.loadString(xmlString).asInstanceOf[Elem])
            catch case _: Exception => None
          }.toList match
            case preserved if preserved.nonEmpty => preserved
            case _ =>
              // Build from Font model if no raw XML or parse failed
              run.font.map { f =>
                import com.tjclp.xl.style.font.Font
                val fontProps = Seq.newBuilder[Elem]

                if f.bold then fontProps += elem("b")()
                if f.italic then fontProps += elem("i")()
                if f.underline then fontProps += elem("u")()

                f.color.foreach { c =>
                  fontProps += elem("color", "rgb" -> c.toHex.drop(1))() // Drop # prefix
                }

                fontProps += elem("sz", "val" -> f.sizePt.toString)()
                fontProps += elem("name", "val" -> f.name)()

                elem("rPr")(fontProps.result()*)
              }.toList

          // Build <t> with optional xml:space="preserve"
          val needsPreserve = needsXmlSpacePreserve(run.text)
          val textElem =
            if needsPreserve then
              Elem(
                null,
                "t",
                PrefixedAttribute("xml", "space", "preserve", Null),
                TopScope,
                true,
                Text(run.text)
              )
            else elem("t")(Text(run.text))

          elem("r")(rPrElems ++ Seq(textElem)*)
        }

        elem("si")(runElems*)
    }

    elem(
      "sst",
      "xmlns" -> nsSpreadsheetML,
      "count" -> totalCount.toString, // Total instances
      "uniqueCount" -> strings.size.toString // Unique entries
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
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def fromStrings(strings: Iterable[String], totalCount: Option[Int] = None): SharedStrings =
    // Stream through strings with LinkedHashSet for O(1) deduplication (50-70% memory reduction)
    import scala.collection.mutable
    val seen = mutable.LinkedHashSet.empty[String]
    var count = 0
    strings.foreach { s =>
      count += 1
      seen += normalize(s)
    }
    val normalized = seen.toVector.map(s => Left(s): SSTEntry)
    val indexMap = seen.toVector.zipWithIndex.toMap
    val finalCount = totalCount.getOrElse(count)
    SharedStrings(normalized, indexMap, finalCount)

  /** Normalize string to NFC form for consistent comparison */
  def normalize(s: String): String =
    Normalizer.normalize(s, Normalizer.Form.NFC)

  /**
   * Build SharedStrings table from workbook text and rich text cells
   *
   * Extracts all CellValue.Text and CellValue.RichText instances, deduplicates by plain text
   * content, and tracks total count.
   *
   * REQUIRES: wb contains valid CellValue.Text or CellValue.RichText cells ENSURES:
   *   - strings contains unique entries (plain text or RichText)
   *   - indexMap maps plain text to SST index (deduplication by content)
   *   - totalCount = number of text/richtext cell instances in workbook
   *   - totalCount >= strings.size (equality only when no duplicates)
   *   - Iteration order is stable (sheets processed in Vector order)
   * DETERMINISTIC: Yes (stable Vector traversal, deterministic deduplication) ERROR CASES: None
   * (total function, handles empty workbook → empty SST)
   *
   * @param wb
   *   Workbook to extract strings from
   * @return
   *   SharedStrings with deduplicated entries and total count
   */
  def fromWorkbook(wb: Workbook): SharedStrings =
    // Stream entries using iterator for lazy evaluation
    val allEntries = wb.sheets.iterator.flatMap { sheet =>
      sheet.cells.values.iterator.flatMap { cell =>
        cell.value match
          case com.tjclp.xl.cell.CellValue.Text(s) => Iterator.single(Left(s): SSTEntry)
          case com.tjclp.xl.cell.CellValue.RichText(rt) => Iterator.single(Right(rt): SSTEntry)
          case _ => Iterator.empty
      }
    }
    fromEntries(allEntries.to(Iterable), None)

  /**
   * Create SharedStrings from SST entries with deduplication.
   *
   * Deduplicates by plain text content (RichText entries are keyed by toPlainText).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  def fromEntries(entries: Iterable[SSTEntry], totalCount: Option[Int] = None): SharedStrings =
    import scala.collection.mutable
    val seen = mutable.LinkedHashSet.empty[String]
    val entryVec = Vector.newBuilder[SSTEntry]
    var count = 0

    entries.foreach { entry =>
      count += 1
      val key = entry match
        case Left(text) => normalize(text)
        case Right(richText) => normalize(richText.toPlainText)

      if seen.add(key) then entryVec += entry
    }

    val normalized = entryVec.result()
    val indexMap = normalized.zipWithIndex.map { case (entry, idx) =>
      val key = entry match
        case Left(text) => text
        case Right(richText) => richText.toPlainText
      normalize(key) -> idx
    }.toMap

    SharedStrings(normalized, indexMap, totalCount.getOrElse(count))

  /**
   * Determine if using SST saves space vs inline strings
   *
   * Heuristic: Use SST if total string bytes saved > SST overhead. SST overhead ≈ 200 bytes + 50
   * bytes per unique string. Inline overhead ≈ 30 bytes per cell + string length.
   */
  def shouldUseSST(wb: Workbook): Boolean =
    val textCells = wb.sheets.flatMap { sheet =>
      sheet.cells.values.flatMap { cell =>
        cell.value match
          case com.tjclp.xl.cell.CellValue.Text(s) => Some(s)
          case _ => None
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

    val entries = siElems.map { si =>
      val rElems = getChildren(si, "r")

      if rElems.nonEmpty then
        // RichText: parse runs with formatting
        parseTextRuns(rElems).map(rt => Right(rt): SSTEntry)
      else
        // Simple text: extract from <t>
        (si \ "t").headOption.map(_.text) match
          case Some(text) => Right(Left(normalize(text)): SSTEntry)
          case None => Left("SharedString <si> missing <t> element and has no <r> runs")
    }

    val errors = entries.collect { case Left(err) => err }

    if errors.nonEmpty then Left(s"SharedStrings parse errors: ${errors.mkString(", ")}")
    else
      val entryVec = entries.collect { case Right(entry) => entry }.toVector

      // Build index map using plain text representation
      val indexMap = entryVec.zipWithIndex.map { case (entry, idx) =>
        val key = entry match
          case Left(text) => text
          case Right(richText) => richText.toPlainText
        normalize(key) -> idx
      }.toMap

      // Try to read totalCount from count attribute, fall back to uniqueCount
      val totalCount = elem
        .attribute("count")
        .flatMap(_.text.toIntOption)
        .getOrElse(entryVec.size) // Default to unique count if missing

      Right(SharedStrings(entryVec, indexMap, totalCount))
