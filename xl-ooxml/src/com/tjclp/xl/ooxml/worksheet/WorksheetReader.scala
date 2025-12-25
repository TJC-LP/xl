package com.tjclp.xl.ooxml.worksheet

import scala.xml.Elem

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.ooxml.XmlUtil.{
  getAttr,
  getAttrOpt,
  getChild,
  getChildren,
  getTextPreservingWhitespace,
  parseTextRuns
}
import com.tjclp.xl.ooxml.{SharedStrings, XmlReadable}

/** Reader for parsing OoxmlWorksheet from XML */
object WorksheetReader extends XmlReadable[OoxmlWorksheet]:

  /** Parse worksheet from XML (XmlReadable trait compatibility) */
  def fromXml(elem: Elem): Either[String, OoxmlWorksheet] =
    fromXmlWithSST(elem, None)

  /** Parse worksheet from XML with optional SharedStrings table */
  def fromXmlWithSST(elem: Elem, sst: Option[SharedStrings]): Either[String, OoxmlWorksheet] =
    for
      // Parse sheetData (required)
      sheetDataElem <- getChild(elem, "sheetData")
      rowElems = getChildren(sheetDataElem, "row")
      rows <- parseRows(rowElems, sst)

      // Parse mergeCells (optional)
      mergedRanges <- parseMergeCells(elem)

      // Extract ALL preserved metadata elements (all optional)
      sheetPr = (elem \ "sheetPr").headOption.collect { case e: Elem => cleanNamespaces(e) }
      dimension = (elem \ "dimension").headOption.collect { case e: Elem => cleanNamespaces(e) }
      sheetViews = (elem \ "sheetViews").headOption.collect { case e: Elem => cleanNamespaces(e) }
      sheetFormatPr = (elem \ "sheetFormatPr").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }
      cols = (elem \ "cols").headOption.collect { case e: Elem => cleanNamespaces(e) }

      conditionalFormatting = elem.child.collect {
        case e: Elem if e.label == "conditionalFormatting" => cleanNamespaces(e)
      }
      printOptions = (elem \ "printOptions").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }
      rowBreaks = (elem \ "rowBreaks").headOption.collect { case e: Elem => cleanNamespaces(e) }
      colBreaks = (elem \ "colBreaks").headOption.collect { case e: Elem => cleanNamespaces(e) }
      customPropertiesWs = (elem \ "customProperties").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }

      pageMargins = (elem \ "pageMargins").headOption.collect { case e: Elem => cleanNamespaces(e) }
      pageSetup = (elem \ "pageSetup").headOption.collect { case e: Elem => cleanNamespaces(e) }
      headerFooter = (elem \ "headerFooter").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }

      drawing = (elem \ "drawing").headOption.collect { case e: Elem => cleanNamespaces(e) }
      legacyDrawing = (elem \ "legacyDrawing").headOption.collect { case e: Elem =>
        cleanNamespaces(e)
      }
      picture = (elem \ "picture").headOption.collect { case e: Elem => cleanNamespaces(e) }
      oleObjects = (elem \ "oleObjects").headOption.collect { case e: Elem => cleanNamespaces(e) }
      controls = (elem \ "controls").headOption.collect { case e: Elem => cleanNamespaces(e) }

      tableParts = (elem \ "tableParts").headOption.collect { case e: Elem => cleanNamespaces(e) }

      extLst = (elem \ "extLst").headOption.collect { case e: Elem => cleanNamespaces(e) }

      // Collect any other elements we don't explicitly handle
      knownElements = Set(
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "sheetData",
        "mergeCells",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "drawing",
        "legacyDrawing",
        "picture",
        "oleObjects",
        "controls",
        "extLst",
        // Additional elements from OOXML spec
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "phoneticPr",
        "conditionalFormatting",
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "legacyDrawingHF",
        "webPublishItems",
        "tableParts"
      )
      otherElements = elem.child.collect {
        case e: Elem if !knownElements.contains(e.label) => e
      }
    yield OoxmlWorksheet(
      rows,
      mergedRanges,
      sheetPr,
      dimension,
      sheetViews,
      sheetFormatPr,
      cols,
      conditionalFormatting,
      printOptions,
      rowBreaks,
      colBreaks,
      customPropertiesWs,
      pageMargins,
      pageSetup,
      headerFooter,
      drawing,
      legacyDrawing,
      picture,
      oleObjects,
      controls,
      tableParts,
      extLst,
      otherElements.toSeq,
      rootAttributes = elem.attributes,
      rootScope = Option(elem.scope).getOrElse(defaultWorksheetScope)
    )

  private def parseMergeCells(worksheetElem: Elem): Either[String, Set[CellRange]] =
    // mergeCells is optional
    (worksheetElem \ "mergeCells").headOption match
      case None => Right(Set.empty)
      case Some(mergeCellsElem: Elem) =>
        val mergeCellElems = getChildren(mergeCellsElem, "mergeCell")
        val parsed = mergeCellElems.map { elem =>
          for
            refStr <- getAttr(elem, "ref")
            range <- CellRange.parse(refStr)
          yield range
        }
        val errors = parsed.collect { case Left(err) => err }
        if errors.nonEmpty then Left(s"MergeCell parse errors: ${errors.mkString(", ")}")
        else Right(parsed.collect { case Right(range) => range }.toSet)
      case _ => Right(Set.empty) // Non-Elem node, ignore

  private def parseRows(
    elems: Seq[Elem],
    sst: Option[SharedStrings]
  ): Either[String, Seq[OoxmlRow]] =
    val parsed = elems.map { e =>
      for
        rStr <- getAttr(e, "r")
        rowIdx <- rStr.toIntOption.toRight(s"Invalid row index: $rStr")

        // Optimization: Extract all attributes once into Map for O(1) lookups (was O(n) per attr = O(11n) total)
        // This avoids 11 DOM traversals per row (1-2% speedup for 10k rows)
        attrs = e.attributes.asAttrMap

        // Extract ALL row attributes for byte-perfect preservation
        spans = attrs.get("spans")
        style = attrs.get("s").flatMap(_.toIntOption)
        height = attrs.get("ht").flatMap(_.toDoubleOption)
        customHeight = attrs.get("customHeight").contains("1")
        customFormat = attrs.get("customFormat").contains("1")
        hidden = attrs.get("hidden").contains("1")
        outlineLevel = attrs.get("outlineLevel").flatMap(_.toIntOption)
        collapsed = attrs.get("collapsed").contains("1")
        thickBot = attrs.get("thickBot").contains("1")
        thickTop = attrs.get("thickTop").contains("1")
        dyDescent = attrs.get("x14ac:dyDescent").flatMap(_.toDoubleOption)

        cellElems = getChildren(e, "c")
        cells <- parseCells(cellElems, sst)
      yield OoxmlRow(
        rowIdx,
        cells,
        spans,
        style,
        height,
        customHeight,
        customFormat,
        hidden,
        outlineLevel,
        collapsed,
        thickBot,
        thickTop,
        dyDescent
      )
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"Row parse errors: ${errors.mkString(", ")}")
    else Right(parsed.collect { case Right(row) => row })

  private def parseCells(
    elems: Seq[Elem],
    sst: Option[SharedStrings]
  ): Either[String, Seq[OoxmlCell]] =
    val parsed = elems.map { e =>
      for
        refStr <- getAttr(e, "r")
        ref <- ARef.parse(refStr)
        cellType = getAttrOpt(e, "t").getOrElse("")
        styleIdx = getAttrOpt(e, "s").flatMap(_.toIntOption)
        value <- parseCellValue(e, cellType, sst)
      yield OoxmlCell(ref, value, styleIdx, cellType)
    }

    val errors = parsed.collect { case Left(err) => err }
    if errors.nonEmpty then Left(s"Cell parse errors: ${errors.mkString(", ")}")
    else Right(parsed.collect { case Right(cell) => cell })

  private def parseCellValue(
    elem: Elem,
    cellType: String,
    sst: Option[SharedStrings]
  ): Either[String, CellValue] =
    // Check for formula element first (before cellType dispatch)
    (elem \ "f").headOption.map(_.text.trim) match
      case Some(formulaExpr) if formulaExpr.nonEmpty =>
        // Formula cell - parse cached value from <v> if present
        val cachedValue: Option[CellValue] = (elem \ "v").headOption.map(_.text).flatMap { vText =>
          // Infer cached value type from cellType attribute
          cellType match
            case "n" | "" =>
              try Some(CellValue.Number(BigDecimal(vText)))
              catch case _: NumberFormatException => None
            case "b" =>
              vText match
                case "1" | "true" => Some(CellValue.Bool(true))
                case "0" | "false" => Some(CellValue.Bool(false))
                case _ => None
            case "e" =>
              import com.tjclp.xl.cells.CellError
              CellError.parse(vText).toOption.map(CellValue.Error.apply)
            case "str" | "inlineStr" =>
              Some(CellValue.Text(vText))
            case _ => None
        }
        Right(CellValue.Formula(formulaExpr, cachedValue))

      case _ =>
        // No formula - dispatch on cellType as before
        parseCellValueWithoutFormula(elem, cellType, sst)

  /** Parse cell value when no formula element is present */
  private def parseCellValueWithoutFormula(
    elem: Elem,
    cellType: String,
    sst: Option[SharedStrings]
  ): Either[String, CellValue] =
    cellType match
      case "inlineStr" | "str" =>
        // Both "inlineStr" and "str" cell types use <is> for inline strings
        // Check for rich text (<is><r>) vs simple text (<is><t>)
        (elem \ "is").headOption match
          case None =>
            // Fallback: "str" type may have text in <v> element (preserving whitespace)
            (elem \ "v").headOption
              .collect { case e: Elem => e }
              .map(getTextPreservingWhitespace) match
              case Some(text) => Right(CellValue.Text(text))
              case None => Left(s"$cellType cell missing <is> element and <v> element")
          case Some(isElem: Elem) =>
            val rElems = getChildren(isElem, "r")

            if rElems.nonEmpty then
              // Rich text: parse runs with formatting
              parseTextRuns(rElems).map(CellValue.RichText.apply)
            else
              // Simple text: extract from <t> (preserving whitespace)
              (isElem \ "t").headOption
                .collect { case e: Elem => e }
                .map(getTextPreservingWhitespace) match
                case Some(text) => Right(CellValue.Text(text))
                case None => Left(s"$cellType <is> missing <t> element and has no <r> runs")
          case _ => Left(s"$cellType <is> is not an Elem")

      case "s" =>
        // SST index - resolve using SharedStrings table
        (elem \ "v").headOption.map(_.text) match
          case Some(idxStr) =>
            idxStr.toIntOption match
              case Some(idx) =>
                (for {
                  sharedStrings <- sst
                  entry <- sharedStrings.apply(idx)
                } yield sharedStrings.toCellValue(entry)) match
                  case Some(cellValue) => Right(cellValue)
                  case None =>
                    // SST index out of bounds -> CellError.Ref (not parse failure)
                    Right(CellValue.Error(com.tjclp.xl.cells.CellError.Ref))
              case None =>
                // Invalid SST index format -> CellError.Value
                Right(CellValue.Error(com.tjclp.xl.cells.CellError.Value))
          case None => Left("SST cell missing <v>")

      case "n" | "" =>
        // Number
        (elem \ "v").headOption.map(_.text) match
          case Some(numStr) =>
            try Right(CellValue.Number(BigDecimal(numStr)))
            catch case _: NumberFormatException => Left(s"Invalid number: $numStr")
          case None => Right(CellValue.Empty) // Empty numeric cell

      case "b" =>
        // Boolean
        (elem \ "v").headOption.map(_.text) match
          case Some("1") | Some("true") => Right(CellValue.Bool(true))
          case Some("0") | Some("false") => Right(CellValue.Bool(false))
          case other => Left(s"Invalid boolean value: $other")

      case "e" =>
        // Error
        (elem \ "v").headOption.map(_.text) match
          case Some(errStr) =>
            import com.tjclp.xl.cells.CellError
            CellError.parse(errStr).map(CellValue.Error.apply)
          case None => Left("Error cell missing <v>")

      case other =>
        Left(s"Unsupported cell type: $other")
