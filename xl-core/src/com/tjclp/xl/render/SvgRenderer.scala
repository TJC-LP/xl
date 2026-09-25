package com.tjclp.xl.render

import com.tjclp.xl.addressing.{ARef, CellRange, Column, Row}
import com.tjclp.xl.cells.{Cell, CellValue}
import com.tjclp.xl.richtext.TextRun
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.styles.alignment.{HAlign, VAlign}
import com.tjclp.xl.styles.border.{Border, BorderSide, BorderStyle}
import com.tjclp.xl.styles.color.{Color, ThemePalette}
import com.tjclp.xl.styles.fill.Fill
import com.tjclp.xl.styles.font.{Font, Underline}

/**
 * Renders Excel sheets to SVG images with styled cells.
 *
 * SVG output can be viewed directly by Claude as images, making it useful for visualizing formatted
 * Excel data including colors, fonts, and borders.
 */
object SvgRenderer:
  import RenderUtils.*

  /**
   * Export a sheet range to an SVG image.
   *
   * @param sheet
   *   The sheet to export
   * @param range
   *   The cell range to export
   * @param includeStyles
   *   Whether to include cell styling (colors, fonts, borders)
   * @param theme
   *   Theme palette for resolving theme colors (default: Office theme)
   * @param showLabels
   *   Whether to show column letters (A, B, C...) and row numbers (1, 2, 3...) (default: false)
   * @param showGridlines
   *   Whether to show cell gridlines (default: false, matches HTML behavior). A sheet-level
   *   `SheetView(showGridLines = false)` suppresses gridlines even when this flag is set, matching
   *   how Excel renders such sheets (GH-258).
   * @return
   *   SVG string
   */
  def toSvg(
    sheet: Sheet,
    range: CellRange,
    includeStyles: Boolean = true,
    theme: ThemePalette = ThemePalette.office,
    showLabels: Boolean = false,
    showGridlines: Boolean = false
  ): String =
    toSvgResolving(ResolvedCell(_, _))(
      sheet,
      range,
      includeStyles,
      theme,
      showLabels,
      showGridlines
    )

  /**
   * [[toSvg]] with the per-cell resolver injected. Each rendered cell is resolved exactly once
   * (`RenderUtilsSpec` counts through this seam): every overflow, anchor, hash and text decision
   * reads the one [[ResolvedCell]], so a Custom format code is parsed once per cell.
   */
  private[render] def toSvgResolving(resolve: (Cell, Sheet) => ResolvedCell)(
    sheet: Sheet,
    range: CellRange,
    includeStyles: Boolean,
    theme: ThemePalette,
    showLabels: Boolean,
    showGridlines: Boolean
  ): String =
    // GH-258: the sheet's own view settings win when they disable gridlines (templates that are
    // gridline-free in Excel must stay gridline-free in exports).
    val gridlinesEnabled = showGridlines && sheet.viewSettings.forall(_.showGridLines)
    val startCol = range.start.col.index0
    val endCol = range.end.col.index0
    val startRow = range.start.row.index0
    val endRow = range.end.row.index0

    // Calculate column widths and row heights using shared utilities
    val colWidths = calculateColumnWidths(sheet, range)
    val rowHeights = calculateRowHeights(sheet, range)

    // Calculate x/y offsets based on whether labels are shown
    val xOffset = if showLabels then HeaderWidth else 0
    val yOffset = if showLabels then HeaderHeight else 0

    // Pre-calculate positions; full boundary scans (n+1 entries) also locate grid edges
    // for shared-edge border resolution (GH-298)
    val colBoundaries = colWidths.scanLeft(xOffset)(_ + _)
    val rowBoundaries = rowHeights.scanLeft(yOffset)(_ + _)
    val colXPositions = colBoundaries.dropRight(1)
    val rowYPositions = rowBoundaries.dropRight(1)

    // Calculate total dimensions
    val totalWidth = xOffset + colWidths.sum
    val totalHeight = yOffset + rowHeights.sum

    val sb = new StringBuilder

    // SVG header
    sb.append(
      s"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 $totalWidth $totalHeight" """
    )
    sb.append(s"""width="$totalWidth" height="$totalHeight">\n""")

    // Embedded styles - note: 11pt = ~15px (11 * 4/3)
    // Gridlines are now applied via inline stroke attributes for reliable cross-renderer support.
    // IMPORTANT: .cell-text must NOT declare font-family/font-size — CSS class rules outrank
    // presentation attributes in the cascade, which silently overrode per-cell fonts in every
    // CSS-aware rasterizer (GH-255). Cell text carries explicit font attributes on every path
    // instead; the class remains for semantic grouping only.
    sb.append(s"""  <style>
    .header { fill: #E0E0E0; stroke: #999999; stroke-width: 1; }
    .header-text { font-family: 'Segoe UI', Arial, sans-serif; font-size: 15px; fill: #333333; }
  </style>
""")

    // Clip path buffer - will be rendered in <defs> after cell loop
    val clipPathBuffer = new StringBuilder

    // Column headers (A, B, C...) - only if showLabels
    if showLabels then
      sb.append("  <g class=\"col-headers\">\n")
      (startCol to endCol).foreach { col =>
        val colIdx = col - startCol
        val colLetter = Column.from0(col).toLetter
        val width = colWidths(colIdx)
        val xPos = colXPositions(colIdx)
        sb.append(
          s"""    <rect x="$xPos" y="0" width="$width" height="$HeaderHeight" class="header"/>\n"""
        )
        val textX = xPos + width / 2
        val textY = HeaderHeight / 2 + 4
        sb.append(
          s"""    <text x="$textX" y="$textY" text-anchor="middle" class="header-text">$colLetter</text>\n"""
        )
      }
      sb.append("  </g>\n")

      // Row headers (1, 2, 3...)
      sb.append("  <g class=\"row-headers\">\n")
      (startRow to endRow).foreach { row =>
        val rowIdx = row - startRow
        val rowNum = row + 1
        val y = rowYPositions(rowIdx)
        val rowHeight = rowHeights(rowIdx)
        sb.append(
          s"""    <rect x="0" y="$y" width="$HeaderWidth" height="$rowHeight" class="header"/>\n"""
        )
        val textX = HeaderWidth / 2
        val textY = y + rowHeight / 2 + 4
        sb.append(
          s"""    <text x="$textX" y="$textY" text-anchor="middle" class="header-text">$rowNum</text>\n"""
        )
      }
      sb.append("  </g>\n")
    end if

    // Three-pass rendering: backgrounds first, then borders, then text on top
    // This ensures text overflow is visible and not covered by adjacent cell backgrounds
    val textBuffer = new StringBuilder
    // Border declarations keyed by unit grid edge: (boundary index, cell index along the
    // edge). Adjacent cells declaring the same edge are resolved heavier-wins (GH-298)
    // instead of double-drawn with painter's order deciding.
    val hEdges = scala.collection.mutable.Map.empty[(Int, Int), EdgeDecl]
    val vEdges = scala.collection.mutable.Map.empty[(Int, Int), EdgeDecl]

    // Clip paths will be rendered in defs section after cell loop completes

    // Pass 1: Cell backgrounds (also generates clip paths in clipPathBuffer)
    sb.append("  <g class=\"cells\">\n")
    val gridStroke = """ stroke="#D0D0D0" stroke-width="0.5""""
    (startRow to endRow).foreach { row =>
      val rowIdx = row - startRow
      val y = rowYPositions(rowIdx)
      val rowHeight = rowHeights(rowIdx)

      // The row's drawn cells — every column except the interior of a merge — each resolved
      // ONCE: the spans below, and the fill, anchor, text and hash after them, all read it, and
      // resolving re-formats the value and re-parses a Custom code.
      val drawn = (startCol to endCol).flatMap { col =>
        val ref = ARef.from0(col, row)
        val mergeRange = sheet.getMergedRange(ref)
        if mergeRange.exists(_.start != ref) then None
        else Some((col, ref, mergeRange, sheet.cells.get(ref).map(resolve(_, sheet))))
      }

      // Text spill spans, decided before any rect is drawn: a leftward spill covers cells
      // drawn earlier in the row, and no gridline may cross the text.
      val spans = drawn.flatMap { case (col, ref, _, resolvedOpt) =>
        resolvedOpt
          .map { resolved =>
            val width = colWidths(col - startCol)
            col -> overflowSpan(resolved, ref, width, colWidths, sheet, startCol, endCol)
          }
          .filter(_._2.spills)
      }.toMap
      val spannedCols = spans.flatMap { case (col, span) =>
        (col - span.left) to (col + span.right)
      }.toSet

      drawn.foreach { case (col, ref, mergeRange, resolvedOpt) =>
        val colIdx = col - startCol
        val xPos = colXPositions(colIdx)
        val cellOpt = resolvedOpt.map(_.cell)

        // A cell's own box: its column, or its merge clamped to the window. A spill never
        // widens it — every cell under a spilled text keeps its own fill and borders.
        val (effectiveWidth, effectiveHeight) = mergeRange match
          case Some(range) =>
            val mergeEndCol = math.min(range.end.col.index0, endCol)
            val mergeEndRow = math.min(range.end.row.index0, endRow)
            (
              (colIdx to (mergeEndCol - startCol)).map(colWidths).sum,
              (rowIdx to (mergeEndRow - startRow)).map(rowHeights).sum
            )
          case None => (colWidths(colIdx), rowHeight)

        // The text's clip: its own box, or the columns its text spills across. Keyed by cell
        // ref: pixel positions collide when hidden rows/columns collapse to the same
        // coordinates, emitting duplicate ids — invalid SVG (GH-298).
        val (clipX, clipWidth) = spans.get(col) match
          case Some(span) =>
            val left = colBoundaries(colIdx - span.left)
            (left, colBoundaries(colIdx + span.right + 1) - left)
          case None => (xPos, effectiveWidth)
        val clipId = s"clip-${ref.toA1}"
        clipPathBuffer.append(
          s"""    <clipPath id="$clipId"><rect x="$clipX" y="$y" width="$clipWidth" height="$effectiveHeight"/></clipPath>\n"""
        )

        // Cell background (borders rendered separately as line elements)
        val fillAttr = cellOpt
          .flatMap(c => if includeStyles then cellStyleToSvg(c, sheet, theme) else None)
          .getOrElse("""fill="#FFFFFF"""")

        // Gridlines: explicit stroke attributes (CSS-only approach unreliable across renderers);
        // a spanned cell's are drawn once around its whole span below
        val strokeAttr =
          if gridlinesEnabled && !spannedCols.contains(col) then gridStroke else ""

        sb.append(
          s"""    <rect x="$xPos" y="$y" width="$effectiveWidth" height="$effectiveHeight" """
        )
        sb.append(s"""$fillAttr$strokeAttr class="cell"/>\n""")

        // Collect border declarations for second pass, decomposed into the unit grid edges
        // of the cell's own (merge-expanded) rect, so shared edges resolve to the heavier
        // declaration (GH-298)
        if includeStyles then
          resolvedOpt.flatMap(_.style).foreach { style =>
            if style.border != Border.none then
              val spanCols = mergeRange match
                case Some(range) => math.min(range.end.col.index0, endCol) - col + 1
                case None => 1
              val spanRows = mergeRange match
                case Some(range) => math.min(range.end.row.index0, endRow) - row + 1
                case None => 1
              declareCellEdges(hEdges, vEdges, style.border, col, row, spanCols, spanRows)
          }

        // Collect text for third pass (skip hidden rows/cols), clipped to its clip box
        if effectiveHeight > 0 && effectiveWidth > 0 then
          resolvedOpt.foreach { resolved =>
            val box = CellBox(xPos, y, effectiveWidth, effectiveHeight)
            textBuffer.append(cellText(resolved, sheet, box, clipId, includeStyles, theme))
          }
      }

      if gridlinesEnabled then
        spans.toList.sortBy(_._1).foreach { case (col, span) =>
          val left = colBoundaries(col - startCol - span.left)
          val width = colBoundaries(col - startCol + span.right + 1) - left
          sb.append(
            s"""    <rect x="$left" y="$y" width="$width" height="$rowHeight" fill="none"$gridStroke/>\n"""
          )
        }
    }
    sb.append("  </g>\n")

    // Rebuild SVG with correct structure: header -> styles -> defs -> content
    // Extract the header and styles portion
    val content = sb.toString
    val styleEndMarker = "</style>\n"
    val styleEndIdx = content.indexOf(styleEndMarker) + styleEndMarker.length

    val result = new StringBuilder
    result.append(content.substring(0, styleEndIdx))

    // Insert defs with clip paths before content
    result.append("  <defs>\n")
    result.append(clipPathBuffer)
    result.append("  </defs>\n")

    // Append remaining content (headers and cells)
    result.append(content.substring(styleEndIdx))

    // Pass 2: Cell borders (rendered above backgrounds, below text), one line per
    // resolved grid edge run
    result.append("  <g class=\"cell-borders\">\n")
    result.append(
      renderResolvedBorders(hEdges, vEdges, colBoundaries, rowBoundaries, startCol, startRow, theme)
    )
    result.append("  </g>\n")

    // Pass 3: Cell text (rendered on top of all backgrounds and borders, clipped to cell boundaries)
    result.append("  <g class=\"cell-text-layer\">\n")
    result.append(textBuffer)
    result.append("  </g>\n")

    result.append("</svg>")
    result.toString

  /** A cell's own drawing box in pixels: its column, or its merge. */
  private final case class CellBox(x: Int, y: Int, width: Int, height: Int)

  /**
   * The text of one cell, anchored in its own `box` and clipped to `clipId` — the box, or the
   * columns an overflowing text spills across.
   */
  private def cellText(
    resolved: ResolvedCell,
    sheet: Sheet,
    box: CellBox,
    clipId: String,
    includeStyles: Boolean,
    theme: ThemePalette
  ): String =
    val out = new StringBuilder
    val cell = resolved.cell
    val style = resolved.style
    val cellPt = style.fold(DefaultFontSize.toDouble)(_.font.sizePt)
    val (textX, anchor) = textAlignment(resolved, box.x, box.width)
    // Rich text shares one baseline: its largest run must fit the row like Excel autofits it
    val drawnPt = if includeStyles then largestFontPt(cell.value, cellPt) else cellPt
    val textY = textYPosition(style, drawnPt, box.y, box.height)
    val clipAttr = s""" clip-path="url(#$clipId)""""

    cell.value match
      case CellValue.RichText(rt) if includeStyles && rt.runs.nonEmpty =>
        // Get cell's base font for inheritance by unstyled runs
        val baseFont = style.map(_.font)

        // Inter-run gap to account for AWT vs SVG font metric differences
        // SVG renders slightly wider than AWT measures, so add extra spacing
        val interRunGap = InterRunGapPx

        // Calculate total width of all runs (including gaps) for proper alignment
        val runWidths =
          rt.runs.map(run => measureTextWidth(run.text, run.font.orElse(baseFont)))
        val totalGaps = interRunGap * math.max(0, rt.runs.size - 1)
        val totalWidth = runWidths.sum + totalGaps

        // Adjust starting x based on alignment (anchor doesn't work with explicit tspan x)
        val adjustedTextX = anchor match
          case "middle" => textX - totalWidth / 2
          case "end" => textX - totalWidth
          case _ => textX // "start"

        // Explicit base font attrs (was CSS .cell-text): tspan runs override per-run
        out.append(
          s"""    <text y="$textY" class="cell-text" font-size="15px" font-family="Calibri"$clipAttr>"""
        )
        rt.runs.zipWithIndex.foldLeft(adjustedTextX) { case (currentX, (run, idx)) =>
          val runStyle = runToSvgStyle(run, baseFont, theme)
          val escapedText = escapeXml(run.text)
          out.append(s"""<tspan x="$currentX"$runStyle>$escapedText</tspan>""")
          val gap = if idx < rt.runs.size - 1 then interRunGap else 0
          currentX + measureTextWidth(run.text, run.font.orElse(baseFont)) + gap
        }
        out.append("</text>\n")

      case _ =>
        val formatted = resolved.content.text
        if formatted.nonEmpty then
          // includeStyles=false still needs explicit font attrs now that the
          // .cell-text CSS rule no longer declares them (GH-255 cascade fix)
          val textStyle =
            if includeStyles then cellTextStyle(cell, sheet, theme)
            else """ fill="#000000" font-size="15px" font-family="Calibri""""
          val shouldWrap = style.exists(_.align.wrapText)

          // Excel's #### marker: a clipped numeral reads as a different, plausible
          // number (GH-459)
          val text = hashOverflowText(resolved.content, style, box.width).getOrElse(formatted)

          if shouldWrap then
            val availableWidth = box.width - CellPaddingX * 2
            val font = style.map(_.font)
            val lines = wrapText(text, availableWidth, font)
            // The line pitch autofit sizes the row by, so n lines fill an autofitted row
            val lh = lineHeightPx(sheet, cellPt)
            val firstLineY = textYPositionWrapped(style, cellPt, box.y, box.height, lines.size, lh)

            out.append(
              s"""    <text x="$textX" text-anchor="$anchor" class="cell-text"$textStyle$clipAttr>"""
            )
            lines.zipWithIndex.foreach { (line, idx) =>
              val lineY = firstLineY + idx * lh
              val escapedLine = escapeXml(line)
              out.append(s"""<tspan x="$textX" y="$lineY">$escapedLine</tspan>""")
            }
            out.append("</text>\n")
          else
            // The right anchor sits CellPaddingX inside the cell, so text measuring in
            // (width - CellPaddingX, width] would start left of the cell and lose its
            // leading glyph — a sheared digit or, worse, a dropped minus sign (GH-459).
            // Clamp the anchor to the cell's left edge for exactly that band.
            //
            // The clamp must NOT fire once the text is wider than the cell: Excel
            // right-anchors an overflowing right-aligned value, spills it left into empty
            // neighbours and cuts its head at the first one that blocks. Only Text reaches
            // here at full width — Numeric, Bool and Error hash above (GH-459, GH-500), and
            // Bool and Error are centred anyway — and pulling its anchor right would show the
            // head and push the tail out of the cell (it would also contradict the RichText
            // branch, which left-shifts by the run width).
            //
            // Precision note: the clamp lands the left edge exactly on the cell edge, so it
            // relies on the rasterizer laying the string out no wider than AWT measured it.
            // Batik can run ~2px wider than AWT's stringWidth, which can still shave the flag
            // off a leading '1' when the slack is under ~2px (rsvg-convert and resvg do not).
            // A blind safety margin is not the fix — it would only move the shave to the
            // trailing digit. Tracked as GH-505.
            val textW = measureTextWidth(text, style.map(_.font))
            val anchorX =
              if anchor == "end" && textW <= box.width then math.max(textX, box.x + textW)
              else textX
            val escapedText = escapeXml(text)
            out.append(
              s"""    <text x="$anchorX" y="$textY" text-anchor="$anchor" class="cell-text"$textStyle$clipAttr>"""
            )
            out.append(s"""$escapedText</text>\n""")
    out.toString

  /**
   * Get SVG fill attribute for a cell's background. Borders are rendered separately as line
   * elements.
   */
  private def cellStyleToSvg(
    cell: Cell,
    sheet: Sheet,
    theme: ThemePalette
  ): Option[String] =
    cell.styleId.flatMap(sheet.styleRegistry.get).map { style =>
      style.fill match
        case Fill.Solid(color) => colorToFillAttrsWithOpacity(color, theme)
        case Fill.Pattern(_, bgColor, _) =>
          // For pattern fills, use the background color as the cell fill (pattern rendering with
          // the foreground is not yet supported in SVG); an automatic background is the window
          // colour, rendered like Fill.None (GH-566)
          bgColor.map(colorToFillAttrsWithOpacity(_, theme)).getOrElse("""fill="#FFFFFF"""")
        case Fill.None => """fill="#FFFFFF""""
    }

  /**
   * A border declaration competing for one unit grid edge.
   *
   * `fromTrailing` marks the declaring cell as the trailing cell of the edge (right of a vertical
   * edge, below a horizontal edge); it breaks weight ties and orients double-border offsets.
   */
  private final case class EdgeDecl(side: BorderSide, fromTrailing: Boolean)

  /**
   * Visual weight for shared-edge resolution (GH-298), ordered: none < hair < thin family
   * (thin/dotted/dashed/dash-dot/dash-dot-dot) < medium family (medium + medium-dashed variants +
   * slant) < thick < double.
   */
  private def borderWeight(style: BorderStyle): Int =
    style match
      case BorderStyle.None => 0
      case BorderStyle.Hair => 1
      case BorderStyle.Thin | BorderStyle.Dotted | BorderStyle.Dashed | BorderStyle.DashDot |
          BorderStyle.DashDotDot =>
        2
      case BorderStyle.Medium | BorderStyle.MediumDashed | BorderStyle.MediumDashDot |
          BorderStyle.MediumDashDotDot | BorderStyle.SlantDashDot =>
        3
      case BorderStyle.Thick => 4
      case BorderStyle.Double => 5

  /**
   * Excel's shared-edge rule: the heavier border wins; ties go to the trailing (right/bottom)
   * cell's declaration, matching Excel's drawing order.
   */
  private def resolveSharedEdge(a: EdgeDecl, b: EdgeDecl): EdgeDecl =
    val aw = borderWeight(a.side.style)
    val bw = borderWeight(b.side.style)
    if bw > aw then b
    else if aw > bw then a
    else if b.fromTrailing then b
    else a

  /**
   * Decompose a cell's border into unit grid edges of its effective rect: the perimeter spans
   * `spanCols` columns and `spanRows` rows starting at (col, row). Top/left declarations come from
   * the trailing cell of their edges; bottom/right from the leading cell.
   */
  private def declareCellEdges(
    hEdges: scala.collection.mutable.Map[(Int, Int), EdgeDecl],
    vEdges: scala.collection.mutable.Map[(Int, Int), EdgeDecl],
    border: Border,
    col: Int,
    row: Int,
    spanCols: Int,
    spanRows: Int
  ): Unit =
    val lastCol = col + spanCols - 1
    val lastRow = row + spanRows - 1
    (col to lastCol).foreach { c =>
      declareEdge(hEdges, (row, c), border.top, fromTrailing = true)
      declareEdge(hEdges, (lastRow + 1, c), border.bottom, fromTrailing = false)
    }
    (row to lastRow).foreach { r =>
      declareEdge(vEdges, (col, r), border.left, fromTrailing = true)
      declareEdge(vEdges, (lastCol + 1, r), border.right, fromTrailing = false)
    }

  private def declareEdge(
    edges: scala.collection.mutable.Map[(Int, Int), EdgeDecl],
    key: (Int, Int),
    side: BorderSide,
    fromTrailing: Boolean
  ): Unit =
    if side.style != BorderStyle.None then
      val incoming = EdgeDecl(side, fromTrailing)
      edges(key) = edges.get(key) match
        case Some(existing) => resolveSharedEdge(existing, incoming)
        case None => incoming

  /**
   * Render the resolved border edges as line elements, one `<line>` per maximal run of contiguous
   * unit edges with identical winning declarations. Zero-length segments (hidden rows/columns) are
   * dropped.
   *
   * Unlike a rect stroke approach, line elements allow different styles/colors per side and support
   * dashed, dotted, and double borders.
   */
  private def renderResolvedBorders(
    hEdges: scala.collection.Map[(Int, Int), EdgeDecl],
    vEdges: scala.collection.Map[(Int, Int), EdgeDecl],
    colBoundaries: IndexedSeq[Int],
    rowBoundaries: IndexedSeq[Int],
    startCol: Int,
    startRow: Int,
    theme: ThemePalette
  ): String =
    val sb = new StringBuilder
    edgeRuns(hEdges).foreach { case (b, c0, c1, decl) =>
      val y = rowBoundaries(b - startRow)
      val x1 = colBoundaries(c0 - startCol)
      val x2 = colBoundaries(c1 + 1 - startCol)
      if x1 != x2 then
        val side = if decl.fromTrailing then "top" else "bottom"
        sb.append(renderBorderLine(side, decl.side, x1, y, x2, y, theme))
    }
    edgeRuns(vEdges).foreach { case (b, r0, r1, decl) =>
      val x = colBoundaries(b - startCol)
      val y1 = rowBoundaries(r0 - startRow)
      val y2 = rowBoundaries(r1 + 1 - startRow)
      if y1 != y2 then
        val side = if decl.fromTrailing then "left" else "right"
        sb.append(renderBorderLine(side, decl.side, x, y1, x, y2, theme))
    }
    sb.toString

  /**
   * Sorted maximal runs `(boundary, first, last, decl)` of contiguous unit edges sharing a boundary
   * line and an identical declaration.
   */
  private def edgeRuns(
    edges: scala.collection.Map[(Int, Int), EdgeDecl]
  ): List[(Int, Int, Int, EdgeDecl)] =
    edges.toList
      .map { case ((b, u), decl) => (b, u, decl) }
      .sortBy { case (b, u, _) => (b, u) }
      .foldLeft(List.empty[(Int, Int, Int, EdgeDecl)]) {
        case ((rb, r0, r1, rd) :: rest, (b, u, d)) if b == rb && u == r1 + 1 && d == rd =>
          (rb, r0, u, rd) :: rest
        case (acc, (b, u, d)) => (b, u, u, d) :: acc
      }
      .reverse

  /**
   * Render a single border line. Handles double borders specially.
   */
  private def renderBorderLine(
    side: String,
    borderSide: BorderSide,
    x1: Double,
    y1: Double,
    x2: Double,
    y2: Double,
    theme: ThemePalette
  ): String =
    if borderSide.style == BorderStyle.Double then
      renderDoubleBorder(side, borderSide, x1, y1, x2, y2, theme)
    else
      val (strokeWidth, dashArray) = borderStyleToSvg(borderSide.style)
      val color = borderSide.color.map(c => colorToHex(c, theme)).getOrElse("#000000")
      val dashAttr = dashArray.map(d => s""" stroke-dasharray="$d"""").getOrElse("")
      s"""      <line x1="$x1" y1="$y1" x2="$x2" y2="$y2" stroke="$color" stroke-width="$strokeWidth"$dashAttr/>\n"""

  /**
   * Render a double border as two parallel lines.
   */
  private def renderDoubleBorder(
    side: String,
    borderSide: BorderSide,
    x1: Double,
    y1: Double,
    x2: Double,
    y2: Double,
    theme: ThemePalette
  ): String =
    val color = borderSide.color.map(c => colorToHex(c, theme)).getOrElse("#000000")
    val offset = 2.0 // Gap between double lines

    side match
      case "top" =>
        s"""      <line x1="$x1" y1="$y1" x2="$x2" y2="$y1" stroke="$color" stroke-width="1"/>
      <line x1="$x1" y1="${y1 + offset}" x2="$x2" y2="${y1 + offset}" stroke="$color" stroke-width="1"/>
"""
      case "bottom" =>
        s"""      <line x1="$x1" y1="${y1 - offset}" x2="$x2" y2="${y1 - offset}" stroke="$color" stroke-width="1"/>
      <line x1="$x1" y1="$y1" x2="$x2" y2="$y1" stroke="$color" stroke-width="1"/>
"""
      case "left" =>
        s"""      <line x1="$x1" y1="$y1" x2="$x1" y2="$y2" stroke="$color" stroke-width="1"/>
      <line x1="${x1 + offset}" y1="$y1" x2="${x1 + offset}" y2="$y2" stroke="$color" stroke-width="1"/>
"""
      case "right" =>
        s"""      <line x1="${x1 - offset}" y1="$y1" x2="${x1 - offset}" y2="$y2" stroke="$color" stroke-width="1"/>
      <line x1="$x1" y1="$y1" x2="$x1" y2="$y2" stroke="$color" stroke-width="1"/>
"""
      case _ => ""

  /**
   * Map BorderStyle to SVG stroke attributes (width and dash pattern).
   */
  private def borderStyleToSvg(style: BorderStyle): (Int, Option[String]) =
    style match
      case BorderStyle.None => (0, None)
      case BorderStyle.Thin => (1, None)
      case BorderStyle.Medium => (2, None)
      case BorderStyle.Thick => (3, None)
      case BorderStyle.Hair => (1, Some("1,1")) // Very fine dotted
      case BorderStyle.Dotted => (1, Some("2,2")) // Dotted
      case BorderStyle.Dashed => (1, Some("4,2")) // Dashed
      case BorderStyle.MediumDashed => (2, Some("4,2")) // Medium dashed
      case BorderStyle.DashDot => (1, Some("4,2,1,2")) // Dash-dot
      case BorderStyle.MediumDashDot => (2, Some("4,2,1,2"))
      case BorderStyle.DashDotDot => (1, Some("4,2,1,2,1,2"))
      case BorderStyle.MediumDashDotDot => (2, Some("4,2,1,2,1,2"))
      case BorderStyle.SlantDashDot => (1, Some("4,2,1,2"))
      case BorderStyle.Double => (1, None) // Handled separately

  /**
   * Get SVG text style attributes for a cell.
   *
   * ALWAYS includes font properties explicitly for exact fidelity (don't rely on CSS defaults).
   */
  private def cellTextStyle(cell: Cell, sheet: Sheet, theme: ThemePalette): String =
    cell.styleId
      .flatMap(sheet.styleRegistry.get)
      .map { style =>
        val attrs = scala.collection.mutable.ArrayBuffer[String]()

        // ALWAYS include font color (even if default black) for exact fidelity
        val fontColor = style.font.color.map(c => colorToHex(c, theme)).getOrElse("#000000")
        attrs += s"""fill="$fontColor""""

        // Font weight
        if style.font.bold then attrs += """font-weight="bold""""

        // Font style
        if style.font.italic then attrs += """font-style="italic""""

        // Underline (SVG uses text-decoration, GH-256)
        if style.font.underline != Underline.None then attrs += """text-decoration="underline""""

        // ALWAYS include font size (don't rely on CSS defaults) - convert pt to px (pt * 4/3)
        val fontSizePx = (style.font.sizePt * 4.0 / 3.0).toInt
        attrs += s"""font-size="${fontSizePx}px""""

        // ALWAYS include font family (even if default) - unquoted in SVG attributes (GH-255)
        attrs += s"""font-family="${svgFontFamily(style.font.name)}""""

        " " + attrs.mkString(" ")
      }
      .getOrElse {
        // No style - use explicit defaults for exact fidelity
        """ fill="#000000" font-size="15px" font-family="Calibri""""
      }

  /**
   * Convert a TextRun to SVG tspan style attributes.
   *
   * When a run has no explicit font, inherits from the cell's base font. This ensures rich text
   * runs without explicit formatting still display with the cell's styling.
   */
  private def runToSvgStyle(run: TextRun, baseFont: Option[Font], theme: ThemePalette): String =
    // Use run's font if present, otherwise fall back to base font
    val effectiveFont = run.font.orElse(baseFont)

    effectiveFont match
      case None => ""
      case Some(f) =>
        val attrs = scala.collection.mutable.ArrayBuffer[String]()

        // Font color - always include for exact fidelity
        val colorAttr = f.color.map(c => colorToHex(c, theme)).getOrElse("#000000")
        attrs += s"""fill="$colorAttr""""

        // Font weight
        if f.bold then attrs += """font-weight="bold""""

        // Font style
        if f.italic then attrs += """font-style="italic""""

        // Underline (SVG uses text-decoration)
        if f.underline != Underline.None then attrs += """text-decoration="underline""""

        // Font size - always include for exact fidelity - convert pt to px (pt * 4/3)
        val fontSizePx = (f.sizePt * 4.0 / 3.0).toInt
        attrs += s"""font-size="${fontSizePx}px""""

        // Font family - always include for exact fidelity - unquoted in SVG attributes (GH-255)
        attrs += s"""font-family="${svgFontFamily(f.name)}""""

        if attrs.nonEmpty then " " + attrs.mkString(" ") else ""

  /**
   * Calculate text y position based on vertical alignment.
   *
   * SVG text baseline is at the y coordinate, so we need to adjust for font metrics. Approximation:
   * ascender ~ 80% of font size.
   */
  private def textYPosition(
    style: Option[CellStyle],
    fontSizePt: Double,
    cellY: Int,
    cellHeight: Int
  ): Int =
    textYPositionWrapped(style, fontSizePt, cellY, cellHeight, lineCount = 1, lh = 0)

  /**
   * Calculate y position for the first line of wrapped text.
   *
   * Accounts for vertical alignment and total text height (lineCount * lineHeight). A
   * bottom-aligned last line sits [[RenderUtils.baselineInsetPx]] above the bottom edge, so a large
   * font's descenders stay inside its autofitted row.
   */
  private def textYPositionWrapped(
    style: Option[CellStyle],
    fontSizePt: Double,
    cellY: Int,
    cellHeight: Int,
    lineCount: Int,
    lh: Int
  ): Int =
    val vAlign = style.map(_.align.vertical).getOrElse(VAlign.Bottom)
    val fontSizePx = (fontSizePt * 4.0 / 3.0).toInt // Convert pt to px

    // Text baseline adjustment (SVG places text at baseline, not top)
    val baselineOffset = (fontSizePx * 0.8).toInt // Approximate ascender height

    // Total height of wrapped text (lineCount - 1 because first line doesn't have spacing above it)
    val totalTextHeight = if lineCount > 1 then (lineCount - 1) * lh + fontSizePx else fontSizePx
    val bottomBaseline = cellY + cellHeight - baselineInsetPx(fontSizePt)

    vAlign match
      case VAlign.Top =>
        cellY + baselineOffset + 4 // Small padding from top
      case VAlign.Middle =>
        // Center the text block vertically
        if lineCount > 1 then cellY + (cellHeight - totalTextHeight) / 2 + baselineOffset
        else cellY + cellHeight / 2 + baselineOffset / 3
      case VAlign.Bottom =>
        // Position last line near bottom, calculate where first line should start
        if lineCount > 1 then bottomBaseline - (lineCount - 1) * lh
        else bottomBaseline
      case VAlign.Justify | VAlign.Distributed =>
        // Same as Middle for wrapped text
        if lineCount > 1 then cellY + (cellHeight - totalTextHeight) / 2 + baselineOffset
        else cellY + cellHeight / 2 + baselineOffset / 3

  /**
   * Calculate text x position and anchor based on alignment and indentation.
   *
   * Indentation adds IndentPxPerLevel pixels per level (Excel uses ~3 characters per level).
   */
  private def textAlignment(cell: ResolvedCell, cellX: Int, cellWidth: Int): (Int, String) =
    val indent = cell.style.map(_.align.indent).getOrElse(0)
    val indentPx = indent * IndentPxPerLevel

    // General resolves from the rendered kind, the same way the hash and colspan decisions do:
    // errors centre like logicals, a number under a text-only format is text (GH-500, GH-501)
    cell.align match
      case HAlign.Center | HAlign.CenterContinuous =>
        // Center alignment: indent shifts content slightly right
        (cellX + cellWidth / 2 + indentPx / 2, "middle")
      case HAlign.Right =>
        // Right alignment: indent typically ignored (Excel behavior)
        (cellX + cellWidth - CellPaddingX, "end")
      case _ =>
        // Left alignment: indent adds to left padding
        (cellX + CellPaddingX + indentPx, "start")
