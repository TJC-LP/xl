package com.tjclp.xl.ooxml

import com.tjclp.xl.addressing.ARef

/**
 * VML drawing generator for Excel comment indicators.
 *
 * Generates legacy VML (Vector Markup Language) XML for comment visual indicators (yellow
 * triangles). VML is deprecated but still required by Excel for traditional "Notes" comments.
 *
 * **Architecture**:
 *   - Comments content: `xl/commentsN.xml` (modern XML)
 *   - Visual indicators: `xl/drawings/vmlDrawingN.vml` (legacy VML) ← This module
 *
 * **VML Structure**:
 *   - `<o:shapelayout>` - Shape layout metadata (one per file)
 *   - `<v:shapetype>` - Template for comment boxes (one per file)
 *   - `<v:shape>` - Individual comment indicator (one per comment)
 *
 * VML is NOT parsed on read (binary passthrough). This module generates VML for new comments only.
 *
 * @since 0.4.0
 */
object VmlDrawing:

  /**
   * VML shapetype template for comment boxes.
   *
   * Fixed template used by all comment shapes. Defines the comment box appearance (rectangular
   * callout with connector).
   */
  private val SHAPE_TYPE_TEMPLATE: String =
    """<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"
      |             path="m,l,21600r21600,l21600,xe">
      |  <v:stroke joinstyle="miter"/>
      |  <v:path gradientshapeok="t" o:connecttype="rect"/>
      | </v:shapetype>""".stripMargin

  /**
   * Generate complete VML XML for a sheet's comment indicators.
   *
   * Creates VML file with:
   *   - Shape layout header
   *   - Shared shapetype template
   *   - Individual shapes for each comment
   *
   * @param comments
   *   OOXML comments for this sheet
   * @param sheetIndex
   *   Sheet index (0-based)
   * @return
   *   Complete VML XML string (compact format, no pretty-printing)
   */
  def generateForComments(comments: OoxmlComments, sheetIndex: Int): String =
    val baseShapeId = 1024 + (sheetIndex * 100) // Unique IDs per sheet
    val shapes = comments.comments.zipWithIndex.map { case (comment, idx) =>
      generateShape(comment, baseShapeId + idx)
    }

    s"""<xml xmlns:v="urn:schemas-microsoft-com:vml"
       |     xmlns:o="urn:schemas-microsoft-com:office:office"
       |     xmlns:x="urn:schemas-microsoft-com:office:excel">
       | <o:shapelayout v:ext="edit">
       |  <o:idmap v:ext="edit" data="${sheetIndex + 1}"/>
       | </o:shapelayout>
       | $SHAPE_TYPE_TEMPLATE
       | ${shapes.mkString("\n ")}
       |</xml>""".stripMargin

  /**
   * Generate VML shape for a single comment indicator.
   *
   * Creates a shape element with:
   *   - Unique ID
   *   - Fixed yellow styling (`#ffffe1`)
   *   - Cell anchoring via ClientData
   *   - Hidden visibility (shows on hover)
   *
   * Position calculation uses simplified heuristics - Excel recalculates exact position on open.
   *
   * @param comment
   *   OOXML comment to generate shape for
   * @param shapeId
   *   Unique shape ID (must be unique across file)
   * @return
   *   VML shape XML string
   */
  private def generateShape(comment: OoxmlComment, shapeId: Int): String =
    val ref = comment.ref
    val col = ref.col.index0 // 0-based
    val row = ref.row.index0 // 0-based

    // Defensive validation (ARef already enforces, but defense-in-depth)
    require(col >= 0 && col <= 16383, s"Column $col exceeds Excel maximum (16383)")
    require(row >= 0 && row <= 1048575, s"Row $row exceeds Excel maximum (1048575)")

    // Simplified positioning (Excel adjusts on open)
    val marginLeft = col * 75.0 + 59.25 // pt units
    val marginTop = row * 15.0 + 1.5 // pt units

    // Anchor: [fromCol, fromColOffset, fromRow, fromRowOffset, toCol, toColOffset, toRow, toRowOffset]
    // Use fixed offsets - Excel adjusts for optimal display
    val anchor = s"$col, 15, $row, 2, ${col + 2}, 15, ${row + 3}, 16"

    s"""<v:shape id="_x0000_s$shapeId" type="#_x0000_t202"
       |          style="position:absolute;margin-left:${marginLeft}pt;margin-top:${marginTop}pt;
       |                 width:108pt;height:59.25pt;z-index:$shapeId;visibility:hidden"
       |          fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">
       |  <v:fill color2="infoBackground [80]"/>
       |  <v:shadow color="none [81]" obscured="t"/>
       |  <v:path o:connecttype="none"/>
       |  <v:textbox style="mso-direction-alt:auto">
       |   <div style="text-align:left"></div>
       |  </v:textbox>
       |  <x:ClientData ObjectType="Note">
       |   <x:MoveWithCells/>
       |   <x:SizeWithCells/>
       |   <x:Anchor>$anchor</x:Anchor>
       |   <x:AutoFill>False</x:AutoFill>
       |   <x:Row>$row</x:Row>
       |   <x:Column>$col</x:Column>
       |  </x:ClientData>
       | </v:shape>""".stripMargin
