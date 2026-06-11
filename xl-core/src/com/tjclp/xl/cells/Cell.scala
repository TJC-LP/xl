package com.tjclp.xl.cells

import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.styles.units.StyleId

object Cell:
  /** Create cell from reference and value */
  def apply(ref: ARef, value: CellValue): Cell = Cell(ref, value, None, None, None)

  /** Create cell from A1 notation and value */
  def fromA1(a1: String, value: CellValue): Either[String, Cell] =
    ARef.parse(a1).map(ref => Cell(ref, value))

  /** Create empty cell */
  def empty(ref: ARef): Cell = Cell(ref, CellValue.Empty)

/**
 * Cell with value and optional metadata.
 *
 * Note: cell comments live in the sheet-level store (`Sheet.comments`) — that is what the OOXML
 * writer serializes. The [[comment]] field here is deprecated: `Sheet.put(cell)` converts it into
 * `Sheet.comments` (as a plain-text comment) and clears the field, so it never silently vanishes on
 * write (GH-295).
 */
final case class Cell(
  ref: ARef,
  value: CellValue = CellValue.Empty,
  styleId: Option[StyleId] = None,
  @deprecated(
    "use Sheet.comment/Sheet.comments — this field is not serialized; Sheet.put(cell) converts it into Sheet.comments and clears it",
    "0.12.1"
  )
  comment: Option[String] = None,
  hyperlink: Option[String] = None
):
  /** Get cell column */
  def col: Column = this.ref.col

  /** Get cell row */
  def row: Row = this.ref.row

  /** Get A1 notation */
  def toA1: String = this.ref.toA1

  /** Update cell value */
  def withValue(newValue: CellValue): Cell = copy(value = newValue)

  /** Update cell style */
  def withStyle(styleId: StyleId): Cell = copy(styleId = Some(styleId))

  /** Clear cell style */
  def clearStyle: Cell = copy(styleId = None)

  /**
   * Add comment to cell.
   *
   * Prefer `Sheet.comment(ref, Comment.plainText(text))`: the cell-level field is not serialized
   * directly — `Sheet.put(cell)` converts it into the sheet-level comment store (GH-295).
   */
  @deprecated(
    "use Sheet.comment(ref, Comment.plainText(text)) — Cell.comment is not serialized; Sheet.put(cell) converts it into Sheet.comments",
    "0.12.1"
  )
  def withComment(text: String): Cell = copy(comment = Some(text))

  /**
   * Clear comment.
   *
   * Prefer `Sheet.removeComment(ref)`: sheet-level comments are the serialized store (GH-295).
   */
  @deprecated(
    "use Sheet.removeComment(ref) — Cell.comment is not serialized; comments live in Sheet.comments",
    "0.12.1"
  )
  def clearComment: Cell = copy(comment = None)

  /** Add hyperlink to cell */
  def withHyperlink(url: String): Cell = copy(hyperlink = Some(url))

  /** Clear hyperlink */
  def clearHyperlink: Cell = copy(hyperlink = None)

  /** Check if cell is empty */
  def isEmpty: Boolean = value == CellValue.Empty

  /** Check if cell is not empty */
  def nonEmpty: Boolean = !isEmpty

  /** Check if cell contains a formula */
  def isFormula: Boolean = value match
    case CellValue.Formula(_, _) => true
    case _ => false

  /** Check if cell contains an error */
  def isError: Boolean = value match
    case CellValue.Error(_) => true
    case _ => false
