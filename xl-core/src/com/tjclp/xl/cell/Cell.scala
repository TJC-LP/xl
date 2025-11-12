package com.tjclp.xl.cell

import com.tjclp.xl.addressing.{ARef, Column, Row}
import com.tjclp.xl.style.units.StyleId

object Cell:
  /** Create cell from reference and value */
  def apply(ref: ARef, value: CellValue): Cell = Cell(ref, value, None, None, None)

  /** Create cell from A1 notation and value */
  def fromA1(a1: String, value: CellValue): Either[String, Cell] =
    ARef.parse(a1).map(ref => Cell(ref, value))

  /** Create empty cell */
  def empty(ref: ARef): Cell = Cell(ref, CellValue.Empty)

/** Cell with value and optional metadata */
case class Cell(
  ref: ARef,
  value: CellValue = CellValue.Empty,
  styleId: Option[StyleId] = None,
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

  /** Add comment to cell */
  def withComment(text: String): Cell = copy(comment = Some(text))

  /** Clear comment */
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
    case CellValue.Formula(_) => true
    case _ => false

  /** Check if cell contains an error */
  def isError: Boolean = value match
    case CellValue.Error(_) => true
    case _ => false
