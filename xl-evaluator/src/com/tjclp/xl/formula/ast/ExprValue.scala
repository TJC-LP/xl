package com.tjclp.xl.formula.ast

import com.tjclp.xl.formula.functions.FunctionSpecs
import com.tjclp.xl.formula.eval.EvalError
import com.tjclp.xl.formula.functions.EvalContext

import com.tjclp.xl.cells.{CellError, CellValue}
import java.time.{LocalDate, LocalDateTime}

sealed trait ExprValue derives CanEqual

object ExprValue:
  final case class Number(value: BigDecimal) extends ExprValue
  final case class Text(value: String) extends ExprValue
  final case class Bool(value: Boolean) extends ExprValue
  final case class Date(value: LocalDate) extends ExprValue
  final case class DateTime(value: LocalDateTime) extends ExprValue
  final case class Cell(value: CellValue) extends ExprValue
  final case class Opaque(value: Any) extends ExprValue

  def from(value: Any): ExprValue = value match
    case exprValue: ExprValue => exprValue
    case cellValue: CellValue => Cell(cellValue)
    case text: String => Text(text)
    case number: BigDecimal => Number(number)
    case number: Int => Number(BigDecimal(number))
    case number: Long => Number(BigDecimal(number))
    case number: Double if number.isFinite => Number(BigDecimal(number))
    // #681: NaN and ±Infinity have no Excel value (BigDecimal(Double) throws on them) — #NUM!
    case _: Double => Cell(CellValue.Error(CellError.Num))
    case bool: Boolean => Bool(bool)
    case date: LocalDate => Date(date)
    case dateTime: LocalDateTime => DateTime(dateTime)
    case other => Opaque(other)
