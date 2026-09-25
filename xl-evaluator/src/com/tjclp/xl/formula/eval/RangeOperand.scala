package com.tjclp.xl.formula.eval

import com.tjclp.xl.addressing.CellRange
import com.tjclp.xl.sheets.Sheet

/**
 * A reference passed as an argument in Excel's reference operand class (an aggregate's, AND's,
 * OR's) from a plain cell: the range itself, unread, so the consumer folds it as it folds a range
 * written in its argument — a whole column bounded to the data, text and blanks skipped by Excel's
 * reference rule (`SUM(IF(A1:A10>2,A1:A10,0))`). Only
 * [[com.tjclp.xl.formula.functions.EvalContext.evalReferenceArg]] produces one, for its consumers.
 */
private[formula] final case class RangeOperand(sheet: Sheet, range: CellRange)
