package com.tjclp.xl.formula

import com.tjclp.xl.CellRange
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.TExpr.*

trait TExprLookupOps:
  /**
   * Smart constructor for VLOOKUP (supports text and numeric lookups).
   *
   * Example: TExpr.vlookup(TExpr.Lit("Widget A"), CellRange("A1:D10"), TExpr.Lit(4),
   * TExpr.Lit(false))
   */
  def vlookup(
    lookup: TExpr[?],
    table: CellRange,
    colIndex: TExpr[Int],
    rangeLookup: TExpr[Boolean] = Lit(true)
  ): TExpr[CellValue] =
    Call(
      FunctionSpecs.vlookup,
      (asCellValueExpr(lookup), RangeLocation.Local(table), colIndex, Some(rangeLookup))
    )

  /**
   * Smart constructor for VLOOKUP with explicit RangeLocation (supports cross-sheet lookups).
   *
   * Example: TExpr.vlookupWithLocation(TExpr.Lit("Widget A"),
   * RangeLocation.CrossSheet(SheetName("Lookup"), CellRange("A1:D10")), TExpr.Lit(4),
   * TExpr.Lit(false))
   */
  def vlookupWithLocation(
    lookup: TExpr[?],
    table: RangeLocation,
    colIndex: TExpr[Int],
    rangeLookup: TExpr[Boolean] = Lit(true)
  ): TExpr[CellValue] =
    Call(
      FunctionSpecs.vlookup,
      (asCellValueExpr(lookup), table, colIndex, Some(rangeLookup))
    )

  /**
   * XLOOKUP: advanced lookup with flexible matching.
   *
   * @param lookupValue
   *   The value to search for
   * @param lookupArray
   *   The range to search in
   * @param returnArray
   *   The range to return values from (same dimensions as lookupArray)
   * @param ifNotFound
   *   Optional value to return if no match (default: #N/A error)
   * @param matchMode
   *   0=exact (default), -1=next smaller, 1=next larger, 2=wildcard
   * @param searchMode
   *   1=first-to-last (default), -1=last-to-first, 2=binary asc, -2=binary desc
   *
   * Example: TExpr.xlookup(TExpr.Lit("Apple"), lookupRange, returnRange)
   */
  def xlookup(
    lookupValue: TExpr[?],
    lookupArray: CellRange,
    returnArray: CellRange,
    ifNotFound: Option[TExpr[?]] = None,
    matchMode: TExpr[Int] = Lit(0),
    searchMode: TExpr[Int] = Lit(1)
  ): TExpr[CellValue] =
    val matchModeOpt = ifNotFound.map(_ => matchMode)
    val searchModeOpt = ifNotFound.map(_ => searchMode)
    Call(
      FunctionSpecs.xlookup,
      (
        lookupValue.asInstanceOf[TExpr[Any]],
        lookupArray,
        returnArray,
        ifNotFound.map(_.asInstanceOf[TExpr[Any]]),
        matchModeOpt,
        searchModeOpt
      )
    )

  /**
   * INDEX: get value at position in array.
   *
   * @param array
   *   The range to index into
   * @param rowNum
   *   1-based row position
   * @param colNum
   *   Optional 1-based column position (defaults to 1 for single-column ranges)
   *
   * Example: TExpr.index(range, TExpr.Lit(2), Some(TExpr.Lit(3)))
   */
  def index(
    array: CellRange,
    rowNum: TExpr[BigDecimal],
    colNum: Option[TExpr[BigDecimal]] = None
  ): TExpr[CellValue] =
    Call(FunctionSpecs.index, (array, rowNum, colNum))

  /**
   * MATCH: find position of value in array.
   *
   * @param lookupValue
   *   The value to search for
   * @param lookupArray
   *   The range to search in
   * @param matchType
   *   1=largest <= (default), 0=exact, -1=smallest >=
   *
   * Example: TExpr.matchExpr(TExpr.Lit("B"), range, TExpr.Lit(0))
   */
  def matchExpr(
    lookupValue: TExpr[?],
    lookupArray: CellRange,
    matchType: TExpr[BigDecimal] = Lit(BigDecimal(1))
  ): TExpr[BigDecimal] =
    Call(
      FunctionSpecs.matchFn,
      (lookupValue.asInstanceOf[TExpr[Any]], lookupArray, Some(matchType))
    )
