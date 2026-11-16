package com.tjclp.xl.formula.perf

import java.util.concurrent.TimeUnit

import com.tjclp.xl.addressing.{ARef, SheetName}
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.Clock
import com.tjclp.xl.formula.eval.Evaluator
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.sheets.Sheet
import org.openjdk.jmh.annotations.*

/** #735: use -prof gc to measure allocation as well as elapsed time. */
@State(Scope.Benchmark)
@BenchmarkMode(Array(Mode.AverageTime))
@OutputTimeUnit(TimeUnit.MILLISECONDS)
@Warmup(iterations = 2)
@Measurement(iterations = 3)
@Fork(value = 1)
class PositionSequenceBenchmark:
  private val empty = Sheet(SheetName.unsafe("Positions"))
  private val sheet = (0 until 100000).foldLeft(empty) { (s, row) =>
    s.put(ARef.from0(0, row), CellValue.Number(BigDecimal(1)))
  }
  private val positions = FormulaParser.parse("=ROW(A:A)")
  private val whole = FormulaParser.parse("=SUM(ROW(A:A))")
  private val weighted = FormulaParser.parse("=SUMPRODUCT((A1:A100000>0)*ROW(A1:A100000))")

  @Benchmark
  def wholeColumnPositions(): Any =
    positions.flatMap(expr => Evaluator.arrayInstance.eval(expr, empty, Clock.system, None, None))

  @Benchmark
  def wholeColumnSum(): Any =
    whole.flatMap(expr => Evaluator.arrayInstance.eval(expr, empty, Clock.system, None, None))

  @Benchmark
  def weightedPositions(): Any =
    weighted.flatMap(expr => Evaluator.arrayInstance.eval(expr, sheet, Clock.system, None, None))
