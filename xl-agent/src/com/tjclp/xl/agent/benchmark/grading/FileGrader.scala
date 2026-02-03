package com.tjclp.xl.agent.benchmark.grading

import cats.effect.IO
import cats.syntax.all.*
import com.tjclp.xl.agent.benchmark.{ComparableValue, Evaluator, RangeResult}

import java.nio.file.{Files, Path}

// ============================================================================
// File-Based Grader (SpreadsheetBench-style)
// ============================================================================

/** Grader that compares output files against expected answer files */
final class SpreadsheetBenchGrader(xlCliPath: String) extends Grader[Score.BinaryScore]:

  val name: String = "file"

  def canHandle(ctx: GraderContext): Boolean =
    ctx.outputPath.isDefined &&
      ctx.answerPath.isDefined &&
      ctx.answerPosition.isDefined

  def grade(ctx: GraderContext): IO[GradeResult[Score.BinaryScore]] =
    (ctx.outputPath, ctx.answerPath, ctx.answerPosition) match
      case (Some(out), Some(ans), Some(pos)) =>
        gradeInternal(ctx.taskId, out, ans, pos)
      case _ =>
        IO.pure(
          GradeResult(
            graderName = name,
            score = Score.BinaryScore.Fail,
            details = GradeDetails.fail("Missing output, answer, or position")
          )
        )

  private def gradeInternal(
    taskId: String,
    outputPath: Path,
    answerPath: Path,
    answerPosition: String
  ): IO[GradeResult[Score.BinaryScore]] =
    // Check if output file exists
    if !Files.exists(outputPath) then
      IO.pure(
        GradeResult(
          graderName = name,
          score = Score.BinaryScore.Fail,
          details = GradeDetails.fail(s"Output file not found: $outputPath"),
          metadata = Map("taskId" -> taskId)
        )
      )
    else
      // Use the existing Evaluator for comparison
      Evaluator
        .compare(outputPath, answerPath, answerPosition, xlCliPath)
        .map { rangeResults =>
          val passed = rangeResults.forall(_.passed)
          val mismatches = extractMismatches(rangeResults)
          val reasoning =
            if passed then "All cells match expected values"
            else s"${mismatches.length} cell(s) did not match"

          GradeResult(
            graderName = name,
            score = Score.BinaryScore.fromBoolean(passed),
            details = GradeDetails(
              passed = passed,
              reasoning = Some(reasoning),
              mismatches = mismatches
            ),
            metadata = Map(
              "taskId" -> taskId,
              "rangesChecked" -> rangeResults.length.toString,
              "rangesPassed" -> rangeResults.count(_.passed).toString
            )
          )
        }
        .handleError { e =>
          GradeResult(
            graderName = name,
            score = Score.BinaryScore.Fail,
            details = GradeDetails.fail(s"Evaluation error: ${e.getMessage}"),
            metadata = Map("taskId" -> taskId, "error" -> e.getMessage)
          )
        }

  private def extractMismatches(rangeResults: List[RangeResult]): List[CellMismatch] =
    rangeResults.flatMap(_.mismatches).map { m =>
      CellMismatch(
        ref = m.ref,
        expected = formatComparableValue(m.expected),
        actual = formatComparableValue(m.actual)
      )
    }

  private def formatComparableValue(cv: ComparableValue): String =
    cv match
      case ComparableValue.Empty => "<empty>"
      case ComparableValue.Number(v) => v.toString
      case ComparableValue.Text(v) => s"\"$v\""
      case ComparableValue.Bool(v) => v.toString
      case ComparableValue.Error(v) => s"#$v"

object SpreadsheetBenchGrader:
  def apply(xlCliPath: String = "xl"): SpreadsheetBenchGrader =
    new SpreadsheetBenchGrader(xlCliPath)

  def default: SpreadsheetBenchGrader = apply()

// ============================================================================
// Multi-Case File Grader (for tasks with multiple test cases)
// ============================================================================

/** Grade multiple test cases and produce a fractional score */
final class MultiCaseFileGrader(xlCliPath: String) extends Grader[Score.FractionalScore]:

  private val singleGrader = SpreadsheetBenchGrader(xlCliPath)

  val name: String = "multi-case-file"

  def canHandle(ctx: GraderContext): Boolean =
    singleGrader.canHandle(ctx)

  def grade(ctx: GraderContext): IO[GradeResult[Score.FractionalScore]] =
    // This grader is typically called with a single case at a time,
    // with results aggregated externally. For direct use, just wrap single result.
    singleGrader.grade(ctx).map { result =>
      val score =
        if result.passed then Score.FractionalScore(1, 1)
        else Score.FractionalScore(0, 1)
      GradeResult(
        graderName = name,
        score = score,
        details = result.details,
        metadata = result.metadata
      )
    }

  /** Grade multiple contexts and aggregate results */
  def gradeAll(contexts: List[GraderContext]): IO[GradeResult[Score.FractionalScore]] =
    contexts.traverse(singleGrader.grade).map { results =>
      val passing = results.count(_.passed)
      val total = results.length
      val allMismatches = results.flatMap(_.details.mismatches)
      val allReasons = results.flatMap(_.details.reasoning)

      GradeResult(
        graderName = name,
        score = Score.FractionalScore(passing, total),
        details = GradeDetails(
          passed = passing == total && total > 0,
          reasoning = Some(s"$passing/$total test cases passed"),
          mismatches = allMismatches
        ),
        metadata = Map(
          "casesTotal" -> total.toString,
          "casesPassed" -> passing.toString
        )
      )
    }

object MultiCaseFileGrader:
  def apply(xlCliPath: String = "xl"): MultiCaseFileGrader =
    new MultiCaseFileGrader(xlCliPath)

  def default: MultiCaseFileGrader = apply()
