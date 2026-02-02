package com.tjclp.xl.agent

import com.tjclp.xl.agent.benchmark.execution.{CaseDetails, CaseResult, ExecutionResult, TokenUsage}
import com.tjclp.xl.agent.benchmark.grading.{CellMismatch, Score}
import com.tjclp.xl.agent.benchmark.task.TaskId
import org.scalacheck.{Arbitrary, Gen}

/** ScalaCheck generators for xl-agent test data */
object Generators:

  // --------------------------------------------------------------------------
  // Token Usage Generators
  // --------------------------------------------------------------------------

  val genTokenUsage: Gen[TokenUsage] =
    for
      input <- Gen.choose(0L, 100_000L)
      output <- Gen.choose(0L, 50_000L)
      cacheCreate <- Gen.choose(0L, 10_000L)
      cacheRead <- Gen.choose(0L, 10_000L)
    yield TokenUsage(input, output, cacheCreate, cacheRead)

  given Arbitrary[TokenUsage] = Arbitrary(genTokenUsage)

  val genSmallTokenUsage: Gen[TokenUsage] =
    for
      input <- Gen.choose(100L, 1000L)
      output <- Gen.choose(50L, 500L)
    yield TokenUsage(input, output, 0, 0)

  // --------------------------------------------------------------------------
  // Cell Mismatch Generators
  // --------------------------------------------------------------------------

  val genCellRef: Gen[String] =
    for
      col <- Gen.choose('A', 'Z')
      row <- Gen.choose(1, 100)
    yield s"$col$row"

  val genCellValue: Gen[String] =
    Gen.oneOf(
      Gen.choose(0, 10000).map(_.toString),
      Gen.alphaNumStr.map(s => s"\"${s.take(20)}\""),
      Gen.const("<empty>"),
      Gen.oneOf("#REF!", "#VALUE!", "#N/A").map(e => s"$e")
    )

  val genCellMismatch: Gen[CellMismatch] =
    for
      ref <- genCellRef
      expected <- genCellValue
      actual <- genCellValue
    yield CellMismatch(ref, expected, actual)

  given Arbitrary[CellMismatch] = Arbitrary(genCellMismatch)

  // --------------------------------------------------------------------------
  // Case Details Generators
  // --------------------------------------------------------------------------

  val genCaseDetails: Gen[CaseDetails] =
    Gen.oneOf(
      Gen.const(CaseDetails.NoDetails),
      Gen.listOfN(Gen.choose(0, 5).sample.getOrElse(0), genCellMismatch).map(CaseDetails.FileComparison.apply),
      for
        response <- Gen.alphaNumStr.map(_.take(100))
        expected <- Gen.option(Gen.alphaNumStr.map(_.take(50)))
      yield CaseDetails.TokenComparison(response, expected)
    )

  given Arbitrary[CaseDetails] = Arbitrary(genCaseDetails)

  // --------------------------------------------------------------------------
  // Case Result Generators
  // --------------------------------------------------------------------------

  val genCaseResult: Gen[CaseResult] =
    for
      caseNum <- Gen.choose(1, 10)
      passed <- Gen.oneOf(true, false)
      usage <- genSmallTokenUsage
      latencyMs <- Gen.choose(100L, 30_000L)
      details <- genCaseDetails
      error <- Gen.option(Gen.alphaNumStr.map(s => s"Error: ${s.take(50)}"))
    yield CaseResult(
      caseNum = caseNum,
      passed = passed,
      usage = usage,
      latencyMs = latencyMs,
      details = details,
      tracePath = None,
      error = if passed then None else error
    )

  given Arbitrary[CaseResult] = Arbitrary(genCaseResult)

  // --------------------------------------------------------------------------
  // Task ID Generators
  // --------------------------------------------------------------------------

  val genTaskId: Gen[TaskId] =
    Gen.choose(1000, 9999).map(n => TaskId(n.toString))

  given Arbitrary[TaskId] = Arbitrary(genTaskId)

  // --------------------------------------------------------------------------
  // Score Generators
  // --------------------------------------------------------------------------

  val genBinaryScore: Gen[Score.BinaryScore] =
    Gen.oneOf(Score.BinaryScore.Pass, Score.BinaryScore.Fail)

  val genFractionalScore: Gen[Score.FractionalScore] =
    for
      passing <- Gen.choose(0, 10)
      total <- Gen.choose(passing, 10)
    yield Score.FractionalScore(passing, total.max(1))

  val genLetterGrade: Gen[Score.LetterGrade] =
    Gen.oneOf(
      Score.LetterGrade.A,
      Score.LetterGrade.B,
      Score.LetterGrade.C,
      Score.LetterGrade.D,
      Score.LetterGrade.F
    )

  given Arbitrary[Score.BinaryScore] = Arbitrary(genBinaryScore)
  given Arbitrary[Score.FractionalScore] = Arbitrary(genFractionalScore)
  given Arbitrary[Score.LetterGrade] = Arbitrary(genLetterGrade)

  // --------------------------------------------------------------------------
  // Execution Result Generators
  // --------------------------------------------------------------------------

  val genExecutionResult: Gen[ExecutionResult] =
    for
      taskId <- genTaskId
      skill <- Gen.oneOf("xl", "xlsx")
      numCases <- Gen.choose(1, 5)
      caseResults <- Gen.listOfN(numCases, genCaseResult).map(_.toVector)
      usage <- genTokenUsage
      latencyMs <- Gen.choose(1000L, 60_000L)
      error <- Gen.option(Gen.alphaNumStr.map(s => s"Error: ${s.take(30)}"))
    yield
      val passed = caseResults.count(_.passed)
      ExecutionResult(
        taskId = taskId,
        skill = skill,
        caseResults = caseResults.zipWithIndex.map { case (cr, i) => cr.copy(caseNum = i + 1) },
        aggregateScore = Score.FractionalScore(passed, numCases),
        usage = usage,
        latencyMs = latencyMs,
        error = error
      )

  given Arbitrary[ExecutionResult] = Arbitrary(genExecutionResult)
