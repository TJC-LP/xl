package com.tjclp.xl.agent.benchmark.grading

import munit.CatsEffectSuite

import java.nio.file.Path

class FileGraderSpec extends CatsEffectSuite:

  // --------------------------------------------------------------------------
  // SpreadsheetBenchGrader canHandle Tests
  // --------------------------------------------------------------------------

  test("SpreadsheetBenchGrader.canHandle returns true when all paths present") {
    val grader = SpreadsheetBenchGrader("xl")
    val ctx = GraderContext.forFile(
      taskId = "test-task",
      skill = "xl",
      caseNum = 1,
      outputPath = Path.of("/tmp/output.xlsx"),
      answerPath = Path.of("/tmp/answer.xlsx"),
      answerPosition = "A1"
    )
    assert(grader.canHandle(ctx))
  }

  test("SpreadsheetBenchGrader.canHandle returns false for LLM context") {
    val grader = SpreadsheetBenchGrader("xl")
    val ctx = GraderContext.forLLM(
      taskId = "test-task",
      skill = "xl",
      responseText = "The answer is 42",
      expectedAnswer = "42"
    )
    // LLM context doesn't have outputPath, answerPath, answerPosition
    assert(!grader.canHandle(ctx))
  }

  // --------------------------------------------------------------------------
  // SpreadsheetBenchGrader.grade Tests (Missing File)
  // --------------------------------------------------------------------------

  test("SpreadsheetBenchGrader.grade fails gracefully when output file missing") {
    val grader = SpreadsheetBenchGrader("xl")
    val ctx = GraderContext.forFile(
      taskId = "test-task",
      skill = "xl",
      caseNum = 1,
      outputPath = Path.of("/nonexistent/output.xlsx"),
      answerPath = Path.of("/tmp/answer.xlsx"),
      answerPosition = "A1"
    )

    grader.grade(ctx).map { result =>
      assertEquals(result.score, Score.BinaryScore.Fail)
      assert(!result.passed)
      assert(result.details.reasoning.exists(_.contains("not found")))
    }
  }

  test("SpreadsheetBenchGrader.grade fails gracefully when used with LLM context") {
    val grader = SpreadsheetBenchGrader("xl")
    val ctx = GraderContext.forLLM(
      taskId = "test-task",
      skill = "xl",
      responseText = "The answer is 42",
      expectedAnswer = "42"
    )

    grader.grade(ctx).map { result =>
      assertEquals(result.score, Score.BinaryScore.Fail)
      assert(!result.passed)
      assert(result.details.reasoning.exists(_.contains("Missing")))
    }
  }

  // --------------------------------------------------------------------------
  // MultiCaseFileGrader Tests
  // --------------------------------------------------------------------------

  test("MultiCaseFileGrader.name is correct") {
    val grader = MultiCaseFileGrader("xl")
    assertEquals(grader.name, "multi-case-file")
  }

  test("MultiCaseFileGrader.canHandle delegates to SpreadsheetBenchGrader") {
    val grader = MultiCaseFileGrader("xl")
    val ctx = GraderContext.forFile(
      taskId = "test-task",
      skill = "xl",
      caseNum = 1,
      outputPath = Path.of("/tmp/output.xlsx"),
      answerPath = Path.of("/tmp/answer.xlsx"),
      answerPosition = "A1"
    )
    assert(grader.canHandle(ctx))
  }

  // --------------------------------------------------------------------------
  // GraderContext Factory Tests
  // --------------------------------------------------------------------------

  test("GraderContext.forFile creates context with file fields") {
    val ctx = GraderContext.forFile(
      taskId = "task-1",
      skill = "xl",
      caseNum = 3,
      outputPath = Path.of("/out.xlsx"),
      answerPath = Path.of("/ans.xlsx"),
      answerPosition = "B2:D10"
    )

    assertEquals(ctx.taskId, "task-1")
    assertEquals(ctx.skill, "xl")
    assertEquals(ctx.caseNum, Some(3))
    assertEquals(ctx.outputPath, Some(Path.of("/out.xlsx")))
    assertEquals(ctx.answerPath, Some(Path.of("/ans.xlsx")))
    assertEquals(ctx.answerPosition, Some("B2:D10"))
    assertEquals(ctx.responseText, None)
    assertEquals(ctx.expectedAnswer, None)
  }

  test("GraderContext.forLLM creates context with LLM fields") {
    val ctx = GraderContext.forLLM(
      taskId = "task-2",
      skill = "xlsx",
      responseText = "The answer is 42",
      expectedAnswer = "42",
      taskInstruction = Some("Calculate the sum")
    )

    assertEquals(ctx.taskId, "task-2")
    assertEquals(ctx.skill, "xlsx")
    assertEquals(ctx.responseText, Some("The answer is 42"))
    assertEquals(ctx.expectedAnswer, Some("42"))
    assertEquals(ctx.taskInstruction, Some("Calculate the sum"))
    assertEquals(ctx.outputPath, None)
    assertEquals(ctx.answerPath, None)
    assertEquals(ctx.caseNum, None)
  }

  test("GraderContext.forComposite creates context with all fields") {
    val ctx = GraderContext.forComposite(
      taskId = "task-3",
      skill = "xl",
      caseNum = 2,
      outputPath = Path.of("/out.xlsx"),
      answerPath = Path.of("/ans.xlsx"),
      answerPosition = "A1",
      responseText = "Result: 100",
      expectedAnswer = "100"
    )

    assertEquals(ctx.taskId, "task-3")
    assertEquals(ctx.skill, "xl")
    assertEquals(ctx.caseNum, Some(2))
    assertEquals(ctx.outputPath, Some(Path.of("/out.xlsx")))
    assertEquals(ctx.answerPath, Some(Path.of("/ans.xlsx")))
    assertEquals(ctx.responseText, Some("Result: 100"))
    assertEquals(ctx.expectedAnswer, Some("100"))
  }

  // --------------------------------------------------------------------------
  // CellMismatch Tests
  // --------------------------------------------------------------------------

  test("CellMismatch holds reference and values") {
    val mismatch = CellMismatch(
      ref = "A1",
      expected = "100",
      actual = "99"
    )

    assertEquals(mismatch.ref, "A1")
    assertEquals(mismatch.expected, "100")
    assertEquals(mismatch.actual, "99")
  }

  // --------------------------------------------------------------------------
  // GradeDetails Tests
  // --------------------------------------------------------------------------

  test("GradeDetails.passed creates passing details") {
    val details = GradeDetails.passed
    assert(details.passed)
    assertEquals(details.mismatches, Nil)
  }

  test("GradeDetails.fail creates failing details with reason") {
    val details = GradeDetails.fail("Something went wrong")
    assert(!details.passed)
    assertEquals(details.reasoning, Some("Something went wrong"))
    assertEquals(details.mismatches, Nil)
  }

  test("GradeDetails.withMismatches adds mismatches") {
    val mismatches = List(
      CellMismatch("A1", "10", "20"),
      CellMismatch("B2", "hello", "world")
    )
    val details = GradeDetails.fail("Cells don't match").withMismatches(mismatches)

    assert(!details.passed)
    assertEquals(details.mismatches, mismatches)
  }

  // --------------------------------------------------------------------------
  // GraderType Tests
  // --------------------------------------------------------------------------

  test("GraderType.fromString parses file") {
    assertEquals(GraderType.fromString("file"), Right(GraderType.File))
  }

  test("GraderType.fromString parses llm") {
    assertEquals(GraderType.fromString("llm"), Right(GraderType.LLM))
  }

  test("GraderType.fromString parses both") {
    assertEquals(GraderType.fromString("file,llm"), Right(GraderType.FileAndLLM))
    assertEquals(GraderType.fromString("both"), Right(GraderType.FileAndLLM))
  }

  test("GraderType.fromString handles custom") {
    assertEquals(GraderType.fromString("custom-grader"), Right(GraderType.Custom("custom-grader")))
  }
