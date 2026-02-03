package com.tjclp.xl.agent.benchmark.skills.builtin

import cats.effect.{Clock, IO}
import cats.syntax.all.*
import com.tjclp.xl.agent.{Agent, AgentConfig, AgentTask}
import com.tjclp.xl.agent.anthropic.AnthropicClientIO
import com.tjclp.xl.agent.approach.XlsxApproachStrategy
import com.tjclp.xl.agent.benchmark.{Evaluator, RangeResult}
import com.tjclp.xl.agent.benchmark.execution.*
import com.tjclp.xl.agent.benchmark.grading.{CellMismatch, Score}
import com.tjclp.xl.agent.benchmark.skills.{Skill, SkillContext, SkillRegistry}
import com.tjclp.xl.agent.benchmark.task.*
import com.tjclp.xl.agent.benchmark.tracing.ConversationTracer

import java.nio.file.{Files, Path}

/** Skill using Anthropic's built-in xlsx skill (openpyxl) */
object XlsxSkill extends Skill:

  override val name: String = "xlsx"
  override val displayName: String = "OpenPyXL"
  override val description: String = "Anthropic built-in xlsx skill using openpyxl"

  override def setup(client: AnthropicClientIO, config: AgentConfig): IO[SkillContext] =
    IO.println(s"  [$name] Using built-in xlsx skill (no setup required)") *>
      IO.pure(SkillContext.empty)

  override def teardown(client: AnthropicClientIO, ctx: SkillContext): IO[Unit] =
    IO.unit // Nothing to clean up

  override def createAgent(
    client: AnthropicClientIO,
    ctx: SkillContext,
    config: AgentConfig
  ): Agent =
    Agent.create(client, config, XlsxApproachStrategy())

  override def execute(
    task: BenchmarkTask,
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[ExecutionResult] =
    // Default implementation uses enumerateCases + executeCase for compatibility
    // The engine now calls executeCase directly for flattened scheduling
    for
      startTime <- Clock[IO].monotonic.map(_.toMillis)
      cases <- enumerateCases(task)
      caseResults <-
        if cases.isEmpty then IO.pure(Vector.empty[CaseResult])
        else
          cases.toList.traverse(c => executeCase(c, task, ctx, client, agentConfig, engineConfig))
      endTime <- Clock[IO].monotonic.map(_.toMillis)
      latencyMs = endTime - startTime
      totalUsage = caseResults.foldLeft(TokenUsage.zero)(_ + _.usage)
      traces = caseResults.flatMap(_.tracePath).toVector
    yield
      if cases.isEmpty then ExecutionResult.skipped(task.id, name, "No test cases found")
      else
        ExecutionResult
          .fromCases(task.id, name, caseResults.toVector, totalUsage, latencyMs)
          .withTraces(traces)

  override def executeCase(
    testCase: TestCaseFile,
    task: BenchmarkTask,
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[CaseResult] =
    // Determine execution type based on task's input source
    task.inputSource match
      case InputSource.SingleFile(_) =>
        // Token comparison task (analysis) - no answer file comparison
        runTokenComparisonCase(testCase, task, ctx, client, agentConfig, engineConfig)
      case _ =>
        // File comparison task (SpreadsheetBench style)
        runFileComparisonCase(testCase, task, ctx, client, agentConfig, engineConfig)

  /** Run a file comparison test case (SpreadsheetBench style) */
  private def runFileComparisonCase(
    testCase: TestCaseFile,
    task: BenchmarkTask,
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[CaseResult] =
    val agent = createAgent(client, ctx, agentConfig)
    val outputDir =
      engineConfig.outputDir.resolve("outputs").resolve(task.taskIdValue).resolve(name)
    val outputPath = outputDir.resolve(s"${testCase.caseNum}_output.xlsx")

    (for
      _ <- IO.blocking(Files.createDirectories(outputDir))
      startTime <- Clock[IO].monotonic.map(_.toMillis)

      // Create tracer for conversation logs
      tracer <- ConversationTracer.create(
        outputDir = engineConfig.outputDir,
        taskId = task.taskIdValue,
        skillName = name,
        caseNum = testCase.caseNum,
        streaming = engineConfig.stream
      )

      agentTask = AgentTask(
        instruction = task.instruction,
        inputFile = testCase.inputPath,
        outputFile = outputPath,
        answerPosition = task.evaluation.answerPosition
      )

      // Run agent with streaming to capture conversation
      result <- agent.runStreaming(agentTask, tracer.onEvent)

      rangeResults <- result.outputPath match
        case Some(path) =>
          val answerPos = task.evaluation.answerPosition.getOrElse("A1")
          Evaluator.compare(path, testCase.answerPath, answerPos, engineConfig.xlCliPath)
        case None =>
          IO.pure(List(RangeResult("N/A", false, Nil)))

      passed = rangeResults.forall(_.passed)

      endTime <- Clock[IO].monotonic.map(_.toMillis)
      latencyMs = endTime - startTime

      // Complete and save tracer
      _ <- tracer.complete(result.usage, passed, result.error)
      tracePath <- tracer.save()

      mismatches = rangeResults.flatMap(_.mismatches).map { m =>
        CellMismatch(
          ref = m.ref,
          expected = m.expected.toString,
          actual = m.actual.toString
        )
      }

      details =
        if passed then CaseDetails.filePass
        else CaseDetails.fileFail(mismatches)
    yield CaseResult(
      caseNum = testCase.caseNum,
      passed = passed,
      usage = TokenUsage.fromAgentUsage(result.usage),
      latencyMs = latencyMs,
      details = details,
      tracePath = Some(tracePath)
    )).handleErrorWith { e =>
      IO.pure(
        CaseResult(
          caseNum = testCase.caseNum,
          passed = false,
          usage = TokenUsage.zero,
          latencyMs = 0,
          details = CaseDetails.NoDetails,
          error = Some(s"${e.getClass.getSimpleName}: ${e.getMessage}")
        )
      )
    }

  /** Run a token comparison test case (TokenBenchmark/analysis style) */
  private def runTokenComparisonCase(
    testCase: TestCaseFile,
    task: BenchmarkTask,
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[CaseResult] =
    val agent = createAgent(client, ctx, agentConfig)
    val outputDir =
      engineConfig.outputDir.resolve("outputs").resolve(task.taskIdValue).resolve(name)
    val outputPath = outputDir.resolve(s"${testCase.caseNum}_output.xlsx")

    (for
      _ <- IO.blocking(Files.createDirectories(outputDir))
      startTime <- Clock[IO].monotonic.map(_.toMillis)

      // Create tracer for conversation logs
      tracer <- ConversationTracer.create(
        outputDir = engineConfig.outputDir,
        taskId = task.taskIdValue,
        skillName = name,
        caseNum = testCase.caseNum,
        streaming = engineConfig.stream
      )

      agentTask = AgentTask(
        instruction = task.instruction,
        inputFile = testCase.inputPath,
        outputFile = outputPath,
        answerPosition = task.evaluation.answerPosition
      )

      // Run agent with streaming to capture conversation
      result <- agent.runStreaming(agentTask, tracer.onEvent)

      endTime <- Clock[IO].monotonic.map(_.toMillis)
      latencyMs = endTime - startTime

      usage = TokenUsage.fromAgentUsage(result.usage)

      // For analysis tasks, "success" means agent ran without error.
      // Actual correctness is determined by LLM grading (via CaseDetails.TokenComparison).
      agentSucceeded = result.error.isEmpty

      // Complete and save tracer
      _ <- tracer.complete(result.usage, agentSucceeded, result.error)
      tracePath <- tracer.save()
    yield CaseResult(
      caseNum = testCase.caseNum,
      passed = agentSucceeded, // Agent ran without error; grading determines correctness
      usage = usage,
      latencyMs = latencyMs,
      details = CaseDetails.token(
        result.responseText.getOrElse(""),
        task.evaluation.expectedAnswer
      ),
      tracePath = Some(tracePath)
    )).handleErrorWith { e =>
      IO.pure(
        CaseResult(
          caseNum = testCase.caseNum,
          passed = false,
          usage = TokenUsage.zero,
          latencyMs = 0,
          details = CaseDetails.NoDetails,
          error = Some(s"${e.getClass.getSimpleName}: ${e.getMessage}")
        )
      )
    }
