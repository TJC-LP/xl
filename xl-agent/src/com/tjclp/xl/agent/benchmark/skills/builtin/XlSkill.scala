package com.tjclp.xl.agent.benchmark.skills.builtin

import cats.effect.{Clock, IO}
import cats.effect.syntax.concurrent.*
import cats.syntax.all.*
import com.tjclp.xl.agent.{Agent, AgentConfig, AgentTask, UploadedFile}
import com.tjclp.xl.agent.anthropic.{AnthropicClientIO, SkillsApi}
import com.tjclp.xl.agent.approach.XlApproachStrategy
import com.tjclp.xl.agent.benchmark.{Evaluator, RangeResult}
import com.tjclp.xl.agent.benchmark.common.FileManager
import com.tjclp.xl.agent.benchmark.execution.*
import com.tjclp.xl.agent.benchmark.grading.{CellMismatch, Score}
import com.tjclp.xl.agent.benchmark.skills.{Skill, SkillContext, SkillRegistry}
import com.tjclp.xl.agent.benchmark.task.*
import com.tjclp.xl.agent.benchmark.tracing.ConversationTracer

import java.nio.file.{Files, Path}

/** Skill using the custom xl-cli tool with the Skills API */
object XlSkill extends Skill:

  override val name: String = "xl"
  override val displayName: String = "xl-cli"
  override val description: String = "Custom xl-cli tool for Excel operations"

  override def setup(client: AnthropicClientIO, config: AgentConfig): IO[SkillContext] =
    for
      // Resolve paths
      binaryPath <- FileManager.resolveBinaryPath(config.xlBinaryPath)
      skillPath <- FileManager.resolveSkillPath(config.xlSkillPath)

      _ <- IO.println(s"  [$name] Checking binary: ${binaryPath.getFileName}")
      binaryFile <- client.uploadFileIfNeeded(binaryPath)

      _ <- IO.println(s"  [$name] Registering skill via Skills API...")
      apiKey <- AnthropicClientIO.loadApiKey
      skillId <- SkillsApi.getOrCreateXlSkill(apiKey, skillPath)
      _ <- IO.println(s"  [$name] Skill ID: $skillId")
    yield SkillContext(
      fileIds = List(binaryFile.id),
      skillId = Some(skillId),
      metadata = Map(
        "binaryFilename" -> binaryFile.filename,
        "binaryFileId" -> binaryFile.id
      )
    )

  override def teardown(client: AnthropicClientIO, ctx: SkillContext): IO[Unit] =
    ctx.fileIds.traverse_(id => client.deleteFile(id).attempt.void)

  override def createAgent(
    client: AnthropicClientIO,
    ctx: SkillContext,
    config: AgentConfig
  ): Agent =
    val binaryFile = UploadedFile(
      id = ctx.metadata.getOrElse("binaryFileId", ""),
      filename = ctx.metadata.getOrElse("binaryFilename", "xl")
    )
    val strategy = XlApproachStrategy(binaryFile, ctx.skillId.getOrElse(""))
    Agent.create(client, config, strategy)

  override def execute(
    task: BenchmarkTask,
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[ExecutionResult] =
    task.inputSource match
      case InputSource.TestCases(cases) =>
        runTestCases(task, cases.toList, ctx, client, agentConfig, engineConfig)

      case InputSource.SingleFile(path) =>
        runSingleFile(task, path, ctx, client, agentConfig, engineConfig)

      case InputSource.NoInput =>
        runNoInput(task, ctx, client, agentConfig, engineConfig)

      case InputSource.DataDirectory(dir, pattern) =>
        // Resolve files dynamically
        IO.blocking {
          import scala.jdk.CollectionConverters.*
          val glob = dir.getFileSystem.getPathMatcher(s"glob:$pattern")
          Files
            .walk(dir)
            .iterator()
            .asScala
            .filter(p => Files.isRegularFile(p) && glob.matches(p.getFileName))
            .map(p => TestCaseFile(1, p, p)) // Placeholder - single case
            .toVector
        }.flatMap { cases =>
          if cases.isEmpty then
            IO.pure(ExecutionResult.skipped(task.id, name, "No files matched pattern"))
          else runTestCases(task, cases.toList, ctx, client, agentConfig, engineConfig)
        }

  /** Run task with multiple test cases (SpreadsheetBench style) */
  private def runTestCases(
    task: BenchmarkTask,
    cases: List[TestCaseFile],
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[ExecutionResult] =
    for
      startTime <- Clock[IO].monotonic.map(_.toMillis)

      // Run test cases in parallel (each with its own agent instance)
      caseResults <- cases.parTraverseN(engineConfig.parallelism) { testCase =>
        val agent = createAgent(client, ctx, agentConfig)
        runSingleTestCase(
          agent = agent,
          task = task,
          testCase = testCase,
          engineConfig = engineConfig,
          agentConfig = agentConfig
        )
      }

      endTime <- Clock[IO].monotonic.map(_.toMillis)
      latencyMs = endTime - startTime

      // Aggregate usage
      totalUsage = caseResults.foldLeft(TokenUsage.zero)(_ + _.usage)

      // Build result
      result = ExecutionResult.fromCases(
        taskId = task.id,
        skill = name,
        caseResults = caseResults.toVector,
        usage = totalUsage,
        latencyMs = latencyMs
      )

      // Collect trace paths if any
      traces = caseResults.flatMap(_.tracePath).toVector
    yield result.withTraces(traces)

  /** Run a single test case and evaluate against expected answer */
  private def runSingleTestCase(
    agent: Agent,
    task: BenchmarkTask,
    testCase: TestCaseFile,
    engineConfig: EngineConfig,
    agentConfig: AgentConfig
  ): IO[CaseResult] =
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

      // Create agent task
      agentTask = AgentTask(
        instruction = task.instruction,
        inputFile = testCase.inputPath,
        outputFile = outputPath,
        answerPosition = task.evaluation.answerPosition
      )

      // Run agent with streaming to capture conversation
      result <- agent.runStreaming(agentTask, tracer.onEvent)

      // Evaluate output against answer
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

      // Convert mismatches to grading format
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
          details = CaseDetails.NoDetails
        )
      )
    }

  /** Run task with single input file (TokenBenchmark style) */
  private def runSingleFile(
    task: BenchmarkTask,
    inputPath: Path,
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[ExecutionResult] =
    for
      startTime <- Clock[IO].monotonic.map(_.toMillis)

      agent = createAgent(client, ctx, agentConfig)

      outputDir = engineConfig.outputDir.resolve("outputs").resolve(task.taskIdValue).resolve(name)
      outputPath = outputDir.resolve("output.xlsx")
      _ <- IO.blocking(Files.createDirectories(outputDir))

      // Create tracer for conversation logs
      tracer <- ConversationTracer.create(
        outputDir = engineConfig.outputDir,
        taskId = task.taskIdValue,
        skillName = name,
        caseNum = 1,
        streaming = engineConfig.stream
      )

      agentTask = AgentTask(
        instruction = task.instruction,
        inputFile = inputPath,
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

      // For single file tasks (analysis), we defer correctness to LLM grading.
      // Set passed = true here; the grader will update aggregateScore based on response quality.
      caseResult = CaseResult(
        caseNum = 1,
        passed = agentSucceeded, // Agent ran without error; grading determines correctness
        usage = usage,
        latencyMs = latencyMs,
        details = CaseDetails.token(
          result.responseText.getOrElse(""),
          task.evaluation.expectedAnswer
        ),
        tracePath = Some(tracePath)
      )
    yield ExecutionResult
      .fromSingleCase(
        taskId = task.id,
        skill = name,
        caseResult = caseResult,
        usage = usage,
        latencyMs = latencyMs
      )
      .withTraces(Vector(tracePath))

  /** Run task without input file */
  private def runNoInput(
    task: BenchmarkTask,
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[ExecutionResult] =
    // For no-input tasks, we'd need a different execution path
    // For now, mark as skipped
    IO.pure(ExecutionResult.skipped(task.id, name, "NoInput tasks not yet implemented"))
