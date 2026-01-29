package com.tjclp.xl.agent.benchmark.spreadsheetbench

import cats.effect.{ExitCode, IO, IOApp, Resource}
import cats.syntax.all.*
import cats.effect.std.Semaphore
import com.tjclp.xl.agent.{Agent, AgentConfig, AgentEvent, AgentTask, TokenUsage, UploadedFile}
import com.tjclp.xl.agent.anthropic.{AnthropicClientIO, SkillsApi}
import com.tjclp.xl.agent.approach.{ApproachStrategy, XlApproachStrategy, XlsxApproachStrategy}
import com.tjclp.xl.agent.benchmark.*
import com.tjclp.xl.agent.benchmark.common.{
  ApproachResult,
  ComparisonReport,
  ComparisonReporter,
  ComparisonStats,
  FileManager,
  ModelPricing,
  TaskComparison,
  UsageBreakdown
}
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.{Files, Path, StandardOpenOption}
import java.time.{Instant, ZoneId}
import java.time.format.DateTimeFormatter

/** SpreadsheetBench benchmark runner */
object Runner extends IOApp:

  override def run(args: List[String]): IO[ExitCode] =
    parseArgs(args)
      .flatMap { config =>
        runBenchmark(config).as(ExitCode.Success)
      }
      .handleErrorWith { e =>
        IO.println(s"Error: ${e.getMessage}").as(ExitCode.Error)
      }

  def runBenchmark(config: BenchConfig): IO[Unit] =
    AnthropicClientIO.fromEnv.use { client =>
      for
        // Create output directories
        _ <- IO.blocking(Files.createDirectories(config.outputDir))
        logsDir = config.outputDir.resolve("logs")
        _ <- IO.blocking(Files.createDirectories(logsDir))

        // Load tasks
        tasks <- TaskLoader.load(config)

        // Branch on comparison mode
        _ <-
          if config.compare then runComparisonBenchmark(config, client, tasks, logsDir)
          else runSingleApproachBenchmark(config, client, tasks, logsDir)
      yield ()
    }

  /** Run benchmark with a single approach (original behavior) */
  private def runSingleApproachBenchmark(
    config: BenchConfig,
    client: AnthropicClientIO,
    tasks: List[BenchTask],
    logsDir: Path
  ): IO[Unit] =
    for
      // Create approach strategy and track files for cleanup
      _ <- IO.println(s"Approach: ${config.approach}")
      strategyAndCleanup <- createStrategy(config, client)
      (strategy, cleanupFiles) = strategyAndCleanup

      _ <- IO.println(s"Running ${tasks.length} task(s) with parallelism ${config.parallelism}")
      _ <- IO.println("-" * 60)

      // Run tasks
      agentConfig = AgentConfig(
        verbose = config.verbose,
        xlBinaryPath = config.xlBinary,
        xlSkillPath = config.xlSkill
      )
      agent = Agent.create(client, agentConfig, strategy)

      results <-
        if config.parallelism == 1 then
          tasks.traverse(task => runTask(agent, task, config, logsDir))
        else
          Semaphore[IO](config.parallelism.toLong).flatMap { sem =>
            tasks.parTraverse(task => sem.permit.use(_ => runTask(agent, task, config, logsDir)))
          }

      // Build and save report
      report <- buildReport(config, results)
      _ <- saveReport(config, report)
      _ <- printSummary(report)

      // Cleanup uploaded files
      _ <- cleanupFiles.traverse_(id => client.deleteFile(id).attempt)
    yield ()

  /** Run benchmark comparing both xl and xlsx approaches */
  private def runComparisonBenchmark(
    config: BenchConfig,
    client: AnthropicClientIO,
    tasks: List[BenchTask],
    logsDir: Path
  ): IO[Unit] =
    for
      _ <- IO.println("Comparison mode: running both xl and xlsx approaches")
      _ <- IO.println("-" * 60)

      // Setup xl approach (unless xlsx-only)
      xlSetup <-
        if config.xlsxOnly then IO.pure(None)
        else
          for
            binaryPath <- resolveBinaryPath(config)
            skillPath <- resolveSkillPath(config)
            _ <- IO.println(s"Setting up xl approach...")
            binary <- client.uploadFile(binaryPath)
            _ <- IO.println(s"  Binary uploaded: ${binary.id}")
            apiKey <- AnthropicClientIO.loadApiKey
            skillId <- SkillsApi.getOrCreateXlSkill(apiKey, skillPath)
            _ <- IO.println(s"  Skill ID: $skillId")
            strategy = XlApproachStrategy(binary, skillId)
          yield Some((strategy, binary.id))

      // Setup xlsx approach (unless xl-only)
      xlsxStrategy <-
        if config.xlOnly then IO.pure(None)
        else
          IO.println("Setting up xlsx approach (openpyxl)...") *>
            IO.pure(Some(XlsxApproachStrategy()))

      _ <- IO.println(s"Running ${tasks.length} task(s) with parallelism ${config.parallelism}")
      _ <- IO.println("-" * 60)

      agentConfig = AgentConfig(
        verbose = config.verbose,
        xlBinaryPath = config.xlBinary,
        xlSkillPath = config.xlSkill
      )

      // Run tasks with both approaches in parallel
      comparisons <-
        if config.parallelism == 1 then
          tasks.traverse(task =>
            runTaskComparison(client, agentConfig, xlSetup, xlsxStrategy, task, config, logsDir)
          )
        else
          Semaphore[IO](config.parallelism.toLong).flatMap { sem =>
            tasks.parTraverse(task =>
              sem.permit.use(_ =>
                runTaskComparison(client, agentConfig, xlSetup, xlsxStrategy, task, config, logsDir)
              )
            )
          }

      // Print comparison table
      pricing = ModelPricing.forModel(agentConfig.model)
      _ <- ComparisonReporter.printTable("SpreadsheetBench: xl vs xlsx", comparisons, pricing)

      // Save comparison report
      _ <- saveComparisonReport(config, agentConfig.model, comparisons, pricing)

      // Cleanup
      _ <- xlSetup.traverse_ { case (_, fileId) => client.deleteFile(fileId).attempt }
    yield ()

  /** Save comparison report to JSON file */
  private def saveComparisonReport(
    config: BenchConfig,
    model: String,
    comparisons: List[TaskComparison],
    pricing: ModelPricing
  ): IO[Unit] =
    IO.blocking {
      val timestamp = DateTimeFormatter
        .ofPattern("yyyy-MM-dd'T'HH-mm-ss")
        .format(Instant.now.atZone(ZoneId.systemDefault()))

      val report = ComparisonReport(
        timestamp = timestamp,
        dataset = config.dataset,
        model = model,
        pricing = pricing,
        stats = ComparisonStats.fromResults(comparisons, pricing),
        results = comparisons
      )

      Files.createDirectories(config.outputDir)
      val reportPath = config.outputDir.resolve(s"comparison_${timestamp}.json")
      Files.writeString(reportPath, report.toJsonPretty)
      println(s"\nComparison report saved to: $reportPath")
    }

  /** Run a single task with both approaches */
  private def runTaskComparison(
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    xlSetup: Option[(ApproachStrategy, String)],
    xlsxStrategy: Option[ApproachStrategy],
    task: BenchTask,
    config: BenchConfig,
    logsDir: Path
  ): IO[TaskComparison] =
    if task.isVba && config.taskIds.isEmpty then
      IO.println(s"[${task.id}] SKIPPED (VBA/Macro task)") *>
        IO.pure(TaskComparison(task.id, task.instruction.take(40), None, None))
    else
      for
        _ <- IO.println(s"[${task.id}] Starting comparison...")

        // Run xl approach
        xlResult <- xlSetup.traverse { case (strategy, _) =>
          val agent = Agent.create(client, agentConfig, strategy)
          runTaskForComparison(agent, task, config, logsDir, "xl")
        }

        // Run xlsx approach
        xlsxResult <- xlsxStrategy.traverse { strategy =>
          val agent = Agent.create(client, agentConfig, strategy)
          runTaskForComparison(agent, task, config, logsDir, "xlsx")
        }

        // Log summary
        xlSummary = xlResult
          .map(r => s"${r.usage.inputTokens}in/${r.usage.outputTokens}out")
          .getOrElse("-")
        xlsxSummary = xlsxResult
          .map(r => s"${r.usage.inputTokens}in/${r.usage.outputTokens}out")
          .getOrElse("-")
        xlPass = xlResult.map(r => if r.passed then "\u2713" else "\u2717").getOrElse("-")
        xlsxPass = xlsxResult.map(r => if r.passed then "\u2713" else "\u2717").getOrElse("-")
        _ <- IO.println(s"[${task.id}] xl=$xlSummary $xlPass, xlsx=$xlsxSummary $xlsxPass")
      yield TaskComparison(task.id, task.instruction.take(40), xlResult, xlsxResult)

  /** Run a task with a single approach and convert to ApproachResult */
  private def runTaskForComparison(
    agent: Agent,
    task: BenchTask,
    config: BenchConfig,
    logsDir: Path,
    approach: String
  ): IO[ApproachResult] =
    (for
      startTime <- IO.realTimeInstant.map(_.toEpochMilli)

      // Run all test cases
      testResults <- task.testCases.traverse { testCase =>
        val outputDir = config.outputDir.resolve("outputs").resolve(task.id).resolve(approach)
        val outputPath = outputDir.resolve(s"${testCase.caseNum}_output.xlsx")

        IO.blocking(Files.createDirectories(outputDir)) *>
          runSingleTestCase(agent, task, testCase, outputPath, config)
      }

      endTime <- IO.realTimeInstant.map(_.toEpochMilli)
      latencyMs = endTime - startTime

      // Aggregate results
      totalUsage = testResults.foldLeft(TokenUsage.zero)((acc, r) => acc + r._1)
      allPassed = testResults.forall(_._2)
    yield ApproachResult(
      approach = approach,
      success = true,
      passed = allPassed,
      usage = UsageBreakdown.fromTokenUsage(totalUsage.inputTokens, totalUsage.outputTokens),
      latencyMs = latencyMs
    )).handleErrorWith { e =>
      IO.pure(
        ApproachResult(
          approach = approach,
          success = false,
          passed = false,
          usage = UsageBreakdown(0, 0),
          latencyMs = 0,
          error = Some(e.getMessage)
        )
      )
    }

  /** Run a single test case and return (usage, passed) */
  private def runSingleTestCase(
    agent: Agent,
    task: BenchTask,
    testCase: TestCase,
    outputPath: Path,
    config: BenchConfig
  ): IO[(TokenUsage, Boolean)] =
    val agentTask = AgentTask(
      instruction = task.instruction,
      inputFile = testCase.inputPath,
      outputFile = outputPath,
      answerPosition = Some(task.answerPosition)
    )

    for
      result <- agent.run(agentTask)

      // Evaluate
      rangeResults <- result.outputPath match
        case Some(path) =>
          Evaluator.compare(path, testCase.answerPath, task.answerPosition)
        case None =>
          IO.pure(List(RangeResult(task.answerPosition, false, Nil)))

      passed = rangeResults.forall(_.passed)
    yield (result.usage, passed)

  /** Create approach strategy and return file IDs to cleanup */
  private def createStrategy(
    config: BenchConfig,
    client: AnthropicClientIO
  ): IO[(ApproachStrategy, List[String])] =
    config.approach match
      case Approach.Xl =>
        for
          // Resolve paths
          binaryPath <- resolveBinaryPath(config)
          skillPath <- resolveSkillPath(config)

          _ <- IO.println(s"Uploading xl binary...")
          _ <- IO.println(
            s"  Binary: $binaryPath${if config.xlBinary.isDefined then " (dev override)" else ""}"
          )

          binary <- client.uploadFile(binaryPath)
          _ <- IO.println(s"  Binary uploaded: ${binary.id}")

          // Register skill via Skills API (like TokenRunner)
          apiKey <- AnthropicClientIO.loadApiKey
          _ <- IO.println(s"Registering xl-cli skill via Skills API...")
          _ <- IO.println(
            s"  Skill: $skillPath${if config.xlSkill.isDefined then " (dev override)" else ""}"
          )
          skillId <- SkillsApi.getOrCreateXlSkill(apiKey, skillPath)
          _ <- IO.println(s"  Skill ID: $skillId")

          strategy = XlApproachStrategy(binary, skillId)
        yield (strategy, List(binary.id)) // No skill file to cleanup

      case Approach.Xlsx =>
        IO.println("Using built-in xlsx skill (openpyxl)") *>
          IO.pure((XlsxApproachStrategy(), Nil))

  private def runTask(
    agent: Agent,
    task: BenchTask,
    config: BenchConfig,
    logsDir: Path
  ): IO[TaskResult] =
    if task.isVba && config.taskIds.isEmpty then
      IO.println(s"[${task.id}] SKIPPED (VBA/Macro task - use --task to run explicitly)") *>
        IO.pure(TaskResult.skipped(task, "VBA/Macro task"))
    else
      (for
        _ <- IO.println(s"[${task.id}] Starting: ${task.instruction.take(60)}...")
        startTime <- IO.realTimeInstant.map(_.toEpochMilli)

        // Run all 3 test cases
        testResults <- task.testCases.traverse { testCase =>
          runTestCase(agent, task, testCase, config, logsDir)
        }

        endTime <- IO.realTimeInstant.map(_.toEpochMilli)
        latencyMs = endTime - startTime

        // Aggregate usage and collect transcripts
        totalUsage = testResults.foldLeft(TokenUsage.zero)((acc, r) => acc + r._2)
        responseTexts = testResults.map(_._3)
        transcripts = testResults.map(_._4)

        score = TaskScore.fromResults(testResults.map(_._1))
        statusEmoji =
          if score.hardRestriction == 1 then "\u2713"
          else if score.softRestriction > 0 then "\u25d0"
          else "\u2717"
        _ <- IO.println(
          s"[${task.id}] $statusEmoji soft=${score.softRestriction}, hard=${score.hardRestriction}, ${latencyMs}ms"
        )

        flatResponses = responseTexts.flatten
        result = TaskResult(
          taskId = task.id,
          instruction = task.instruction,
          instructionType = task.instructionType,
          score = score,
          testCaseResults = testResults.map(_._1),
          usage = totalUsage,
          latencyMs = latencyMs,
          responseText =
            if flatResponses.nonEmpty then Some(flatResponses.mkString("\n---\n")) else None
        )

        approach = config.approach.toString.toLowerCase
        _ <- saveTaskLog(logsDir, task, result, responseTexts, transcripts, approach)
      yield result).handleErrorWith { e =>
        IO.println(s"[${task.id}] ERROR: ${e.getMessage}") *>
          IO.pure(
            TaskResult(
              taskId = task.id,
              instruction = task.instruction,
              instructionType = task.instructionType,
              score = TaskScore(0.0, 0),
              testCaseResults = Nil,
              usage = TokenUsage.zero,
              latencyMs = 0,
              error = Some(e.getMessage)
            )
          )
      }

  private def runTestCase(
    agent: Agent,
    task: BenchTask,
    testCase: TestCase,
    config: BenchConfig,
    logsDir: Path
  ): IO[(TestCaseResult, TokenUsage, Option[String], Vector[AgentEvent])] =
    val outputDir = config.outputDir.resolve("outputs").resolve(task.id)
    val outputPath = outputDir.resolve(s"${testCase.caseNum}_output.xlsx")

    (for
      _ <- IO.blocking(Files.createDirectories(outputDir))
      _ <- IO.println(s"  [${task.id}:${testCase.caseNum}] Running...")

      agentTask = AgentTask(
        instruction = task.instruction,
        inputFile = testCase.inputPath,
        outputFile = outputPath,
        answerPosition = Some(task.answerPosition)
      )

      result <- agent.run(agentTask)

      _ <- IO.println(
        s"  [${task.id}:${testCase.caseNum}] Tokens: ${result.usage.inputTokens}in/${result.usage.outputTokens}out"
      )

      // Evaluate the result
      rangeResults <- result.outputPath match
        case Some(path) =>
          Evaluator.compare(path, testCase.answerPath, task.answerPosition)
        case None =>
          IO.pure(List(RangeResult(task.answerPosition, false, Nil)))

      passed = rangeResults.forall(_.passed)
      _ <- IO.println(s"  [${task.id}:${testCase.caseNum}] ${if passed then "PASS" else "FAIL"}")

      // Log mismatches
      _ <- rangeResults.flatMap(_.mismatches).take(3).traverse_ { m =>
        IO.println(s"    Mismatch at ${m.ref}: expected=${m.expected}, got=${m.actual}")
      }
    yield (
      TestCaseResult(testCase.caseNum, passed, rangeResults),
      result.usage,
      result.responseText,
      result.transcript
    ))
      .handleErrorWith { e =>
        IO.println(s"  [${task.id}:${testCase.caseNum}] ERROR: ${e.getMessage}") *>
          IO.pure(
            (
              TestCaseResult(testCase.caseNum, false, Nil, Some(e.getMessage)),
              TokenUsage.zero,
              None,
              Vector.empty
            )
          )
      }

  private def buildReport(config: BenchConfig, results: List[TaskResult]): IO[BenchmarkReport] =
    IO {
      val timestamp = DateTimeFormatter
        .ofPattern("yyyy-MM-dd'T'HH-mm-ss")
        .format(
          Instant.now.atZone(ZoneId.systemDefault())
        )
      BenchmarkReport(
        timestamp = timestamp,
        dataset = config.dataset,
        model = "claude-opus-4-5-20251101",
        approach = config.approach.toString.toLowerCase,
        aggregate = AggregateStats.fromResults(results),
        results = results
      )
    }

  private def saveReport(config: BenchConfig, report: BenchmarkReport): IO[Unit] =
    IO.blocking {
      Files.createDirectories(config.outputDir)
      val reportPath =
        config.outputDir.resolve(s"report_${report.approach}_${report.timestamp}.json")
      Files.writeString(reportPath, report.toJsonPretty)
      println(s"\nReport saved to: $reportPath")
    }

  private def printSummary(report: BenchmarkReport): IO[Unit] =
    IO.println(s"""
${"=" * 60}
SpreadsheetBench Results Summary
${"=" * 60}
Dataset: ${report.dataset}
Model: ${report.model}
Approach: ${report.approach}
Timestamp: ${report.timestamp}
${"-" * 60}
Total Tasks: ${report.aggregate.totalTasks}
Completed: ${report.aggregate.completedTasks}
Skipped: ${report.aggregate.skippedTasks}
${"-" * 60}
Avg Soft Restriction: ${f"${report.aggregate.avgSoftRestriction}%.2f"}
Avg Hard Restriction: ${f"${report.aggregate.avgHardRestriction}%.2f"}
${"-" * 60}
Total Input Tokens: ${report.aggregate.totalInputTokens}
Total Output Tokens: ${report.aggregate.totalOutputTokens}
Avg Latency: ${report.aggregate.avgLatencyMs}ms
${"=" * 60}
""")

  private def saveTaskLog(
    logsDir: Path,
    task: BenchTask,
    result: TaskResult,
    responses: List[Option[String]],
    transcripts: List[Vector[AgentEvent]],
    approach: String
  ): IO[Unit] =
    IO.blocking {
      import io.circe.syntax.*

      // Markdown log (with approach in filename)
      val logPath = logsDir.resolve(s"${task.id}_${approach}.md")
      val content = new StringBuilder()

      content.append(s"# Task ${task.id}\n\n")
      content.append(s"**Approach:** $approach\n\n")
      content.append(s"**Type:** ${task.instructionType}\n\n")
      content.append(
        s"**Score:** soft=${result.score.softRestriction}, hard=${result.score.hardRestriction}\n\n"
      )
      content.append(
        s"**Tokens:** ${result.usage.inputTokens} input, ${result.usage.outputTokens} output\n\n"
      )
      content.append(s"**Latency:** ${result.latencyMs}ms\n\n")

      content.append("## Instruction\n\n")
      content.append(task.instruction)
      content.append("\n\n")

      content.append(s"## Answer Position\n\n`${task.answerPosition}`\n\n")

      content.append("## Test Case Results\n\n")
      result.testCaseResults.zipWithIndex.foreach { case (tc, idx) =>
        content.append(s"### Test Case ${tc.caseNum}\n\n")
        content.append(s"**Passed:** ${tc.passed}\n\n")
        tc.error.foreach(e => content.append(s"**Error:** $e\n\n"))
        if tc.rangeResults.flatMap(_.mismatches).nonEmpty then
          content.append("**Mismatches:**\n\n")
          tc.rangeResults.flatMap(_.mismatches).foreach { m =>
            content.append(s"- `${m.ref}`: expected `${m.expected}`, got `${m.actual}`\n")
          }
          content.append("\n")

        // Tool calls from transcript
        transcripts.lift(idx).foreach { events =>
          val toolCalls = events.collect { case AgentEvent.ToolInvocation(_, _, _, Some(cmd)) =>
            cmd
          }
          if toolCalls.nonEmpty then
            content.append(s"**Tool Calls:** ${toolCalls.length}\n\n")
            toolCalls.zipWithIndex.foreach { case (cmd, i) =>
              val truncated = if cmd.length > 500 then cmd.take(500) + "..." else cmd
              content.append(s"${i + 1}. ```bash\n$truncated\n```\n\n")
            }
        }

        responses.lift(idx).flatten.foreach { resp =>
          content.append("**Model Response:**\n\n```\n")
          content.append(resp)
          content.append("\n```\n\n")
        }
      }

      Files.writeString(
        logPath,
        content.toString,
        StandardOpenOption.CREATE,
        StandardOpenOption.TRUNCATE_EXISTING
      )

      // JSON transcript (with approach in filename)
      val transcriptPath = logsDir.resolve(s"${task.id}_${approach}_transcript.json")
      val timestamp = java.time.format.DateTimeFormatter
        .ofPattern("yyyy-MM-dd'T'HH-mm-ss")
        .format(java.time.Instant.now.atZone(java.time.ZoneId.systemDefault()))

      val transcriptJson = io.circe.Json.obj(
        "taskId" -> task.id.asJson,
        "approach" -> approach.asJson,
        "timestamp" -> timestamp.asJson,
        "testCases" -> transcripts.zipWithIndex.map { case (events, idx) =>
          io.circe.Json.obj(
            "caseNum" -> (idx + 1).asJson,
            "eventCount" -> events.length.asJson,
            "events" -> events.asJson
          )
        }.asJson
      )

      Files.writeString(
        transcriptPath,
        transcriptJson.spaces2,
        StandardOpenOption.CREATE,
        StandardOpenOption.TRUNCATE_EXISTING
      )
    }

  private def resolveBinaryPath(config: BenchConfig): IO[Path] =
    FileManager.resolveBinaryPath(config.xlBinary)

  private def resolveSkillPath(config: BenchConfig): IO[Path] =
    FileManager.resolveSkillPath(config.xlSkill)

  private def parseArgs(args: List[String]): IO[BenchConfig] =
    IO {
      var config = BenchConfig()
      var remaining = args

      while remaining.nonEmpty do
        remaining match
          case "--dataset" :: value :: rest =>
            config = config.copy(dataset = value)
            remaining = rest
          case "--limit" :: value :: rest =>
            config = config.copy(limit = Some(value.toInt))
            remaining = rest
          case "--task" :: value :: rest =>
            config = config.copy(taskIds = Some(List(value)))
            remaining = rest
          case "--tasks" :: value :: rest =>
            config = config.copy(taskIds = Some(value.split(",").map(_.trim).toList))
            remaining = rest
          case "--skip" :: value :: rest =>
            config = config.copy(skipIds = value.split(",").map(_.trim).toSet)
            remaining = rest
          case "--category" :: value :: rest =>
            config = config.copy(category = Some(value))
            remaining = rest
          case "--parallelism" :: value :: rest =>
            config = config.copy(parallelism = value.toInt)
            remaining = rest
          case "--verbose" :: rest =>
            config = config.copy(verbose = true)
            remaining = rest
          case "--output" :: value :: rest =>
            config = config.copy(outputDir = Path.of(value))
            remaining = rest
          case "--xl-binary" :: value :: rest =>
            config = config.copy(xlBinary = Some(Path.of(value)))
            remaining = rest
          case "--xl-skill" :: value :: rest =>
            config = config.copy(xlSkill = Some(Path.of(value)))
            remaining = rest
          case "--approach" :: value :: rest =>
            val approach = value.toLowerCase match
              case "xl" => Approach.Xl
              case "xlsx" => Approach.Xlsx
              case _ =>
                println(s"Unknown approach: $value (valid: xl, xlsx)")
                Approach.Xl
            config = config.copy(approach = approach)
            remaining = rest
          case "--compare" :: rest =>
            config = config.copy(compare = true)
            remaining = rest
          case "--xl-only" :: rest =>
            config = config.copy(xlOnly = true)
            remaining = rest
          case "--xlsx-only" :: rest =>
            config = config.copy(xlsxOnly = true)
            remaining = rest
          case "--help" :: _ =>
            println(HelpText)
            throw new RuntimeException("Help requested")
          case unknown :: rest =>
            println(s"Unknown argument: $unknown")
            remaining = rest
          case Nil =>
            ()

      config
    }

  val HelpText: String = """
SpreadsheetBench - Evaluate LLM spreadsheet capabilities

Usage:
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.spreadsheetbench.Runner -- [options]

Task Selection:
  --task <id>           Run a single task by ID
  --tasks <id1,id2,...> Run specific tasks (comma-separated IDs)
  --skip <id1,id2,...>  Skip specific tasks (comma-separated IDs)
  --category <type>     Filter by category (Cell-Level, Sheet-Level)
  --limit <n>           Limit number of tasks

Configuration:
  --dataset <name>      Dataset to use (default: sample_data_200)
  --approach <name>     Approach to use: xl (custom xl-cli) or xlsx (built-in openpyxl)
  --parallelism <n>     Number of parallel tasks (default: 4, use 1 for debugging)
  --output <dir>        Output directory for results (default: results)
  --verbose             Show full model responses

Comparison Mode:
  --compare             Run BOTH xl and xlsx approaches and compare results
  --xl-only             With --compare: only run xl approach
  --xlsx-only           With --compare: only run xlsx approach

Development:
  --xl-binary <path>    Use local xl binary instead of released version
  --xl-skill <path>     Use local xl skill zip instead of released version

Examples:
  # Run a single task with xl-cli approach (default)
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.spreadsheetbench.Runner -- --task 59196

  # Run a single task with xlsx approach (openpyxl)
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.spreadsheetbench.Runner -- --task 59196 --approach xlsx

  # Compare both approaches on the same task (single command!)
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.spreadsheetbench.Runner -- --task 59196 --compare

  # Compare on multiple tasks in parallel
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.spreadsheetbench.Runner -- --tasks 59196,59197 --compare --parallelism 4

  # Run first 5 tasks sequentially with verbose output
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.spreadsheetbench.Runner -- --limit 5 --parallelism 1 --verbose

Output:
  results/
    report_<timestamp>.json     # Full benchmark results
    logs/<task_id>.md           # Per-task detailed logs with model responses
    outputs/<task_id>/          # Output files for evaluation
"""
