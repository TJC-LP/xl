package com.tjclp.xl.agent.benchmark.spreadsheetbench

import cats.effect.{ExitCode, IO, IOApp, Resource}
import cats.syntax.all.*
import cats.effect.std.Semaphore
import com.tjclp.xl.agent.{Agent, AgentConfig, AgentEvent, AgentTask, TokenUsage, UploadedFile}
import com.tjclp.xl.agent.anthropic.AnthropicClientIO
import com.tjclp.xl.agent.benchmark.*
import com.tjclp.xl.agent.benchmark.common.FileManager
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

        // Resolve and upload binary and skill files
        binaryPath <- resolveBinaryPath(config)
        skillPath <- resolveSkillPath(config)
        _ <- IO.println(s"Uploading xl binary and skill...")
        _ <- IO.println(
          s"  Binary: $binaryPath${if config.xlBinary.isDefined then " (dev override)" else ""}"
        )
        _ <- IO.println(
          s"  Skill: $skillPath${if config.xlSkill.isDefined then " (dev override)" else ""}"
        )

        binary <- client.uploadFile(binaryPath)
        _ <- IO.println(s"  Binary uploaded: ${binary.id}")
        skill <- client.uploadFile(skillPath)
        _ <- IO.println(s"  Skill uploaded: ${skill.id}")

        // Load tasks
        tasks <- TaskLoader.load(config)
        _ <- IO.println(s"Running ${tasks.length} task(s) with parallelism ${config.parallelism}")
        _ <- IO.println("-" * 60)

        // Run tasks
        agentConfig = AgentConfig(
          verbose = config.verbose,
          xlBinaryPath = config.xlBinary,
          xlSkillPath = config.xlSkill
        )
        agent = Agent.create(client, agentConfig, binary, skill)

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

        // Cleanup
        _ <- client.deleteFile(binary.id).attempt
        _ <- client.deleteFile(skill.id).attempt
      yield ()
    }

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

        // Aggregate usage
        totalUsage = testResults.foldLeft(TokenUsage.zero)((acc, r) => acc + r._2)
        responseTexts = testResults.map(_._3)

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

        _ <- saveTaskLog(logsDir, task, result, responseTexts)
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
  ): IO[(TestCaseResult, TokenUsage, Option[String])] =
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
      result.responseText
    ))
      .handleErrorWith { e =>
        IO.println(s"  [${task.id}:${testCase.caseNum}] ERROR: ${e.getMessage}") *>
          IO.pure(
            (
              TestCaseResult(testCase.caseNum, false, Nil, Some(e.getMessage)),
              TokenUsage.zero,
              None
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
        aggregate = AggregateStats.fromResults(results),
        results = results
      )
    }

  private def saveReport(config: BenchConfig, report: BenchmarkReport): IO[Unit] =
    IO.blocking {
      Files.createDirectories(config.outputDir)
      val reportPath = config.outputDir.resolve(s"report_${report.timestamp}.json")
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
    responses: List[Option[String]]
  ): IO[Unit] =
    IO.blocking {
      val logPath = logsDir.resolve(s"${task.id}.md")
      val content = new StringBuilder()

      content.append(s"# Task ${task.id}\n\n")
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
  --parallelism <n>     Number of parallel tasks (default: 4, use 1 for debugging)
  --output <dir>        Output directory for results (default: results)
  --verbose             Show full model responses

Development:
  --xl-binary <path>    Use local xl binary instead of released version
  --xl-skill <path>     Use local xl skill zip instead of released version

Examples:
  # Run a single task
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.spreadsheetbench.Runner -- --task 59196

  # Run first 5 tasks sequentially with verbose output
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.spreadsheetbench.Runner -- --limit 5 --parallelism 1 --verbose

Output:
  results/
    report_<timestamp>.json     # Full benchmark results
    logs/<task_id>.md           # Per-task detailed logs with model responses
    outputs/<task_id>/          # Output files for evaluation
"""
