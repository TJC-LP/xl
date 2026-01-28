package spreadsheetbench

import zio.*
import zio.json.*
import com.anthropic.client.AnthropicClient
import com.anthropic.client.okhttp.AnthropicOkHttpClient
import com.anthropic.core.JsonValue
import com.anthropic.models.messages.*
import com.anthropic.models.beta.AnthropicBeta
import com.anthropic.models.beta.files.{FileMetadata, FileUploadParams, FileDeleteParams}
import io.github.cdimascio.dotenv.Dotenv

import java.nio.file.{Files, Path, Paths, StandardOpenOption}
import java.time.{Instant, ZoneId}
import java.time.format.DateTimeFormatter
import scala.jdk.CollectionConverters.*
import scala.util.Try

// Exception for clean exit on --help
case class HelpRequestedException() extends Exception("Help requested")

object SpreadsheetBench extends ZIOAppDefault:

  // ============================================================================
  // Configuration
  // ============================================================================

  // Paths relative to project root or script directory
  lazy val BinaryPath: Path = findFile("xl-0.8.0-linux-amd64", List(
    "../benchmark",
    "examples/anthropic-sdk/benchmark",
    "."
  ))
  lazy val SkillPath: Path = findFile("xl-skill-0.8.0.zip", List(
    "../benchmark",
    "examples/anthropic-sdk/benchmark",
    "."
  ))

  private def findFile(name: String, dirs: List[String]): Path =
    dirs.view
      .map(d => Paths.get(d, name))
      .find(p => Files.exists(p))
      .getOrElse(Paths.get(dirs.head, name)) // fallback to first

  val Model = "claude-sonnet-4-5-20250929"

  // ============================================================================
  // Main Entry Point
  // ============================================================================

  override val run: ZIO[ZIOAppArgs & Scope, Any, Any] =
    (for
      args   <- ZIOAppArgs.getArgs
      config <- parseArgs(args.toList)
      _      <- Console.printLine(s"SpreadsheetBench starting...")
      _      <- Console.printLine(s"  Dataset: ${config.dataset}")
      _      <- Console.printLine(s"  Output: ${config.outputDir}")
      _      <- runBenchmark(config)
    yield ()).catchSome {
      case _: HelpRequestedException => ZIO.succeed(())
    }

  // ============================================================================
  // Benchmark Runner
  // ============================================================================

  def runBenchmark(config: BenchConfig): Task[Unit] =
    for
      apiKey  <- loadApiKey
      client  <- ZIO.attempt(buildClient(apiKey))

      // Create output directory
      _       <- ZIO.attemptBlocking(Files.createDirectories(config.outputDir))
      logsDir  = config.outputDir.resolve("logs")
      _       <- ZIO.attemptBlocking(Files.createDirectories(logsDir))

      _       <- Console.printLine("Uploading xl binary and skill...")
      binary  <- uploadFile(client, BinaryPath)
      _       <- Console.printLine(s"  Binary uploaded: ${binary.id()}")

      tasks   <- TaskLoader.load(config)
      _       <- Console.printLine(s"Running ${tasks.length} task(s) with parallelism ${config.parallelism}")
      _       <- Console.printLine("-" * 60)

      results <- (if config.parallelism == 1 then
                    // Sequential for easier debugging
                    ZIO.foreach(tasks)(task => runTask(client, binary.id(), task, config, logsDir))
                  else
                    ZIO.foreachPar(tasks)(task => runTask(client, binary.id(), task, config, logsDir))
                      .withParallelism(config.parallelism))

      report  <- buildReport(config, results)
      _       <- saveReport(config, report)
      _       <- printSummary(report)

      // Cleanup uploaded files
      _       <- ZIO.attempt(client.beta().files().delete(FileDeleteParams.builder().fileId(binary.id()).build()))
                   .catchAll(_ => ZIO.unit)
    yield ()

  // ============================================================================
  // Task Execution
  // ============================================================================

  def runTask(
      client: AnthropicClient,
      binaryId: String,
      task: BenchTask,
      config: BenchConfig,
      logsDir: Path
  ): Task[TaskResult] =
    if task.isVba then
      Console.printLine(s"[${task.id}] SKIPPED (VBA/Macro task)") *>
        ZIO.succeed(TaskResult.skipped(task, "VBA/Macro task"))
    else
      for
        _         <- Console.printLine(s"[${task.id}] Starting: ${task.instruction.take(60)}...")
        startTime <- Clock.currentTime(java.util.concurrent.TimeUnit.MILLISECONDS)

        // Run all 3 test cases
        testResults <- ZIO.foreach(task.testCases) { testCase =>
          runTestCase(client, binaryId, task, testCase, config, logsDir)
        }

        endTime   <- Clock.currentTime(java.util.concurrent.TimeUnit.MILLISECONDS)
        latencyMs  = endTime - startTime

        // Aggregate usage across test cases
        totalUsage = testResults.foldLeft(TokenUsage(0, 0)) { (acc, r) =>
          TokenUsage(acc.inputTokens + r._2.inputTokens, acc.outputTokens + r._2.outputTokens)
        }

        // Collect response texts for logging
        responseTexts = testResults.map(_._3)

        score = TaskScore.fromResults(testResults.map(_._1))
        statusEmoji = if score.hardRestriction == 1 then "✓" else if score.softRestriction > 0 then "◐" else "✗"
        _    <- Console.printLine(s"[${task.id}] $statusEmoji soft=${score.softRestriction}, hard=${score.hardRestriction}, ${latencyMs}ms")

        flatResponses = responseTexts.flatten
        result = TaskResult(
          taskId = task.id,
          instruction = task.instruction,
          instructionType = task.instructionType,
          score = score,
          testCaseResults = testResults.map(_._1),
          usage = totalUsage,
          latencyMs = latencyMs,
          responseText = if flatResponses.nonEmpty then Some(flatResponses.mkString("\n---\n")) else None
        )

        // Save detailed task log
        _ <- saveTaskLog(logsDir, task, result, responseTexts)

      yield result

  def runTestCase(
      client: AnthropicClient,
      binaryId: String,
      task: BenchTask,
      testCase: TestCase,
      config: BenchConfig,
      logsDir: Path
  ): Task[(TestCaseResult, TokenUsage, Option[String])] =
    (for
      // Upload input file for this test case
      inputFile <- uploadFile(client, testCase.inputPath)
      _         <- Console.printLine(s"  [${task.id}:${testCase.caseNum}] Input uploaded, sending request...")

      // Build and send the request
      response  <- sendRequest(client, binaryId, inputFile.id(), task, testCase)

      // Extract response text for logging
      responseText = extractResponseText(response)
      _         <- if config.verbose then Console.printLine(s"  [${task.id}:${testCase.caseNum}] Response:\n$responseText") else ZIO.unit

      // Extract usage
      usage      = TokenUsage(response.usage().inputTokens(), response.usage().outputTokens())
      _         <- Console.printLine(s"  [${task.id}:${testCase.caseNum}] Tokens: ${usage.inputTokens}in/${usage.outputTokens}out")

      // Download the output file from the container
      outputPath <- downloadOutput(client, response, task, testCase, config)

      // Evaluate the result
      rangeResults <- Evaluator.compare(outputPath, testCase.answerPath, task.answerPosition)
      passed = rangeResults.forall(_.passed)
      _         <- Console.printLine(s"  [${task.id}:${testCase.caseNum}] ${if passed then "PASS" else "FAIL"}")

      // Log mismatches if any
      _ <- ZIO.foreach(rangeResults.flatMap(_.mismatches).take(3)) { m =>
        Console.printLine(s"    Mismatch at ${m.ref}: expected=${m.expected}, got=${m.actual}")
      }

      // Cleanup
      _ <- ZIO.attempt(client.beta().files().delete(FileDeleteParams.builder().fileId(inputFile.id()).build()))
             .catchAll(_ => ZIO.unit)

    yield (TestCaseResult(testCase.caseNum, passed, rangeResults), usage, Some(responseText)))
      .catchAll { e =>
        Console.printLine(s"  [${task.id}:${testCase.caseNum}] ERROR: ${e.getMessage}") *>
          ZIO.succeed((TestCaseResult(testCase.caseNum, false, Nil, Some(e.getMessage)), TokenUsage(0, 0), None))
      }

  // ============================================================================
  // API Interactions
  // ============================================================================

  def sendRequest(
      client: AnthropicClient,
      binaryId: String,
      inputFileId: String,
      task: BenchTask,
      testCase: TestCase
  ): Task[Message] =
    ZIO.attemptBlocking {
      val prompt = buildPrompt(task, testCase)

      val params = MessageCreateParams.builder()
        .model(Model)
        .maxTokens(8192L)
        .system(SystemPrompt)
        .addUserMessage("placeholder")
        // Override messages to include container_upload blocks
        .putAdditionalBodyProperty("messages", JsonValue.from(java.util.List.of(
          java.util.Map.of(
            "role", "user",
            "content", java.util.List.of(
              java.util.Map.of("type", "text", "text", prompt),
              java.util.Map.of("type", "container_upload", "file_id", binaryId),
              java.util.Map.of("type", "container_upload", "file_id", inputFileId)
            )
          )
        )))
        // Add code execution tool
        .putAdditionalBodyProperty("tools", JsonValue.from(java.util.List.of(
          java.util.Map.of(
            "type", "code_execution_20250825",
            "name", "code_execution"
          )
        )))
        // Add beta headers
        .putAdditionalHeader("anthropic-beta", "code-execution-2025-08-25,files-api-2025-04-14")
        .build()

      client.messages().create(params)
    }

  def extractResponseText(response: Message): String =
    response.content().asScala
      .filter(_.isText)
      .map(_.asText().text())
      .mkString("\n")

  def downloadOutput(
      client: AnthropicClient,
      response: Message,
      task: BenchTask,
      testCase: TestCase,
      config: BenchConfig
  ): Task[Path] =
    val outputDir = config.outputDir.resolve("outputs").resolve(task.id)
    ZIO.attemptBlocking {
      Files.createDirectories(outputDir)
      val outputPath = outputDir.resolve(s"${testCase.caseNum}_output.xlsx")

      // TODO: Implement actual file download from container
      // For now, copy answer file as placeholder for testing evaluation logic
      if !Files.exists(outputPath) then
        Files.copy(testCase.answerPath, outputPath)

      outputPath
    }

  // ============================================================================
  // Logging
  // ============================================================================

  def saveTaskLog(logsDir: Path, task: BenchTask, result: TaskResult, responses: List[Option[String]]): Task[Unit] =
    ZIO.attemptBlocking {
      val logPath = logsDir.resolve(s"${task.id}.md")
      val content = new StringBuilder()

      content.append(s"# Task ${task.id}\n\n")
      content.append(s"**Type:** ${task.instructionType}\n\n")
      content.append(s"**Score:** soft=${result.score.softRestriction}, hard=${result.score.hardRestriction}\n\n")
      content.append(s"**Tokens:** ${result.usage.inputTokens} input, ${result.usage.outputTokens} output\n\n")
      content.append(s"**Latency:** ${result.latencyMs}ms\n\n")

      content.append("## Instruction\n\n")
      content.append(task.instruction)
      content.append("\n\n")

      content.append(s"## Answer Position\n\n`${task.answerPosition}`\n\n")

      content.append("## Test Case Results\n\n")
      result.testCaseResults.zipWithIndex.foreach { case (tc, idx) =>
        content.append(s"### Test Case ${tc.caseNum}\n\n")
        content.append(s"**Passed:** ${tc.passed}\n\n")
        if tc.error.isDefined then
          content.append(s"**Error:** ${tc.error.get}\n\n")
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

      Files.writeString(logPath, content.toString, StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING)
    }

  // ============================================================================
  // Helpers
  // ============================================================================

  def loadApiKey: Task[String] =
    ZIO.attempt {
      // Try multiple directories for .env file
      val directories = List(".", "..", "examples/anthropic-sdk", "examples/anthropic-sdk/spreadsheetbench")
      val dotenv = directories.view
        .map(dir => Dotenv.configure().directory(dir).ignoreIfMissing().load())
        .find(d => d.get("ANTHROPIC_API_KEY") != null)
        .getOrElse(Dotenv.configure().ignoreIfMissing().load())

      Option(java.lang.System.getenv("ANTHROPIC_API_KEY"))
        .orElse(Option(dotenv.get("ANTHROPIC_API_KEY")))
        .getOrElse(throw new Exception("ANTHROPIC_API_KEY not found. Set in .env or as environment variable."))
    }

  def buildClient(apiKey: String): AnthropicClient =
    AnthropicOkHttpClient.builder()
      .apiKey(apiKey)
      .build()

  def uploadFile(client: AnthropicClient, path: Path): Task[FileMetadata] =
    ZIO.attemptBlocking {
      val params = FileUploadParams.builder()
        .file(path)
        .addBeta(AnthropicBeta.FILES_API_2025_04_14)
        .build()
      client.beta().files().upload(params)
    }

  def buildPrompt(task: BenchTask, testCase: TestCase): String =
    s"""You are a spreadsheet expert with access to the xl CLI for Excel operations.

## Task
${task.instruction}

## Input File
The input file is uploaded at: /mnt/user/${testCase.inputPath.getFileName}

## Output Requirements
Save your solution to: /mnt/user/output.xlsx

The answer will be evaluated at position: ${task.answerPosition}

## Available Commands

First, make the xl binary executable:
```bash
chmod +x /mnt/user/xl-0.8.0-linux-amd64
```

Then use these commands:
```bash
# List sheets
/mnt/user/xl-0.8.0-linux-amd64 -f <file> sheets

# View data (use quotes for sheet names with spaces)
/mnt/user/xl-0.8.0-linux-amd64 -f <file> -s "<sheet>" view <range>

# Write a value
/mnt/user/xl-0.8.0-linux-amd64 -f <file> -o <output> put <ref> <value>

# Write a formula
/mnt/user/xl-0.8.0-linux-amd64 -f <file> -o <output> putf <ref> "=FORMULA"

# Batch operations (JSON on stdin)
echo '[{"op":"put","ref":"A1","value":"Hello"}]' | /mnt/user/xl-0.8.0-linux-amd64 -f <file> -o <output> batch -
```

## Instructions
1. First explore the input file to understand its structure
2. Implement the solution using xl CLI commands
3. Save the result to /mnt/user/output.xlsx
4. Verify your solution by viewing the output cells at ${task.answerPosition}
"""

  val SystemPrompt: String =
    """You are a spreadsheet expert assistant. You have access to the xl CLI tool for manipulating Excel files.

Your goal is to solve spreadsheet tasks efficiently and accurately. Always:
1. Explore the input file first to understand its structure
2. Plan your solution before executing
3. Use the most appropriate xl commands for the task
4. Verify your solution before finishing

The xl CLI supports common operations like viewing data, writing values, writing formulas, and batch operations."""

  // ============================================================================
  // Reporting
  // ============================================================================

  def buildReport(config: BenchConfig, results: List[TaskResult]): Task[BenchmarkReport] =
    ZIO.succeed {
      val timestamp = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH-mm-ss").format(
        Instant.now.atZone(ZoneId.systemDefault())
      )
      BenchmarkReport(
        timestamp = timestamp,
        dataset = config.dataset,
        model = Model,
        aggregate = AggregateStats.fromResults(results),
        results = results
      )
    }

  def saveReport(config: BenchConfig, report: BenchmarkReport): Task[Unit] =
    ZIO.attemptBlocking {
      Files.createDirectories(config.outputDir)
      val reportPath = config.outputDir.resolve(s"report_${report.timestamp}.json")
      Files.writeString(reportPath, report.toJsonPretty)
      println(s"\nReport saved to: $reportPath")
    }

  def printSummary(report: BenchmarkReport): Task[Unit] =
    Console.printLine(s"""
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

  // ============================================================================
  // Argument Parsing
  // ============================================================================

  def parseArgs(args: List[String]): Task[BenchConfig] =
    ZIO.attempt {
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
            // Single task ID
            config = config.copy(taskIds = Some(List(value)))
            remaining = rest
          case "--tasks" :: value :: rest =>
            // Multiple task IDs (comma-separated)
            config = config.copy(taskIds = Some(value.split(",").map(_.trim).toList))
            remaining = rest
          case "--skip" :: value :: rest =>
            // Task IDs to skip (comma-separated)
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
          case "--help" :: _ =>
            println(HelpText)
            throw HelpRequestedException()
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
  scala-cli run . -- [options]

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

Examples:
  # Run a single task
  scala-cli run . -- --task 59196

  # Run multiple specific tasks
  scala-cli run . -- --tasks 59196,52292,58296

  # Run first 5 tasks sequentially with verbose output
  scala-cli run . -- --limit 5 --parallelism 1 --verbose

  # Run all Cell-Level tasks
  scala-cli run . -- --category "Cell-Level Manipulation"

  # Skip problematic tasks
  scala-cli run . -- --limit 10 --skip 99-24,81-41

Output:
  results/
    report_<timestamp>.json     # Full benchmark results
    logs/<task_id>.md           # Per-task detailed logs with model responses
    outputs/<task_id>/          # Output files for evaluation
"""
