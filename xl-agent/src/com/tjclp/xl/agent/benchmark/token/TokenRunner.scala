package com.tjclp.xl.agent.benchmark.token

import cats.effect.{ExitCode, IO, IOApp}
import cats.effect.std.Semaphore
import cats.syntax.all.*
import com.anthropic.client.AnthropicClient as JAnthropicClient
import com.anthropic.core.JsonValue
import com.anthropic.models.beta.AnthropicBeta
import com.anthropic.models.beta.messages.*
import com.tjclp.xl.agent.{TokenUsage, UploadedFile}
import com.tjclp.xl.agent.anthropic.{AnthropicClientIO, SkillsApi}
import com.tjclp.xl.agent.benchmark.common.{BenchmarkUtils, FileManager}
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.{Files, Path, Paths}
import scala.jdk.CollectionConverters.*
import scala.jdk.OptionConverters.*

/**
 * Token Efficiency Benchmark Runner
 *
 * Compares token usage between xl CLI (custom skill) and Anthropic's built-in xlsx skill.
 *
 * Usage: ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- [options]
 */
object TokenRunner extends IOApp:

  private val Model = "claude-sonnet-4-5-20250929"
  private val DefaultSamplePath = Paths.get("examples/anthropic-sdk/sample.xlsx")

  override def run(args: List[String]): IO[ExitCode] =
    parseArgs(args)
      .flatMap { config =>
        AnthropicClientIO.fromEnv.use { client =>
          runBenchmark(client, config)
        }
      }
      .as(ExitCode.Success)
      .handleErrorWith { e =>
        IO.println(s"Error: ${e.getMessage}").as(ExitCode.Error)
      }

  /** Run the benchmark */
  private def runBenchmark(client: AnthropicClientIO, config: TokenConfig): IO[Unit] =
    for
      _ <- BenchmarkUtils.printHeader("Token Efficiency Benchmark: xl CLI vs Anthropic xlsx Skill")

      // Load API key for Skills API
      apiKey <- AnthropicClientIO.loadApiKey

      // Resolve sample file
      samplePath = config.samplePath.getOrElse(DefaultSamplePath)
      _ <- IO.raiseWhen(!Files.exists(samplePath))(
        AgentError.ConfigError(
          s"Sample file not found: $samplePath\nGenerate with: scala-cli run examples/anthropic-sdk/create_sample.sc"
        )
      )

      // Upload sample file
      _ <- IO.println(s"Uploading ${samplePath.getFileName}...")
      sampleFile <- client.uploadFile(samplePath)
      _ <- IO.println(s"  Uploaded: ${sampleFile.id}")

      // Setup for xl approach (if not xlsx-only)
      xlSetup <-
        if config.xlsxOnly then IO.pure(None)
        else
          for
            _ <- IO.println("Setting up xl-cli skill...")
            binaryPath <- FileManager.resolveBinaryPath(config.xlBinary)
            skillPath <- FileManager.resolveSkillPath(config.xlSkill)
            _ <- IO.println(s"  Binary: ${binaryPath.getFileName}")
            _ <- IO.println(s"  Skill: ${skillPath.getFileName}")

            binaryFile <- client.uploadFile(binaryPath)
            _ <- IO.println(s"  Binary uploaded: ${binaryFile.id}")

            skillId <- SkillsApi.getOrCreateXlSkill(apiKey, skillPath)
          yield Some((binaryFile, skillId, FileManager.getFilename(binaryPath)))

      // Get tasks
      tasks = config.taskId match
        case Some(id) => TokenTasks.findById(id).toList
        case None => TokenTasks.allTasks(config.includeLarge)

      _ <- IO.raiseWhen(tasks.isEmpty)(AgentError.ConfigError("No tasks to run"))
      _ <- IO.println(s"\nRunning ${tasks.size} task(s)...")
      _ <- BenchmarkUtils.printSeparator()

      // Run tasks
      results <-
        if config.sequential then
          runTasksSequential(client.underlying, config, tasks, sampleFile, xlSetup)
        else runTasksParallel(client.underlying, config, tasks, sampleFile, xlSetup)

      // Grade results if enabled
      gradedResults <-
        if config.grading && !config.xlsxOnly then
          IO.println("\nGrading responses with Opus 4.5...") *>
            TokenGrader.gradeAll(client.underlying, tasks, results)
        else IO.pure(results)

      // Print comparison table
      _ <- TokenReporter.printComparisonTable(gradedResults)

      // Save report
      outputPath = config.outputPath.getOrElse(
        Paths.get("results", s"token_benchmark_${BenchmarkUtils.formatTimestamp}.json")
      )
      report = TokenReporter.buildReport(gradedResults, Model, samplePath.toString)
      _ <- TokenReporter.saveJsonReport(outputPath, report)

      // Cleanup
      _ <- client.deleteFile(sampleFile.id).attempt
      _ <- xlSetup.traverse_ { case (binaryFile, _, _) =>
        client.deleteFile(binaryFile.id).attempt
      }
    yield ()

  /** Run tasks sequentially */
  private def runTasksSequential(
    client: JAnthropicClient,
    config: TokenConfig,
    tasks: List[TokenTask],
    sampleFile: UploadedFile,
    xlSetup: Option[(UploadedFile, String, String)]
  ): IO[List[TokenTaskResult]] =
    tasks.flatTraverse { task =>
      for
        _ <- IO.println(s"\n[${task.id}] ${task.name}")
        results <- runTaskPair(client, config, task, sampleFile, xlSetup)
      yield results
    }

  /** Run tasks in parallel with concurrency limit */
  private def runTasksParallel(
    client: JAnthropicClient,
    config: TokenConfig,
    tasks: List[TokenTask],
    sampleFile: UploadedFile,
    xlSetup: Option[(UploadedFile, String, String)]
  ): IO[List[TokenTaskResult]] =
    Semaphore[IO](4).flatMap { sem =>
      tasks.parFlatTraverse { task =>
        sem.permit.use { _ =>
          for
            _ <- IO.println(s"[${task.id}] Starting...")
            results <- runTaskPair(client, config, task, sampleFile, xlSetup)
          yield results
        }
      }
    }

  /** Run a single task with both approaches */
  private def runTaskPair(
    client: JAnthropicClient,
    config: TokenConfig,
    task: TokenTask,
    sampleFile: UploadedFile,
    xlSetup: Option[(UploadedFile, String, String)]
  ): IO[List[TokenTaskResult]] =
    val xlResult = xlSetup match
      case Some((binaryFile, skillId, binaryName)) if !config.xlsxOnly =>
        runXlTask(client, task, sampleFile, binaryFile, skillId, binaryName, config.verbose).map(
          List(_)
        )
      case _ => IO.pure(Nil)

    val xlsxResult =
      if !config.xlOnly then runXlsxTask(client, task, sampleFile, config.verbose).map(List(_))
      else IO.pure(Nil)

    (xlResult, xlsxResult).parMapN(_ ++ _)

  /** Run task with xl CLI approach */
  private def runXlTask(
    client: JAnthropicClient,
    task: TokenTask,
    sampleFile: UploadedFile,
    binaryFile: UploadedFile,
    skillId: String,
    binaryName: String,
    verbose: Boolean
  ): IO[TokenTaskResult] =
    val startTime = System.currentTimeMillis()

    IO.blocking {
      val systemPrompt = s"""You have access to the xl CLI tool for Excel operations.

The xl binary has been uploaded to /mnt/user/$binaryName - make it executable first.
The Excel file is at /mnt/user/${sampleFile.filename}

Use xl commands to complete the task. Be concise in your response."""

      // Build custom skill for xl CLI
      val xlSkill = BetaSkillParams
        .builder()
        .skillId(skillId)
        .`type`(BetaSkillParams.Type.CUSTOM)
        .version("latest")
        .build()

      val container = BetaContainerParams
        .builder()
        .addSkill(xlSkill)
        .build()

      val codeExecutionTool = BetaCodeExecutionTool20250825.builder().build()

      val params = MessageCreateParams
        .builder()
        .model(Model)
        .maxTokens(2048L)
        .system(systemPrompt)
        .addUserMessage("placeholder")
        .container(container)
        .addTool(codeExecutionTool)
        .addBeta(AnthropicBeta.of("code-execution-2025-08-25"))
        .addBeta(AnthropicBeta.SKILLS_2025_10_02)
        .addBeta(AnthropicBeta.FILES_API_2025_04_14)
        // Override messages to include container_upload blocks
        .putAdditionalBodyProperty(
          "messages",
          JsonValue.from(
            java.util.List.of(
              java.util.Map.of(
                "role",
                "user",
                "content",
                java.util.List.of(
                  java.util.Map.of("type", "text", "text", task.xlPrompt),
                  java.util.Map.of("type", "container_upload", "file_id", binaryFile.id),
                  java.util.Map.of("type", "container_upload", "file_id", sampleFile.id)
                )
              )
            )
          )
        )
        .build()

      val response = client.beta().messages().create(params)
      val latencyMs = System.currentTimeMillis() - startTime

      val responseText = response
        .content()
        .asScala
        .flatMap(_.text().toScala)
        .map(_.text())
        .mkString("\n")

      TokenTaskResult(
        taskId = task.id,
        taskName = task.name,
        approach = Approach.Xl,
        success = true,
        inputTokens = response.usage().inputTokens(),
        outputTokens = response.usage().outputTokens(),
        totalTokens = response.usage().inputTokens() + response.usage().outputTokens(),
        latencyMs = latencyMs,
        responseText = Some(responseText)
      )
    }.handleErrorWith { e =>
      val latencyMs = System.currentTimeMillis() - startTime
      IO.pure(TokenTaskResult.failed(task.id, task.name, Approach.Xl, e.getMessage, latencyMs))
    }.flatTap { result =>
      IO.println(
        s"  [${task.id}] xl: ${BenchmarkUtils.formatTokens(result.inputTokens)} in / ${BenchmarkUtils.formatTokens(result.outputTokens)} out"
      )
    }

  /** Run task with Anthropic xlsx skill approach */
  private def runXlsxTask(
    client: JAnthropicClient,
    task: TokenTask,
    sampleFile: UploadedFile,
    verbose: Boolean
  ): IO[TokenTaskResult] =
    val startTime = System.currentTimeMillis()

    IO.blocking {
      val systemPrompt = s"""You have access to the xlsx skill for Excel operations.

The Excel file is at /mnt/user/${sampleFile.filename}

Use Python with openpyxl to complete the task. Be concise in your response."""

      // Build built-in xlsx skill
      val xlsxSkill = BetaSkillParams
        .builder()
        .skillId("xlsx")
        .`type`(BetaSkillParams.Type.ANTHROPIC)
        .version("latest")
        .build()

      val container = BetaContainerParams
        .builder()
        .addSkill(xlsxSkill)
        .build()

      val codeExecutionTool = BetaCodeExecutionTool20250825.builder().build()

      val params = MessageCreateParams
        .builder()
        .model(Model)
        .maxTokens(2048L)
        .system(systemPrompt)
        .addUserMessage("placeholder")
        .container(container)
        .addTool(codeExecutionTool)
        .addBeta(AnthropicBeta.of("code-execution-2025-08-25"))
        .addBeta(AnthropicBeta.SKILLS_2025_10_02)
        .addBeta(AnthropicBeta.FILES_API_2025_04_14)
        // Override messages to include container_upload blocks
        .putAdditionalBodyProperty(
          "messages",
          JsonValue.from(
            java.util.List.of(
              java.util.Map.of(
                "role",
                "user",
                "content",
                java.util.List.of(
                  java.util.Map.of("type", "text", "text", task.xlsxPrompt),
                  java.util.Map.of("type", "container_upload", "file_id", sampleFile.id)
                )
              )
            )
          )
        )
        .build()

      val response = client.beta().messages().create(params)
      val latencyMs = System.currentTimeMillis() - startTime

      val responseText = response
        .content()
        .asScala
        .flatMap(_.text().toScala)
        .map(_.text())
        .mkString("\n")

      TokenTaskResult(
        taskId = task.id,
        taskName = task.name,
        approach = Approach.Xlsx,
        success = true,
        inputTokens = response.usage().inputTokens(),
        outputTokens = response.usage().outputTokens(),
        totalTokens = response.usage().inputTokens() + response.usage().outputTokens(),
        latencyMs = latencyMs,
        responseText = Some(responseText)
      )
    }.handleErrorWith { e =>
      val latencyMs = System.currentTimeMillis() - startTime
      IO.pure(TokenTaskResult.failed(task.id, task.name, Approach.Xlsx, e.getMessage, latencyMs))
    }.flatTap { result =>
      IO.println(
        s"  [${task.id}] xlsx: ${BenchmarkUtils.formatTokens(result.inputTokens)} in / ${BenchmarkUtils.formatTokens(result.outputTokens)} out"
      )
    }

  /** Parse command line arguments */
  private def parseArgs(args: List[String]): IO[TokenConfig] =
    IO {
      var config = TokenConfig()
      var remaining = args

      while remaining.nonEmpty do
        remaining match
          case "--task" :: value :: rest =>
            config = config.copy(taskId = Some(value))
            remaining = rest
          case "--xl-only" :: rest =>
            config = config.copy(xlOnly = true)
            remaining = rest
          case "--xlsx-only" :: rest =>
            config = config.copy(xlsxOnly = true)
            remaining = rest
          case "--sequential" :: rest =>
            config = config.copy(sequential = true)
            remaining = rest
          case "--no-grade" :: rest =>
            config = config.copy(grading = false)
            remaining = rest
          case "--large" :: rest =>
            config = config.copy(includeLarge = true)
            remaining = rest
          case "--output" :: value :: rest =>
            config = config.copy(outputPath = Some(Paths.get(value)))
            remaining = rest
          case "--xl-binary" :: value :: rest =>
            config = config.copy(xlBinary = Some(Paths.get(value)))
            remaining = rest
          case "--xl-skill" :: value :: rest =>
            config = config.copy(xlSkill = Some(Paths.get(value)))
            remaining = rest
          case "--sample" :: value :: rest =>
            config = config.copy(samplePath = Some(Paths.get(value)))
            remaining = rest
          case "--verbose" :: rest =>
            config = config.copy(verbose = true)
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

  private val HelpText: String = """
Token Efficiency Benchmark - Compare xl CLI vs Anthropic xlsx Skill

Usage:
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- [options]

Task Selection:
  --task <id>         Run a specific task by ID
  --large             Include large file tasks

Approach Selection:
  --xl-only           Only run xl CLI tests (skip xlsx skill)
  --xlsx-only         Only run xlsx skill tests (skip xl)

Execution:
  --sequential        Run tasks sequentially (default: parallel)
  --no-grade          Skip correctness grading with Opus 4.5

Output:
  --output <path>     Output file for JSON results
  --verbose           Show detailed output

Development:
  --xl-binary <path>  Use local xl binary instead of released version
  --xl-skill <path>   Use local xl skill zip instead of released version
  --sample <path>     Custom sample.xlsx path

Examples:
  # Run all standard tasks
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner

  # Run a single task
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- --task list_sheets

  # Run xl-only without grading
  ./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- --xl-only --no-grade
"""
