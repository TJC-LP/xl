package com.tjclp.xl.agent.benchmark.runner

import cats.effect.{ExitCode, IO, IOApp}
import cats.syntax.all.*
import com.anthropic.client.AnthropicClient as JAnthropicClient
import com.tjclp.xl.agent.AgentConfig
import com.tjclp.xl.agent.anthropic.AnthropicClientIO
import com.tjclp.xl.agent.benchmark.common.BenchmarkUtils
import com.tjclp.xl.agent.benchmark.execution.*
import com.tjclp.xl.agent.benchmark.grading.*
import com.tjclp.xl.agent.benchmark.reporting.*
import com.tjclp.xl.agent.benchmark.skills.{Skill, SkillRegistry}
import com.tjclp.xl.agent.benchmark.task.*
import com.tjclp.xl.agent.benchmark.task.UnifiedTaskLoader.*

import java.nio.file.Path
import java.time.{Duration, Instant}

// ============================================================================
// Unified Benchmark Runner
// ============================================================================

/** Single entry point for all benchmark types */
object UnifiedRunner extends IOApp:

  override def run(args: List[String]): IO[ExitCode] =
    for
      config <- parseArgs(args)
      _ <- config.listSkills.fold(IO.unit)(listSkillsAndExit)
      _ <- IO.println(s"${Console.BOLD}Unified Benchmark Runner${Console.RESET}")
      _ <- IO(SkillRegistry.initialize())
      benchmarkRun <- runBenchmarkWithEngine(config)
      _ <- writeUnifiedReport(config, benchmarkRun)
    yield ExitCode.Success

  /** Run the benchmark using BenchmarkEngine */
  def runBenchmarkWithEngine(config: UnifiedConfig): IO[BenchmarkRun] =
    AnthropicClientIO.fromEnv.use { client =>
      for
        tasks <- loadTasks(config)
        _ <- IO.println(s"Loaded ${tasks.length} tasks")
        skills <- loadSkills(config)
        _ <- IO.println(s"Using skills: ${skills.map(_.name).mkString(", ")}")

        // Create graders for the engine
        // Both graders are always available - CaseDetails determines which is used per-task
        fileGrader: Option[Grader[? <: Score]] =
          Some(SpreadsheetBenchGrader(config.xlCliPath))

        llmGrader: Option[Grader[? <: Score]] =
          if config.enableGrading then Some(OpusLLMGrader(client.underlying))
          else Some(NoOpLLMGrader)

        // Create engine and config
        engine = BenchmarkEngine.default(client, fileGrader, llmGrader)

        agentConfig = AgentConfig(
          model = config.model,
          verbose = false
        )

        engineConfig = EngineConfig(
          parallelism = config.parallelism,
          graderType = config.graderType,
          enableTracing = true,
          outputDir = config.outputDir,
          stream = config.stream,
          xlCliPath = config.xlCliPath
        )

        // Run with streaming callback if enabled
        benchmarkRun <-
          if config.stream then
            StreamingReportWriter.writeHeader(skills.map(_.name), tasks.length * skills.length) >>
              engine.runStreaming(
                tasks = tasks,
                skills = skills,
                agentConfig = agentConfig,
                config = engineConfig,
                onResult = result => IO.println(formatExecutionResult(result))
              )
          else engine.run(tasks, skills, agentConfig, engineConfig)

        _ <- printSummary(benchmarkRun)
      yield benchmarkRun
    }

  /** Format an ExecutionResult for streaming output */
  private def formatExecutionResult(result: ExecutionResult): String =
    val status =
      if result.passed then Console.GREEN + "✓" + Console.RESET
      else if result.error.isDefined then Console.RED + "✗" + Console.RESET
      else Console.YELLOW + "○" + Console.RESET

    f"$status ${result.skill}%-8s ${result.taskIdValue}%-12s " +
      f"${result.passedCases}/${result.totalCases} cases " +
      f"tokens=${result.totalTokens}%6d latency=${result.latencyMs}%5dms"

  /** Print summary after benchmark completion */
  private def printSummary(run: BenchmarkRun): IO[Unit] =
    IO.println(s"""
${Console.BOLD}${"=" * 60}${Console.RESET}
${Console.BOLD}Benchmark Complete${Console.RESET}
${"=" * 60}
Duration: ${formatDuration(run.duration)}
Tasks: ${run.tasks.length}
Overall Pass Rate: ${(run.overallPassRate * 100).toInt}%
Total Tokens: ${run.totalUsage.total}
${"-" * 60}
${run.skillResults.values
        .map { sr =>
          f"${sr.displayName}%-15s ${sr.summary.passed}/${sr.summary.total} passed (${sr.summary.passRatePercent}%%)"
        }
        .mkString("\n")}
${"=" * 60}
""")

  private def formatDuration(d: Duration): String =
    val secs = d.toSeconds
    val mins = secs / 60
    val remSecs = secs % 60
    if mins > 0 then f"${mins}m ${remSecs}s" else s"${secs}s"

  // Legacy method for backward compatibility
  def runBenchmark(config: UnifiedConfig): IO[BenchmarkReport] =
    runBenchmarkWithEngine(config).flatMap { run =>
      buildReportFromRun(config, run)
    }

  // --------------------------------------------------------------------------
  // Task Loading
  // --------------------------------------------------------------------------

  private def loadTasks(config: UnifiedConfig): IO[List[BenchmarkTask]] =
    val source = config.benchmarkType match
      case BenchmarkType.SpreadsheetBench =>
        BenchmarkSource.SpreadsheetBench(config.dataDir, config.dataset)
      case BenchmarkType.Token =>
        BenchmarkSource.TokenBenchmark(config.samplePath)
      case BenchmarkType.Custom =>
        config.tasksFile
          .map(p => BenchmarkSource.CustomFile(p))
          .getOrElse(BenchmarkSource.TokenBenchmark(None))

    val filter = TaskFilter(
      taskIds = config.taskIds,
      skipIds = config.skipIds,
      category = config.category,
      limit = config.limit,
      includeVba = config.includeVba
    )

    UnifiedTaskLoader.load(source, filter)

  // --------------------------------------------------------------------------
  // Skill Loading
  // --------------------------------------------------------------------------

  private def loadSkills(config: UnifiedConfig): IO[List[Skill]] =
    IO(config.skills.map(SkillRegistry.apply))

  // --------------------------------------------------------------------------
  // Report Building and Writing
  // --------------------------------------------------------------------------

  /** Build a BenchmarkReport from a BenchmarkRun (for legacy compatibility) */
  private def buildReportFromRun(config: UnifiedConfig, run: BenchmarkRun): IO[BenchmarkReport] =
    IO {
      // Convert ExecutionResults to TaskResultEntries
      val entries = run.allResults.flatMap { result =>
        result.caseResults.map { caseResult =>
          TaskResultEntry(
            taskId = result.taskIdValue,
            skill = result.skill,
            caseNum = Some(caseResult.caseNum),
            instruction = run.tasks
              .find(_.taskIdValue == result.taskIdValue)
              .map(_.instruction.take(100))
              .getOrElse(""),
            category = run.tasks
              .find(_.taskIdValue == result.taskIdValue)
              .map(_.category.toString)
              .getOrElse(""),
            score = Score.BinaryScore.fromBoolean(caseResult.passed),
            gradeDetails =
              if caseResult.passed then GradeDetails.passed
              else GradeDetails.fail("Case failed").withMismatches(caseResult.mismatches),
            usage = TokenSummary(
              inputTokens = caseResult.usage.inputTokens,
              outputTokens = caseResult.usage.outputTokens
            ),
            latencyMs = caseResult.latencyMs,
            error = result.error
          )
        }
      }.toList

      ReportBuilder()
        .withTitle(s"${config.benchmarkType} Benchmark Results")
        .withBenchmarkType(config.benchmarkType.toString)
        .withStartTime(run.startTime)
        .withModel(config.model)
        .withSkills(run.skillResults.keys.toList)
        .withMetadata(
          ReportMetadata(
            dataset = Some(config.dataset),
            dataDir = Some(config.dataDir.toString)
          )
        )
        .addResults(entries)
        .build()
    }

  /** Write unified report using UnifiedReportWriter */
  private def writeUnifiedReport(config: UnifiedConfig, run: BenchmarkRun): IO[Unit] =
    for
      paths <- UnifiedReportWriter.write(run, config.outputDir)
      _ <- IO.println(s"\nReport written to:")
      _ <- IO.println(s"  JSON: ${paths.json}")
      _ <- IO.println(s"  Markdown: ${paths.markdown}")
    yield ()

  /** Legacy writeReport for backward compatibility */
  private def writeReport(config: UnifiedConfig, report: BenchmarkReport): IO[Unit] =
    for
      paths <- ReportWriter.write(config.outputDir, report)
      _ <- IO.println(s"\nReport written to:")
      _ <- IO.println(s"  JSON: ${paths.json}")
      _ <- IO.println(s"  Markdown: ${paths.markdown}")
    yield ()

  // --------------------------------------------------------------------------
  // CLI Argument Parsing
  // --------------------------------------------------------------------------

  private def parseArgs(args: List[String]): IO[UnifiedConfig] =
    IO {
      @annotation.tailrec
      def parse(remaining: List[String], config: UnifiedConfig): UnifiedConfig =
        remaining match
          case Nil => config
          case "--benchmark" :: value :: rest =>
            parse(rest, config.copy(benchmarkType = BenchmarkType.fromString(value)))
          case "--grader" :: value :: rest =>
            parse(
              rest,
              config.copy(graderType = GraderType.fromString(value).getOrElse(GraderType.File))
            )
          case "--skills" :: value :: rest =>
            parse(rest, config.copy(skills = value.split(",").toList))
          case "--task" :: value :: rest =>
            parse(rest, config.copy(taskIds = Some(value.split(",").toList)))
          case "--category" :: value :: rest =>
            parse(rest, config.copy(category = Some(value)))
          case "--limit" :: value :: rest =>
            parse(rest, config.copy(limit = Some(value.toInt)))
          case "--output" :: value :: rest =>
            parse(rest, config.copy(outputDir = Path.of(value)))
          case "--data-dir" :: value :: rest =>
            parse(rest, config.copy(dataDir = Path.of(value)))
          case "--dataset" :: value :: rest =>
            parse(rest, config.copy(dataset = value))
          case "--sample" :: value :: rest =>
            parse(rest, config.copy(samplePath = Some(Path.of(value))))
          case "--tasks-file" :: value :: rest =>
            parse(rest, config.copy(tasksFile = Some(Path.of(value))))
          case "--model" :: value :: rest =>
            parse(rest, config.copy(model = value))
          case "--stream" :: rest =>
            parse(rest, config.copy(stream = true))
          case "--no-grade" :: rest =>
            parse(rest, config.copy(enableGrading = false))
          case "--include-vba" :: rest =>
            parse(rest, config.copy(includeVba = true))
          case "--parallelism" :: value :: rest =>
            parse(rest, config.copy(parallelism = value.toInt))
          case "--list-skills" :: rest =>
            parse(rest, config.copy(listSkills = Some(())))
          case "--xl-cli" :: value :: rest =>
            parse(rest, config.copy(xlCliPath = value))
          case ("--help" | "-h") :: _ =>
            printHelp()
            throw new RuntimeException("Help requested")
          case arg :: _ if arg.startsWith("-") =>
            throw new IllegalArgumentException(s"Unknown argument: $arg")
          case _ :: rest => parse(rest, config) // Skip positional args

      parse(args, UnifiedConfig())
    }

  private def printHelp(): Unit =
    println("""
      |Unified Benchmark Runner
      |
      |Usage: UnifiedRunner [options]
      |
      |Benchmark Type:
      |  --benchmark <type>     Benchmark type: spreadsheetbench, token, custom
      |
      |Grading:
      |  --grader <type>        Grader type: file, llm, file,llm (both)
      |  --no-grade             Disable LLM grading
      |
      |Skills:
      |  --skills <list>        Comma-separated skills: xl,xlsx
      |  --list-skills          List available skills and exit
      |
      |Task Selection:
      |  --task <id>            Run specific task(s), comma-separated
      |  --category <cat>       Filter by category
      |  --limit <n>            Limit number of tasks
      |  --include-vba          Include VBA tasks (skipped by default)
      |
      |Data Sources:
      |  --data-dir <path>      SpreadsheetBench data directory
      |  --dataset <name>       Dataset name (default: sample_data_200)
      |  --sample <path>        Sample file for token benchmark
      |  --tasks-file <path>    Custom tasks JSON file
      |
      |Output:
      |  --output <dir>         Output directory (default: results/)
      |  --stream               Real-time streaming output
      |
      |Execution:
      |  --model <name>         Model to use (default: claude-sonnet-4-20250514)
      |  --parallelism <n>      Number of parallel tasks (default: 4)
      |  --xl-cli <path>        Path to xl CLI (default: xl)
      |""".stripMargin)

  private def listSkillsAndExit(unit: Unit): IO[Nothing] =
    for
      _ <- IO(SkillRegistry.initialize())
      skills <- IO(SkillRegistry.all)
      _ <- IO.println("Available skills:")
      _ <- skills.traverse_(s => IO.println(s"  ${s.name}: ${s.description}"))
    yield throw new RuntimeException("Listed skills")

// ============================================================================
// Configuration
// ============================================================================

/** Unified benchmark configuration */
case class UnifiedConfig(
  benchmarkType: BenchmarkType = BenchmarkType.SpreadsheetBench,
  graderType: GraderType = GraderType.File,
  skills: List[String] = List("xl"),
  taskIds: Option[List[String]] = None,
  skipIds: Set[String] = Set.empty,
  category: Option[String] = None,
  limit: Option[Int] = None,
  includeVba: Boolean = false,
  dataDir: Path =
    Path.of(Option(System.getenv("HOME")).getOrElse("."), "git/SpreadsheetBench/data"),
  dataset: String = "sample_data_200",
  samplePath: Option[Path] = None,
  tasksFile: Option[Path] = None,
  outputDir: Path = Path.of("results", BenchmarkUtils.formatTimestamp),
  model: String = "claude-sonnet-4-20250514",
  parallelism: Int = 4,
  stream: Boolean = false,
  enableGrading: Boolean = true,
  xlCliPath: String = "xl",
  listSkills: Option[Unit] = None
)

/** Benchmark type enum */
enum BenchmarkType derives CanEqual:
  case SpreadsheetBench
  case Token
  case Custom

object BenchmarkType:
  def fromString(s: String): BenchmarkType =
    s.toLowerCase match
      case "spreadsheetbench" | "spreadsheet" | "sb" => SpreadsheetBench
      case "token" | "tokenefficiency" => Token
      case "custom" => Custom
      case other =>
        throw new IllegalArgumentException(
          s"Unknown benchmark type: '$other'. Valid types: spreadsheetbench, token, custom"
        )
