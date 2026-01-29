package com.tjclp.xl.agent.benchmark.runner

import cats.effect.{ExitCode, IO, IOApp}
import cats.syntax.all.*
import com.anthropic.client.AnthropicClient as JAnthropicClient
import com.tjclp.xl.agent.anthropic.AnthropicClientIO
import com.tjclp.xl.agent.benchmark.common.BenchmarkUtils
import com.tjclp.xl.agent.benchmark.grading.*
import com.tjclp.xl.agent.benchmark.reporting.*
import com.tjclp.xl.agent.benchmark.skills.{Skill, SkillRegistry}
import com.tjclp.xl.agent.benchmark.task.*
import com.tjclp.xl.agent.benchmark.task.UnifiedTaskLoader.*

import java.nio.file.Path
import java.time.Instant

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
      report <- runBenchmark(config)
      _ <- writeReport(config, report)
    yield ExitCode.Success

  /** Run the benchmark with the given configuration */
  def runBenchmark(config: UnifiedConfig): IO[BenchmarkReport] =
    AnthropicClientIO.fromEnv.use { client =>
      for
        startTime <- IO(Instant.now())
        tasks <- loadTasks(config)
        _ <- IO.println(s"Loaded ${tasks.length} tasks")
        skills <- loadSkills(config)
        _ <- IO.println(s"Using skills: ${skills.map(_.name).mkString(", ")}")
        graders <- createGraders(config, client.underlying)
        results <- runAllTasks(config, tasks, skills, graders, client)
        report <- buildReport(config, startTime, skills.map(_.name), results)
      yield report
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
  // Grader Creation
  // --------------------------------------------------------------------------

  private def createGraders(
    config: UnifiedConfig,
    client: JAnthropicClient
  ): IO[Vector[Grader[? <: Score]]] =
    IO.pure {
      val fileGrader: Option[Grader[? <: Score]] =
        if config.graderType == GraderType.File || config.graderType == GraderType.FileAndLLM then
          Some(SpreadsheetBenchGrader(config.xlCliPath))
        else None

      val llmGrader: Option[Grader[? <: Score]] =
        if config.graderType == GraderType.LLM || config.graderType == GraderType.FileAndLLM then
          if config.enableGrading then Some(OpusLLMGrader(client))
          else Some(NoOpLLMGrader)
        else None

      Vector(fileGrader, llmGrader).flatten
    }

  // --------------------------------------------------------------------------
  // Task Execution
  // --------------------------------------------------------------------------

  private def runAllTasks(
    config: UnifiedConfig,
    tasks: List[BenchmarkTask],
    skills: List[Skill],
    graders: Vector[Grader[? <: Score]],
    client: AnthropicClientIO
  ): IO[List[TaskResultEntry]] =
    // Create task-skill combinations
    val combinations = for
      task <- tasks
      skill <- skills
    yield (task, skill)

    if config.stream then
      StreamingReportWriter.writeHeader(skills.map(_.name), combinations.length) >>
        combinations.traverse { case (task, skill) =>
          runSingleTask(config, task, skill, graders, client).flatTap { result =>
            StreamingReportWriter.writeResult(result)
          }
        }
    else
      // Run with parallelism
      combinations
        .grouped(config.parallelism.max(1))
        .toList
        .flatTraverse { batch =>
          batch.parTraverse { case (task, skill) =>
            runSingleTask(config, task, skill, graders, client)
          }
        }

  private def runSingleTask(
    config: UnifiedConfig,
    task: BenchmarkTask,
    skill: Skill,
    graders: Vector[Grader[? <: Score]],
    client: AnthropicClientIO
  ): IO[TaskResultEntry] =
    // For now, create a placeholder result
    // Full implementation would integrate with the existing skill execution
    val startTime = System.currentTimeMillis()

    task.inputSource match
      case InputSource.TestCases(cases) =>
        // Run each test case
        runTestCases(config, task, skill, cases.toList, graders, client)

      case InputSource.SingleFile(path) =>
        // Single file execution
        runSingleFileTask(config, task, skill, path, graders, client)

      case InputSource.NoInput =>
        // No input file needed
        runNoInputTask(config, task, skill, graders, client)

      case InputSource.DataDirectory(dir, pattern) =>
        // Dynamic file resolution
        IO.pure(
          TaskResultEntry.skipped(task.taskIdValue, skill.name, "DataDirectory not implemented")
        )

  private def runTestCases(
    config: UnifiedConfig,
    task: BenchmarkTask,
    skill: Skill,
    cases: List[TestCaseFile],
    graders: Vector[Grader[? <: Score]],
    client: AnthropicClientIO
  ): IO[TaskResultEntry] =
    // This would integrate with the existing SpreadsheetBench Runner logic
    // For now, return a placeholder that indicates integration needed
    IO.pure(
      TaskResultEntry(
        taskId = task.taskIdValue,
        skill = skill.name,
        caseNum = None,
        instruction = task.instruction.take(100),
        category = task.category.toString,
        score = Score.FractionalScore(0, cases.length),
        gradeDetails = GradeDetails.fail("Integration with existing runner needed"),
        usage = TokenSummary.zero,
        latencyMs = 0,
        error = Some("Use spreadsheetbench.Runner for full execution")
      )
    )

  private def runSingleFileTask(
    config: UnifiedConfig,
    task: BenchmarkTask,
    skill: Skill,
    inputPath: Path,
    graders: Vector[Grader[? <: Score]],
    client: AnthropicClientIO
  ): IO[TaskResultEntry] =
    // This would integrate with TokenRunner execution logic
    IO.pure(
      TaskResultEntry(
        taskId = task.taskIdValue,
        skill = skill.name,
        caseNum = None,
        instruction = task.instruction.take(100),
        category = task.category.toString,
        score = Score.BinaryScore.Fail,
        gradeDetails = GradeDetails.fail("Integration with existing runner needed"),
        usage = TokenSummary.zero,
        latencyMs = 0,
        error = Some("Use token.TokenRunner for full execution")
      )
    )

  private def runNoInputTask(
    config: UnifiedConfig,
    task: BenchmarkTask,
    skill: Skill,
    graders: Vector[Grader[? <: Score]],
    client: AnthropicClientIO
  ): IO[TaskResultEntry] =
    IO.pure(TaskResultEntry.skipped(task.taskIdValue, skill.name, "No input file"))

  // --------------------------------------------------------------------------
  // Report Building
  // --------------------------------------------------------------------------

  private def buildReport(
    config: UnifiedConfig,
    startTime: Instant,
    skills: List[String],
    results: List[TaskResultEntry]
  ): IO[BenchmarkReport] =
    IO {
      ReportBuilder()
        .withTitle(s"${config.benchmarkType} Benchmark Results")
        .withBenchmarkType(config.benchmarkType.toString)
        .withStartTime(startTime)
        .withModel(config.model)
        .withSkills(skills)
        .withMetadata(
          ReportMetadata(
            dataset = Some(config.dataset),
            dataDir = Some(config.dataDir.toString)
          )
        )
        .addResults(results)
        .build()
    }

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
      case _ => SpreadsheetBench
