package com.tjclp.xl.agent.benchmark.execution

import cats.effect.IO
import cats.syntax.all.*
import com.tjclp.xl.agent.AgentConfig
import com.tjclp.xl.agent.anthropic.AnthropicClientIO
import com.tjclp.xl.agent.benchmark.grading.*
import com.tjclp.xl.agent.benchmark.reporting.TokenSummary
import com.tjclp.xl.agent.benchmark.skills.{Skill, SkillContext}
import com.tjclp.xl.agent.benchmark.task.*

import java.nio.file.Path
import java.time.{Duration, Instant}

// ============================================================================
// Benchmark Engine - Unified Execution Entry Point
// ============================================================================

/**
 * Configuration for the benchmark engine.
 *
 * @param parallelism
 *   Number of parallel task executions (default: 4)
 * @param graderType
 *   Type of grading to apply (File, LLM, or FileAndLLM)
 * @param enableTracing
 *   Whether to save conversation traces
 * @param outputDir
 *   Base directory for output files
 * @param stream
 *   Whether to stream results in real-time
 * @param verbose
 *   Enable verbose logging
 */
case class EngineConfig(
  parallelism: Int = 4,
  graderType: GraderType = GraderType.File,
  enableTracing: Boolean = true,
  outputDir: Path = Path.of("results"),
  stream: Boolean = false,
  verbose: Boolean = false,
  xlCliPath: String = "xl"
)

object EngineConfig:
  val default: EngineConfig = EngineConfig()

/**
 * Result of a complete benchmark run.
 *
 * @param startTime
 *   When the benchmark started
 * @param endTime
 *   When the benchmark completed
 * @param config
 *   Configuration used for the run
 * @param skillResults
 *   Results grouped by skill
 * @param tasks
 *   The tasks that were executed
 */
case class BenchmarkRun(
  startTime: Instant,
  endTime: Instant,
  config: EngineConfig,
  skillResults: Map[String, SkillRunResult],
  tasks: List[BenchmarkTask]
):
  /** Total duration of the benchmark run */
  def duration: Duration = Duration.between(startTime, endTime)

  /** All execution results across all skills */
  def allResults: Vector[ExecutionResult] =
    skillResults.values.flatMap(_.results).toVector

  /** Overall pass rate across all skills */
  def overallPassRate: Double =
    val all = allResults
    if all.isEmpty then 0.0 else all.count(_.passed).toDouble / all.size

  /** Total token usage across all skills */
  def totalUsage: TokenUsage =
    import cats.kernel.Monoid
    Monoid[TokenUsage].combineAll(allResults.map(_.usage))

/**
 * Results for a single skill's benchmark run.
 *
 * @param skill
 *   The skill name
 * @param displayName
 *   Human-readable skill name
 * @param results
 *   Individual execution results for each task
 * @param summary
 *   Aggregated summary statistics
 */
case class SkillRunResult(
  skill: String,
  displayName: String,
  results: Vector[ExecutionResult],
  summary: SkillSummary
)

/**
 * Summary statistics for a skill's benchmark run.
 */
case class SkillSummary(
  total: Int,
  passed: Int,
  failed: Int,
  totalUsage: TokenUsage,
  avgLatencyMs: Double,
  estimatedCost: Option[BigDecimal]
):
  def passRate: Double = if total > 0 then passed.toDouble / total else 0.0
  def passRatePercent: Int = (passRate * 100).toInt

object SkillSummary:
  def fromResults(
    results: Vector[ExecutionResult],
    pricing: Option[ModelPricing] = None
  ): SkillSummary =
    val completed = results.filterNot(_.error.isDefined)
    val passed = completed.count(_.passed)
    val totalUsage = {
      import cats.kernel.Monoid
      Monoid[TokenUsage].combineAll(completed.map(_.usage))
    }
    val avgLatency =
      if completed.isEmpty then 0.0 else completed.map(_.latencyMs).sum.toDouble / completed.size

    SkillSummary(
      total = completed.length,
      passed = passed,
      failed = completed.length - passed,
      totalUsage = totalUsage,
      avgLatencyMs = avgLatency,
      estimatedCost = pricing.map(p => totalUsage.estimatedCost(p))
    )

// ============================================================================
// Benchmark Engine Trait
// ============================================================================

/**
 * The unified benchmark execution engine.
 *
 * This is the single entry point for running benchmarks. It:
 *   - Executes tasks across multiple skills in parallel
 *   - Handles grading (file, LLM, or both)
 *   - Produces unified ExecutionResult output
 *   - Supports streaming progress updates
 */
trait BenchmarkEngine:
  /**
   * Run a benchmark with the given tasks and skills.
   *
   * @param tasks
   *   Tasks to execute
   * @param skills
   *   Skills to use for execution
   * @param agentConfig
   *   Configuration for the agent
   * @param config
   *   Engine configuration
   * @return
   *   Complete benchmark run results
   */
  def run(
    tasks: List[BenchmarkTask],
    skills: List[Skill],
    agentConfig: AgentConfig,
    config: EngineConfig
  ): IO[BenchmarkRun]

  /**
   * Run a benchmark with streaming progress callbacks.
   *
   * @param tasks
   *   Tasks to execute
   * @param skills
   *   Skills to use for execution
   * @param agentConfig
   *   Configuration for the agent
   * @param config
   *   Engine configuration
   * @param onResult
   *   Callback for each execution result
   * @return
   *   Complete benchmark run results
   */
  def runStreaming(
    tasks: List[BenchmarkTask],
    skills: List[Skill],
    agentConfig: AgentConfig,
    config: EngineConfig,
    onResult: ExecutionResult => IO[Unit]
  ): IO[BenchmarkRun]

// ============================================================================
// Default Implementation
// ============================================================================

object BenchmarkEngine:

  /**
   * Create a default benchmark engine.
   *
   * @param client
   *   Anthropic client for API calls
   * @param fileGrader
   *   Optional file grader for spreadsheet comparison
   * @param llmGrader
   *   Optional LLM grader for response evaluation
   */
  def default(
    client: AnthropicClientIO,
    fileGrader: Option[Grader[? <: Score]] = None,
    llmGrader: Option[Grader[? <: Score]] = None
  ): BenchmarkEngine = new DefaultBenchmarkEngine(client, fileGrader, llmGrader)

private class DefaultBenchmarkEngine(
  client: AnthropicClientIO,
  fileGrader: Option[Grader[? <: Score]],
  llmGrader: Option[Grader[? <: Score]]
) extends BenchmarkEngine:

  override def run(
    tasks: List[BenchmarkTask],
    skills: List[Skill],
    agentConfig: AgentConfig,
    config: EngineConfig
  ): IO[BenchmarkRun] =
    runStreaming(tasks, skills, agentConfig, config, _ => IO.unit)

  override def runStreaming(
    tasks: List[BenchmarkTask],
    skills: List[Skill],
    agentConfig: AgentConfig,
    config: EngineConfig,
    onResult: ExecutionResult => IO[Unit]
  ): IO[BenchmarkRun] =
    for
      startTime <- IO(Instant.now())

      // Setup all skills first
      skillContexts <- skills.traverse { skill =>
        skill.setup(client, agentConfig).map(ctx => skill -> ctx)
      }

      // Create task-skill combinations
      combinations =
        for
          task <- tasks
          (skill, ctx) <- skillContexts
        yield (task, skill, ctx)

      // Execute with parallelism using parTraverseN
      results <- combinations
        .grouped(config.parallelism.max(1))
        .toList
        .flatTraverse { batch =>
          batch.parTraverse { case (task, skill, ctx) =>
            executeTaskWithSkill(task, skill, ctx, agentConfig, config).flatTap(onResult)
          }
        }

      // Teardown all skills
      _ <- skillContexts.traverse_ { case (skill, ctx) =>
        skill.teardown(client, ctx).attempt.void
      }

      // Group results by skill and build summaries
      bySkill = results.groupBy(_.skill)
      skillResults = skills.map { skill =>
        val skillResultsVec = bySkill.getOrElse(skill.name, Nil).toVector
        val pricing = Some(ModelPricing.forModel(agentConfig.model))
        skill.name -> SkillRunResult(
          skill = skill.name,
          displayName = skill.displayName,
          results = skillResultsVec,
          summary = SkillSummary.fromResults(skillResultsVec, pricing)
        )
      }.toMap

      endTime <- IO(Instant.now())
    yield BenchmarkRun(
      startTime = startTime,
      endTime = endTime,
      config = config,
      skillResults = skillResults,
      tasks = tasks
    )

  /**
   * Execute a single task with a specific skill.
   */
  private def executeTaskWithSkill(
    task: BenchmarkTask,
    skill: Skill,
    ctx: SkillContext,
    agentConfig: AgentConfig,
    config: EngineConfig
  ): IO[ExecutionResult] =
    skill
      .execute(task, ctx, client, agentConfig, config)
      .flatMap { result =>
        // Apply grading based on task's evaluation spec
        applyGrading(task, result, config)
      }
      .handleErrorWith { e =>
        IO.pure(
          ExecutionResult.failed(
            taskId = task.id,
            skill = skill.name,
            error = e.getMessage
          )
        )
      }

  /**
   * Apply appropriate grading to an execution result.
   *
   * Grading is driven by CaseDetails (what the execution produced), not global config. This ensures
   * token tasks get LLM grading and file tasks get file grading automatically.
   */
  private def applyGrading(
    task: BenchmarkTask,
    result: ExecutionResult,
    config: EngineConfig
  ): IO[ExecutionResult] =
    // If result already has error, skip grading
    if result.error.isDefined then IO.pure(result)
    else
      // Infer grader from CaseDetails (what execution actually produced)
      val graderAndContext: Option[(Grader[? <: Score], GraderContext)] =
        result.caseResults.headOption.flatMap { caseResult =>
          caseResult.details match
            case CaseDetails.FileComparison(_) =>
              // File grading already done during execution (inline in skill)
              // Could add post-hoc file grading here if needed
              None

            case CaseDetails.TokenComparison(response, expected) =>
              // LLM grading for analysis/token tasks
              for
                grader <- llmGrader
                exp <- expected
              yield (
                grader,
                GraderContext.forLLM(
                  taskId = result.taskIdValue,
                  skill = result.skill,
                  responseText = response,
                  expectedAnswer = exp,
                  taskInstruction = Some(task.instruction)
                )
              )

            case CaseDetails.NoDetails =>
              None
        }

      graderAndContext match
        case Some((grader, ctx)) if grader.canHandle(ctx) =>
          grader.grade(ctx).map { gradeResult =>
            result.withGrade(gradeResult)
          }
        case _ =>
          IO.pure(result)
