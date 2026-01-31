package com.tjclp.xl.agent.benchmark.skills

import cats.effect.IO
import cats.syntax.all.*
import com.tjclp.xl.agent.{Agent, AgentConfig, AgentTask}
import com.tjclp.xl.agent.anthropic.AnthropicClientIO
import com.tjclp.xl.agent.benchmark.execution.{
  CaseResult,
  EngineConfig,
  ExecutionResult,
  TokenUsage
}
import com.tjclp.xl.agent.benchmark.task.{BenchmarkTask, InputSource, TestCaseFile}

import java.nio.file.Files
import scala.jdk.CollectionConverters.*

/** Context created during skill setup, containing resources to clean up */
case class SkillContext(
  fileIds: List[String] = Nil,
  skillId: Option[String] = None,
  metadata: Map[String, String] = Map.empty
)

object SkillContext:
  val empty: SkillContext = SkillContext()

/**
 * Skill abstraction for different benchmark approaches
 *
 * Skills encapsulate:
 *   - Setup: Upload files, register APIs, etc.
 *   - Teardown: Clean up uploaded resources
 *   - Agent creation: Configure agent with skill-specific prompts and tools
 *   - Execution: Run tasks and produce unified ExecutionResult
 *
 * This allows N-skill comparisons without hardcoding approaches.
 */
trait Skill:
  /** Unique identifier for this skill (e.g., "xl", "xlsx", "raw-python") */
  def name: String

  /** Human-readable name (e.g., "xl-cli", "OpenPyXL", "Raw Python") */
  def displayName: String

  /** One-time setup (upload files, register skill via API, etc.) */
  def setup(client: AnthropicClientIO, config: AgentConfig): IO[SkillContext]

  /** Cleanup resources (delete uploaded files, etc.) */
  def teardown(client: AnthropicClientIO, ctx: SkillContext): IO[Unit]

  /** Create an agent configured for this skill */
  def createAgent(
    client: AnthropicClientIO,
    ctx: SkillContext,
    config: AgentConfig
  ): Agent

  /**
   * Execute a benchmark task and return a unified ExecutionResult.
   *
   * This is the primary execution method used by BenchmarkEngine. Skills handle all InputSource
   * types (TestCases, SingleFile, NoInput) and produce consistent ExecutionResult output.
   *
   * Grading is determined by the task's EvaluationSpec, not the skill. The skill simply executes
   * and reports results; the engine handles grading.
   *
   * @param task
   *   The benchmark task to execute
   * @param ctx
   *   Skill context from setup
   * @param client
   *   Anthropic client for API calls
   * @param agentConfig
   *   Configuration for the agent
   * @param engineConfig
   *   Engine configuration for execution options
   * @return
   *   Unified execution result
   */
  def execute(
    task: BenchmarkTask,
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[ExecutionResult]

  /** Description for --list-skills output */
  def description: String = displayName

  // ============================================================================
  // Flattened Work Unit Support
  // ============================================================================

  /**
   * Enumerate test cases for a task without executing them.
   *
   * This allows the engine to flatten all (task, skill, case) combinations into independent work
   * units for parallel scheduling. The default implementation derives cases from InputSource.
   *
   * @param task
   *   The task to enumerate cases for
   * @return
   *   Vector of test case files to execute
   */
  def enumerateCases(task: BenchmarkTask): IO[Vector[TestCaseFile]] =
    task.inputSource match
      case InputSource.TestCases(cases) =>
        IO.pure(cases)

      case InputSource.SingleFile(path) =>
        // Single file tasks have exactly one "case"
        IO.pure(Vector(TestCaseFile(1, path, path)))

      case InputSource.NoInput =>
        // No input means no cases to run
        IO.pure(Vector.empty)

      case InputSource.DataDirectory(dir, pattern) =>
        // Resolve files dynamically using glob pattern
        IO.blocking {
          val glob = dir.getFileSystem.getPathMatcher(s"glob:$pattern")
          Files
            .walk(dir)
            .iterator()
            .asScala
            .filter(p => Files.isRegularFile(p) && glob.matches(p.getFileName))
            .zipWithIndex
            .map { (p, idx) => TestCaseFile(idx + 1, p, p) }
            .toVector
        }

  /**
   * Execute a single test case.
   *
   * This is called by the engine for each work unit. Skills must implement this method to handle
   * single-case execution. The default implementation delegates to the full execute() method for
   * backward compatibility, but skills should override this for better efficiency.
   *
   * @param testCase
   *   The specific test case to execute
   * @param task
   *   The benchmark task this case belongs to
   * @param ctx
   *   Skill context from setup
   * @param client
   *   Anthropic client for API calls
   * @param agentConfig
   *   Configuration for the agent
   * @param engineConfig
   *   Engine configuration for execution options
   * @return
   *   Result for this single case
   */
  def executeCase(
    testCase: TestCaseFile,
    task: BenchmarkTask,
    ctx: SkillContext,
    client: AnthropicClientIO,
    agentConfig: AgentConfig,
    engineConfig: EngineConfig
  ): IO[CaseResult]
