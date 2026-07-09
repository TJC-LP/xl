package com.tjclp.xl.agent.benchmark.execution

import cats.effect.IO
import munit.CatsEffectSuite
import com.tjclp.xl.agent.{Agent, AgentConfig, AgentResult, AgentTask}
import com.tjclp.xl.agent.anthropic.AnthropicClientIO
import com.tjclp.xl.agent.benchmark.skills.{Skill, SkillContext}
import com.tjclp.xl.agent.benchmark.task.*

import java.nio.file.{Files, Path}

/**
 * Engine last-resort failure path (issue #340): a third-party Skill that raises outside its own
 * guarded sections must still produce a total failing CaseResult AND leave a metadata-only trace on
 * disk (no tracer exists at the engine level, so events are unrecoverable by design).
 *
 * Runs fully offline: the stub skill raises before the dummy-key client could reach the network.
 */
class BenchmarkEngineSpec extends CatsEffectSuite:

  private val tempDir = FunFixture[Path](
    setup = _ => Files.createTempDirectory("benchmark-engine-spec"),
    teardown = { dir =>
      import scala.jdk.CollectionConverters.*
      import scala.util.Using
      Using.resource(Files.walk(dir)) { stream =>
        stream.iterator().asScala.toList.reverse.foreach(p => Files.deleteIfExists(p))
      }
    }
  )

  /** A third-party-style skill that raises outside any guarded section */
  private object ExplodingSkill extends Skill:
    override val name: String = "exploding"
    override val displayName: String = "Exploding Skill"

    override def setup(client: AnthropicClientIO, config: AgentConfig): IO[SkillContext] =
      IO.pure(SkillContext.empty)

    override def teardown(client: AnthropicClientIO, ctx: SkillContext): IO[Unit] = IO.unit

    override def createAgent(
      client: AnthropicClientIO,
      ctx: SkillContext,
      config: AgentConfig
    ): Agent = new Agent:
      override def run(task: AgentTask): IO[AgentResult] =
        IO.raiseError(new IllegalStateException("unused"))
      override def runStreaming(
        task: AgentTask,
        onEvent: com.tjclp.xl.agent.AgentEvent => IO[Unit]
      ): IO[AgentResult] =
        IO.raiseError(new IllegalStateException("unused"))

    override def execute(
      task: BenchmarkTask,
      ctx: SkillContext,
      client: AnthropicClientIO,
      agentConfig: AgentConfig,
      engineConfig: EngineConfig
    ): IO[ExecutionResult] =
      IO.raiseError(new IllegalStateException("skill exploded outside its guard"))

    override def executeCase(
      testCase: TestCaseFile,
      task: BenchmarkTask,
      ctx: SkillContext,
      client: AnthropicClientIO,
      agentConfig: AgentConfig,
      engineConfig: EngineConfig
    ): IO[CaseResult] =
      IO.raiseError(new IllegalStateException("skill exploded outside its guard"))

  tempDir.test("a raising Skill yields a failing CaseResult with a metadata-only trace") { dir =>
    val testCase = TestCaseFile(1, dir.resolve("in.xlsx"), dir.resolve("answer.xlsx"))
    val task = BenchmarkTask(
      id = TaskId("999"),
      instruction = "test instruction",
      category = TaskCategory.CellLevel,
      inputSource = InputSource.TestCases(Vector(testCase)),
      evaluation = EvaluationSpec.fileOnly
    )
    AnthropicClientIO.resource("dummy-key-never-used").use { client =>
      BenchmarkEngine
        .default(client)
        .run(
          tasks = List(task),
          skills = List(ExplodingSkill),
          agentConfig = AgentConfig(model = "claude-test"),
          config = EngineConfig.default.copy(outputDir = dir, stream = false)
        )
        .map { run =>
          val execResult = run
            .skillResults("exploding")
            .results
            .headOption
            .getOrElse(fail("missing execution result"))
          val caseResult =
            execResult.caseResults.headOption.getOrElse(fail("missing case result"))

          assertEquals(caseResult.passed, false)
          assert(
            caseResult.error.exists(_.contains("IllegalStateException: skill exploded")),
            s"error must be recorded: ${caseResult.error}"
          )

          // All cases errored -> the task itself is an execution failure (issue #340)
          assert(
            execResult.error.exists(_.contains("skill exploded")),
            s"all-errored task must carry a task-level error: ${execResult.error}"
          )

          // The engine fallback saved a metadata-only trace
          val traceDir =
            caseResult.tracePath.getOrElse(fail("engine fallback must save a trace"))
          val traceJson = Files.readString(traceDir.resolve("conversation.json"))
          val parsed =
            io.circe.parser.parse(traceJson).getOrElse(fail("unparseable trace JSON"))
          val meta = parsed.hcursor.downField("metadata")
          assertEquals(meta.get[Boolean]("passed"), Right(false))
          assertEquals(meta.get[String]("model"), Right("claude-test"))
          assert(
            meta.get[String]("error").exists(_.contains("skill exploded")),
            s"trace metadata must carry the error: $traceJson"
          )
          assertEquals(
            parsed.hcursor.downField("events").as[Vector[io.circe.Json]].map(_.size),
            Right(0),
            "engine fallback traces are metadata-only"
          )
        }
    }
  }
