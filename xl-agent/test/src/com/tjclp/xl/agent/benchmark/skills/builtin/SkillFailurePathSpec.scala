package com.tjclp.xl.agent.benchmark.skills.builtin

import munit.CatsEffectSuite
import com.tjclp.xl.agent.AgentConfig
import com.tjclp.xl.agent.anthropic.AnthropicClientIO
import com.tjclp.xl.agent.benchmark.execution.EngineConfig
import com.tjclp.xl.agent.benchmark.skills.SkillContext
import com.tjclp.xl.agent.benchmark.task.*

import java.nio.file.{Files, Path}

/**
 * End-to-end wiring check for the case failure path (issue #334): a real executeCase that raises
 * must return a total CaseResult carrying the error AND leave the partial conversation trace on
 * disk.
 *
 * Runs fully offline: the dummy-key client never reaches the network because the missing input file
 * makes the upload raise first.
 */
class SkillFailurePathSpec extends CatsEffectSuite:

  private val tempDir = FunFixture[Path](
    setup = _ => Files.createTempDirectory("skill-failure-path-spec"),
    teardown = { dir =>
      import scala.jdk.CollectionConverters.*
      import scala.util.Using
      Using.resource(Files.walk(dir)) { stream =>
        stream.iterator().asScala.toList.reverse.foreach(p => Files.deleteIfExists(p))
      }
    }
  )

  tempDir.test("a failing executeCase records the error and saves the partial trace") { dir =>
    val missingInput = dir.resolve("does-not-exist.xlsx")
    val testCase = TestCaseFile(3, missingInput, dir.resolve("answer.xlsx"))
    val task = BenchmarkTask(
      id = TaskId("13894"),
      instruction = "test instruction",
      category = TaskCategory.CellLevel,
      inputSource = InputSource.TestCases(Vector(testCase)),
      evaluation = EvaluationSpec.fileOnly
    )
    AnthropicClientIO.resource("dummy-key-never-used").use { client =>
      XlsxSkill
        .executeCase(
          testCase,
          task,
          SkillContext.empty,
          client,
          AgentConfig(model = "claude-test"),
          EngineConfig.default.copy(outputDir = dir, stream = false)
        )
        .map { result =>
          assertEquals(result.caseNum, 3)
          assertEquals(result.passed, false)
          assert(
            result.error.exists(_.contains("FileUploadFailed")),
            s"error must be recorded: ${result.error}"
          )
          val traceDir = result.tracePath.getOrElse(fail("partial trace was not saved"))
          val traceJson = Files.readString(traceDir.resolve("conversation.json"))
          val parsed = io.circe.parser.parse(traceJson).getOrElse(fail("unparseable trace JSON"))
          val meta = parsed.hcursor.downField("metadata")
          assertEquals(meta.get[Boolean]("passed"), Right(false))
          assert(
            meta.get[String]("error").exists(_.contains("FileUploadFailed")),
            s"trace metadata must carry the error: $traceJson"
          )
        }
    }
  }
