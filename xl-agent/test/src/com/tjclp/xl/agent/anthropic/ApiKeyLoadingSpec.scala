package com.tjclp.xl.agent.anthropic

import java.nio.file.{Files, Path}
import scala.jdk.CollectionConverters.*
import scala.util.Using

import com.tjclp.xl.agent.error.AgentError
import munit.CatsEffectSuite

class ApiKeyLoadingSpec extends CatsEffectSuite:
  private val temporary = FunFixture[Path](
    setup = _ => Files.createTempDirectory("xl-api-key-spec"),
    teardown = dir =>
      Using.resource(Files.walk(dir)) { paths =>
        paths.iterator().asScala.toList.reverse.foreach(Files.deleteIfExists)
      }
  )

  private def putKey(dir: Path, value: String): Unit =
    Files.createDirectories(dir)
    Files.writeString(dir.resolve(".env"), s"ANTHROPIC_API_KEY=$value\n")
    ()

  temporary.test("GH-838: parent and sibling module credentials are never discovered") { parent =>
    val repo = Files.createDirectory(parent.resolve("repo"))
    putKey(parent, "synthetic-parent")
    putKey(parent.resolve("xl-agent"), "synthetic-sibling")
    AnthropicClientIO.loadApiKey(repo, None).intercept[AgentError.ApiKeyMissing]
  }

  temporary.test("GH-838: removed example directories are not searched") { repo =>
    putKey(repo.resolve("examples/anthropic-sdk"), "synthetic-example")
    putKey(repo.resolve("examples/anthropic-sdk/spreadsheetbench"), "synthetic-benchmark")
    AnthropicClientIO.loadApiKey(repo, None).intercept[AgentError.ApiKeyMissing]
  }

  temporary.test("working directory .env takes precedence over module .env") { repo =>
    putKey(repo, "synthetic-root")
    putKey(repo.resolve("xl-agent"), "synthetic-module")
    AnthropicClientIO.loadApiKey(repo, None).map(key => assertEquals(key, "synthetic-root"))
  }

  temporary.test("module .env is used when the working directory has no key") { repo =>
    Files.writeString(repo.resolve(".env"), "UNRELATED=value\n")
    putKey(repo.resolve("xl-agent"), "synthetic-module")
    AnthropicClientIO.loadApiKey(repo, None).map(key => assertEquals(key, "synthetic-module"))
  }

  temporary.test("environment takes precedence without parsing malformed local files") { repo =>
    Files.writeString(repo.resolve(".env"), "not a dotenv assignment\n")
    AnthropicClientIO
      .loadApiKey(repo, Some("synthetic-environment"))
      .map(key => assertEquals(key, "synthetic-environment"))
  }

  temporary.test("missing local files report ApiKeyMissing") { repo =>
    AnthropicClientIO.loadApiKey(repo, None).intercept[AgentError.ApiKeyMissing]
  }
