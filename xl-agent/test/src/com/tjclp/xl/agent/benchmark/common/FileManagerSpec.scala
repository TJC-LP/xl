package com.tjclp.xl.agent.benchmark.common

import munit.CatsEffectSuite
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.{Files, Path}

class FileManagerSpec extends CatsEffectSuite:

  private val tempDir = FunFixture[Path](
    setup = _ => Files.createTempDirectory("filemanager-spec"),
    teardown = { dir =>
      import scala.jdk.CollectionConverters.*
      import scala.util.Using
      Using.resource(Files.walk(dir)) { stream =>
        stream.iterator().asScala.toList.reverse.foreach(p => Files.deleteIfExists(p))
      }
    }
  )

  tempDir.test("findByPattern picks the highest version among matches") { dir =>
    Files.createFile(dir.resolve("xl-0.9.0-linux-amd64"))
    Files.createFile(dir.resolve("xl-0.12.2-linux-amd64"))
    Files.createFile(dir.resolve("xl-0.10.1-linux-amd64"))

    FileManager
      .findByPattern("xl-*-linux-amd64", List(dir.toString))
      .map(found => assertEquals(found.map(_.getFileName.toString), Some("xl-0.12.2-linux-amd64")))
  }

  tempDir.test("findByPattern compares versions numerically, not lexicographically") { dir =>
    Files.createFile(dir.resolve("xl-skill-0.9.0.zip"))
    Files.createFile(dir.resolve("xl-skill-0.10.0.zip"))

    FileManager
      .findByPattern("xl-skill-*.zip", List(dir.toString))
      .map(found => assertEquals(found.map(_.getFileName.toString), Some("xl-skill-0.10.0.zip")))
  }

  tempDir.test("findByPattern returns None when nothing matches") { dir =>
    FileManager
      .findByPattern("xl-*-linux-amd64", List(dir.toString))
      .map(found => assertEquals(found, None))
  }

  test("findByPattern tolerates missing directories") {
    FileManager
      .findByPattern("xl-*-linux-amd64", List("/nonexistent/path/for/spec"))
      .map(found => assertEquals(found, None))
  }

  // --- the skill is locked to the binary's release (GH-592) --------------------------------------

  test("assetVersion reads the release version out of an asset file name") {
    assertEquals(FileManager.assetVersion("xl-0.20.0-linux-amd64"), Some("0.20.0"))
    assertEquals(FileManager.assetVersion("xl-skill-0.19.3.zip"), Some("0.19.3"))
    assertEquals(FileManager.assetVersion("xl"), None)
    assertEquals(FileManager.assetVersion("xl-skill.zip"), None)
  }

  test(
    "lockSkillToBinary: same release passes, different releases fail, no version has nothing to check"
  ) {
    val binary = Path.of("xl-0.20.0-linux-amd64")
    assertEquals(FileManager.lockSkillToBinary(binary, Path.of("xl-skill-0.20.0.zip")), Right(()))
    FileManager.lockSkillToBinary(binary, Path.of("xl-skill-0.19.0.zip")) match
      case Left(AgentError.ConfigError(msg)) =>
        assert(msg.contains("0.20.0") && msg.contains("0.19.0"), msg)
      case other => fail(s"expected ConfigError, got $other")
    assertEquals(
      FileManager.lockSkillToBinary(Path.of("xl"), Path.of("xl-skill-0.20.0.zip")),
      Right(())
    )
    assertEquals(FileManager.lockSkillToBinary(binary, Path.of("xl-skill.zip")), Right(()))
  }

  tempDir.test("resolveSkillPath locked to a version picks that release, not the highest") { dir =>
    Files.createFile(dir.resolve("xl-skill-0.19.0.zip"))
    Files.createFile(dir.resolve("xl-skill-0.20.0.zip"))
    Files.createFile(dir.resolve("xl-skill-0.21.0.zip"))

    FileManager
      .resolveSkillPath(None, List(dir.toString), autoDownload = false, lockTo = Some("0.20.0"))
      .map(p => assertEquals(p.getFileName.toString, "xl-skill-0.20.0.zip"))
  }

  tempDir.test("resolveSkillPath locked to an absent release fails naming the zip it wanted") {
    dir =>
      Files.createFile(dir.resolve("xl-skill-0.21.0.zip"))

      FileManager
        .resolveSkillPath(None, List(dir.toString), autoDownload = false, lockTo = Some("0.20.0"))
        .attempt
        .map {
          case Left(AgentError.ConfigError(msg)) => assert(msg.contains("xl-skill-0.20.0.zip"), msg)
          case other => fail(s"expected ConfigError, got $other")
        }
  }

  tempDir.test("resolveReleaseAssets pairs the newest binary with the skill of the same release") {
    dir =>
      Files.createFile(dir.resolve("xl-0.19.0-linux-amd64"))
      Files.createFile(dir.resolve("xl-0.20.0-linux-amd64"))
      Files.createFile(dir.resolve("xl-skill-0.19.0.zip"))
      Files.createFile(dir.resolve("xl-skill-0.20.0.zip"))
      Files.createFile(dir.resolve("xl-skill-0.21.0.zip"))

      FileManager
        .resolveReleaseAssets(None, None, List(dir.toString), autoDownload = false)
        .map { (binary, skill) =>
          assertEquals(binary.getFileName.toString, "xl-0.20.0-linux-amd64")
          assertEquals(skill.getFileName.toString, "xl-skill-0.20.0.zip")
        }
  }

  tempDir.test(
    "resolveReleaseAssets refuses a skill override from another release than the binary"
  ) { dir =>
    val binary = Files.createFile(dir.resolve("xl-0.20.0-linux-amd64"))
    val skill = Files.createFile(dir.resolve("xl-skill-0.19.0.zip"))

    FileManager
      .resolveReleaseAssets(Some(binary), Some(skill), List(dir.toString), autoDownload = false)
      .attempt
      .map {
        case Left(AgentError.ConfigError(msg)) =>
          assert(msg.contains("0.20.0") && msg.contains("0.19.0"), msg)
        case other => fail(s"expected ConfigError, got $other")
      }
  }

  tempDir.test(
    "resolveReleaseAssets with an unversioned binary override has no lock: highest skill"
  ) { dir =>
    val binary = Files.createFile(dir.resolve("xl"))
    Files.createFile(dir.resolve("xl-skill-0.19.0.zip"))
    Files.createFile(dir.resolve("xl-skill-0.20.0.zip"))

    FileManager
      .resolveReleaseAssets(Some(binary), None, List(dir.toString), autoDownload = false)
      .map { (b, skill) =>
        assertEquals(b, binary)
        assertEquals(skill.getFileName.toString, "xl-skill-0.20.0.zip")
      }
  }
