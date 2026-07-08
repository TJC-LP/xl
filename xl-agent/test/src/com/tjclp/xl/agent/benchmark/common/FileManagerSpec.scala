package com.tjclp.xl.agent.benchmark.common

import munit.CatsEffectSuite

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
