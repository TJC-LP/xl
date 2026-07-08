package com.tjclp.xl.agent.benchmark.common

import cats.effect.IO
import cats.syntax.all.*
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.{Files, Path, Paths}
import scala.jdk.CollectionConverters.*
import scala.sys.process.*
import scala.util.Using

/** Shared file management utilities for benchmarks */
object FileManager:

  // Release asset patterns. Version-agnostic on purpose: resolution picks the
  // highest version present, and auto-download always fetches the latest
  // release, so there is no pinned version to keep in sync with releases.
  private val BinaryPattern = "xl-*-linux-amd64"
  private val SkillPattern = "xl-skill-*.zip"

  // Default search directories
  private val DefaultSearchDirs = List(
    "../benchmark",
    "examples/anthropic-sdk/benchmark",
    "."
  )

  /** Resolve path to xl binary, optionally downloading the latest release from GitHub */
  def resolveBinaryPath(
    pathOverride: Option[Path] = None,
    searchDirs: List[String] = DefaultSearchDirs,
    autoDownload: Boolean = true
  ): IO[Path] =
    resolveAsset("Binary", BinaryPattern, "--xl-binary", pathOverride, searchDirs, autoDownload)

  /** Resolve path to xl skill zip, optionally downloading the latest release from GitHub */
  def resolveSkillPath(
    pathOverride: Option[Path] = None,
    searchDirs: List[String] = DefaultSearchDirs,
    autoDownload: Boolean = true
  ): IO[Path] =
    resolveAsset("Skill", SkillPattern, "--xl-skill", pathOverride, searchDirs, autoDownload)

  private def resolveAsset(
    label: String,
    pattern: String,
    overrideFlag: String,
    pathOverride: Option[Path],
    searchDirs: List[String],
    autoDownload: Boolean
  ): IO[Path] =
    pathOverride match
      case Some(p) => IO.pure(p)
      case None =>
        findByPattern(pattern, searchDirs).flatMap {
          case Some(p) => IO.pure(p)
          case None if autoDownload =>
            for
              targetDir <- IO.pure(searchDirs.lastOption.getOrElse("."))
              _ <- downloadFromGitHub(pattern, targetDir)
              path <- findByPattern(pattern, searchDirs)
                .flatMap(
                  _.liftTo[IO](
                    AgentError.ConfigError(s"$label not found after download: $pattern")
                  )
                )
            yield path
          case None =>
            IO.raiseError(
              AgentError.ConfigError(
                s"$label not found: $pattern. Use $overrideFlag or place in ${searchDirs.mkString(", ")}"
              )
            )
        }

  /** Find an existing file in the given directories */
  def findExistingFile(name: String, dirs: List[String]): IO[Option[Path]] =
    IO.blocking {
      dirs.view
        .map(d => Paths.get(d, name))
        .find(p => Files.exists(p))
    }

  /**
   * Find a file by pattern (glob) in the given directories.
   *
   * When several files match (e.g. binaries from multiple releases side by side), the one with the
   * highest embedded semantic version wins.
   */
  def findByPattern(pattern: String, dirs: List[String]): IO[Option[Path]] =
    IO.blocking {
      import java.nio.file.FileSystems
      val matcher = FileSystems.getDefault.getPathMatcher(s"glob:$pattern")
      val versionRegex = """(\d+)\.(\d+)\.(\d+)""".r

      def versionKey(p: Path): (Int, Int, Int) =
        versionRegex.findFirstMatchIn(p.getFileName.toString) match
          case Some(m) => (m.group(1).toInt, m.group(2).toInt, m.group(3).toInt)
          case None => (-1, -1, -1)

      dirs
        .flatMap { dir =>
          val dirPath = Paths.get(dir)
          if Files.isDirectory(dirPath) then
            Using.resource(Files.list(dirPath)) { stream =>
              stream
                .iterator()
                .asScala
                .filter(p => matcher.matches(p.getFileName))
                .map(_.toAbsolutePath)
                .toList
            }
          else Nil
        }
        .maxByOption(versionKey)
    }

  /** Download assets from GitHub release using gh CLI */
  def downloadFromGitHub(pattern: String, targetDir: String): IO[Unit] =
    IO.blocking {
      val cmd = Seq(
        "gh",
        "release",
        "download",
        "--repo",
        "TJC-LP/xl",
        "--pattern",
        pattern,
        "-D",
        targetDir
      )

      val exitCode = cmd.!
      if exitCode != 0 then
        throw AgentError.ConfigError(
          s"Failed to download '$pattern' from GitHub (exit code: $exitCode)"
        )
    }

  /** Download both binary and skill from GitHub release */
  def downloadReleaseAssets(targetDir: String): IO[(Path, Path)] =
    for
      _ <- IO.println(s"   Downloading xl binary from GitHub...")
      _ <- downloadFromGitHub("xl-*-linux-amd64", targetDir)

      _ <- IO.println(s"   Downloading xl skill from GitHub...")
      _ <- downloadFromGitHub("xl-skill-*.zip", targetDir)

      binary <- findByPattern("xl-*-linux-amd64", List(targetDir))
        .flatMap(
          _.liftTo[IO](AgentError.ConfigError("Binary download succeeded but file not found"))
        )

      skill <- findByPattern("xl-skill-*.zip", List(targetDir))
        .flatMap(
          _.liftTo[IO](AgentError.ConfigError("Skill download succeeded but file not found"))
        )
    yield (binary, skill)

  /** Ensure a directory exists, creating it if necessary */
  def ensureDirectory(path: Path): IO[Unit] =
    IO.blocking(Files.createDirectories(path)).void

  /** Get the filename from a path for use in prompts */
  def getFilename(path: Path): String =
    path.getFileName.toString
