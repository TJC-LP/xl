package com.tjclp.xl.agent.benchmark.common

import cats.effect.IO
import cats.syntax.all.*
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.{Files, Path, Paths}
import scala.jdk.OptionConverters.*
import scala.sys.process.*

/** Shared file management utilities for benchmarks */
object FileManager:

  // Default binary/skill versions
  private val DefaultBinaryVersion = "0.9.0"
  private val DefaultBinaryName = s"xl-$DefaultBinaryVersion-linux-amd64"
  private val DefaultSkillName = s"xl-skill-$DefaultBinaryVersion.zip"

  // Default search directories
  private val DefaultSearchDirs = List(
    "../benchmark",
    "examples/anthropic-sdk/benchmark",
    "."
  )

  /** Resolve path to xl binary, optionally downloading from GitHub */
  def resolveBinaryPath(
    pathOverride: Option[Path] = None,
    searchDirs: List[String] = DefaultSearchDirs,
    autoDownload: Boolean = true
  ): IO[Path] =
    pathOverride match
      case Some(p) => IO.pure(p)
      case None =>
        findExistingFile(DefaultBinaryName, searchDirs).flatMap {
          case Some(p) => IO.pure(p)
          case None if autoDownload =>
            for
              targetDir <- IO.pure(searchDirs.lastOption.getOrElse("."))
              _ <- downloadFromGitHub("xl-*-linux-amd64", targetDir)
              path <- findExistingFile(DefaultBinaryName, searchDirs)
                .flatMap(
                  _.liftTo[IO](
                    AgentError.ConfigError(s"Binary not found after download: $DefaultBinaryName")
                  )
                )
            yield path
          case None =>
            IO.raiseError(
              AgentError.ConfigError(
                s"Binary not found: $DefaultBinaryName. Use --xl-binary or place in ${searchDirs.mkString(", ")}"
              )
            )
        }

  /** Resolve path to xl skill zip, optionally downloading from GitHub */
  def resolveSkillPath(
    pathOverride: Option[Path] = None,
    searchDirs: List[String] = DefaultSearchDirs,
    autoDownload: Boolean = true
  ): IO[Path] =
    pathOverride match
      case Some(p) => IO.pure(p)
      case None =>
        findExistingFile(DefaultSkillName, searchDirs).flatMap {
          case Some(p) => IO.pure(p)
          case None if autoDownload =>
            for
              targetDir <- IO.pure(searchDirs.lastOption.getOrElse("."))
              _ <- downloadFromGitHub("xl-skill-*.zip", targetDir)
              path <- findExistingFile(DefaultSkillName, searchDirs)
                .flatMap(
                  _.liftTo[IO](
                    AgentError.ConfigError(s"Skill not found after download: $DefaultSkillName")
                  )
                )
            yield path
          case None =>
            IO.raiseError(
              AgentError.ConfigError(
                s"Skill not found: $DefaultSkillName. Use --xl-skill or place in ${searchDirs.mkString(", ")}"
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

  /** Find a file by pattern (glob) in the given directories */
  def findByPattern(pattern: String, dirs: List[String]): IO[Option[Path]] =
    IO.blocking {
      import java.nio.file.FileSystems
      val matcher = FileSystems.getDefault.getPathMatcher(s"glob:$pattern")

      dirs.view.flatMap { dir =>
        val dirPath = Paths.get(dir)
        if Files.isDirectory(dirPath) then
          Files
            .list(dirPath)
            .filter(p => matcher.matches(p.getFileName))
            .findFirst()
            .map(_.toAbsolutePath)
            .toScala
        else None
      }.headOption
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
