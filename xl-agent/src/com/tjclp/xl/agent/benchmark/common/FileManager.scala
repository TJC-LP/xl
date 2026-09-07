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

  // Release asset patterns. The binary is version-agnostic on purpose: resolution picks the
  // highest version present and auto-download fetches the latest release, so there is no pinned
  // version to keep in sync with releases. The skill zip is then locked to the binary's release
  // (GH-592): a skill from another release teaches a contract the binary does not have.
  private val BinaryPattern = "xl-*-linux-amd64"
  private val SkillPattern = "xl-skill-*.zip"

  /** The skill zip of one release, as `release.yml` names it. */
  def skillPatternFor(version: String): String = s"xl-skill-$version.zip"

  private val VersionInName = """\d+\.\d+\.\d+""".r

  /**
   * The release version embedded in an asset name: `xl-0.20.0-linux-amd64`, `xl-skill-0.20.0.zip`.
   */
  def assetVersion(fileName: String): Option[String] = VersionInName.findFirstIn(fileName)

  private def fileName(path: Path): String = Option(path.getFileName).fold("")(_.toString)

  /**
   * Left when both assets name a release and the releases differ. Nothing to check when either name
   * carries no version (a local build, a hand-made skill): the caller chose them explicitly.
   */
  def lockSkillToBinary(binary: Path, skill: Path): Either[AgentError, Unit] =
    (assetVersion(fileName(binary)), assetVersion(fileName(skill))) match
      case (Some(b), Some(s)) if b != s =>
        Left(
          AgentError.ConfigError(
            s"skill ${fileName(skill)} is release $s but binary ${fileName(binary)} is release $b: " +
              s"the skill must match the binary (use --xl-skill ${skillPatternFor(b)})"
          )
        )
      case _ => Right(())

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

  /**
   * Resolve path to xl skill zip, optionally downloading it from GitHub. With `lockTo`, only the
   * zip of that release qualifies (and the download is that release's, not the latest).
   */
  def resolveSkillPath(
    pathOverride: Option[Path] = None,
    searchDirs: List[String] = DefaultSearchDirs,
    autoDownload: Boolean = true,
    lockTo: Option[String] = None
  ): IO[Path] =
    resolveAsset(
      "Skill",
      lockTo.fold(SkillPattern)(skillPatternFor),
      "--xl-skill",
      pathOverride,
      searchDirs,
      autoDownload,
      tag = lockTo.map(v => s"v$v")
    )

  /**
   * The binary, then the skill zip of the same release (GH-592). Overrides are honoured but still
   * checked against each other: a versioned skill from another release than the binary is refused.
   */
  def resolveReleaseAssets(
    binaryOverride: Option[Path],
    skillOverride: Option[Path],
    searchDirs: List[String] = DefaultSearchDirs,
    autoDownload: Boolean = true
  ): IO[(Path, Path)] =
    for
      binary <- resolveBinaryPath(binaryOverride, searchDirs, autoDownload)
      skill <- resolveSkillPath(
        skillOverride,
        searchDirs,
        autoDownload,
        lockTo = assetVersion(fileName(binary))
      )
      _ <- IO.fromEither(lockSkillToBinary(binary, skill))
    yield (binary, skill)

  private def resolveAsset(
    label: String,
    pattern: String,
    overrideFlag: String,
    pathOverride: Option[Path],
    searchDirs: List[String],
    autoDownload: Boolean,
    tag: Option[String] = None
  ): IO[Path] =
    pathOverride match
      case Some(p) => IO.pure(p)
      case None =>
        findByPattern(pattern, searchDirs).flatMap {
          case Some(p) => IO.pure(p)
          case None if autoDownload =>
            for
              targetDir <- IO.pure(searchDirs.lastOption.getOrElse("."))
              _ <- downloadFromGitHub(pattern, targetDir, tag)
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

  /** Download assets from a GitHub release using the gh CLI: the latest, or the tagged one. */
  def downloadFromGitHub(pattern: String, targetDir: String, tag: Option[String] = None): IO[Unit] =
    IO.blocking {
      val cmd = Seq("gh", "release", "download") ++ tag.toList ++ Seq(
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
          s"Failed to download '$pattern'${tag.fold("")(t => s" of release $t")} from GitHub (exit code: $exitCode)"
        )
    }

  /** Download the latest binary and the skill of the same release from GitHub */
  def downloadReleaseAssets(targetDir: String): IO[(Path, Path)] =
    for
      _ <- IO.println(s"   Downloading xl binary from GitHub...")
      _ <- downloadFromGitHub(BinaryPattern, targetDir)
      binary <- findByPattern(BinaryPattern, List(targetDir))
        .flatMap(
          _.liftTo[IO](AgentError.ConfigError("Binary download succeeded but file not found"))
        )

      version = assetVersion(fileName(binary))
      _ <- IO.println(s"   Downloading xl skill ${version.getOrElse("(latest)")} from GitHub...")
      _ <- downloadFromGitHub(
        version.fold(SkillPattern)(skillPatternFor),
        targetDir,
        version.map(v => s"v$v")
      )
      skill <- findByPattern(version.fold(SkillPattern)(skillPatternFor), List(targetDir))
        .flatMap(
          _.liftTo[IO](AgentError.ConfigError("Skill download succeeded but file not found"))
        )
      _ <- IO.fromEither(lockSkillToBinary(binary, skill))
    yield (binary, skill)

  /** Ensure a directory exists, creating it if necessary */
  def ensureDirectory(path: Path): IO[Unit] =
    IO.blocking(Files.createDirectories(path)).void

  /** Get the filename from a path for use in prompts */
  def getFilename(path: Path): String =
    path.getFileName.toString
