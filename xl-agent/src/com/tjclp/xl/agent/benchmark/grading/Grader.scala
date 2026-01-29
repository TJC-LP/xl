package com.tjclp.xl.agent.benchmark.grading

import cats.effect.IO
import cats.syntax.all.*
import io.circe.*
import io.circe.generic.semiauto.*

import java.nio.file.Path

// ============================================================================
// Grader Context
// ============================================================================

/** Context provided to graders for evaluation */
trait GraderContext:
  /** Unique task identifier */
  def taskId: String

  /** Skill name used for the task */
  def skill: String

  /** Test case number (for SpreadsheetBench-style tasks) */
  def caseNum: Option[Int]

  /** Path to the agent's output file (if produced) */
  def outputPath: Option[Path]

  /** Path to the expected answer file */
  def answerPath: Option[Path]

  /** Cell range(s) containing the answer */
  def answerPosition: Option[String]

  /** Raw text response from the agent */
  def responseText: Option[String]

  /** Expected answer for LLM grading */
  def expectedAnswer: Option[String]

object GraderContext:

  /** Create a context for file-based grading (SpreadsheetBench) */
  def forFile(
    taskId: String,
    skill: String,
    caseNum: Int,
    outputPath: Path,
    answerPath: Path,
    answerPosition: String
  ): GraderContext = FileGraderContext(
    taskId = taskId,
    skill = skill,
    caseNum = Some(caseNum),
    outputPath = Some(outputPath),
    answerPath = Some(answerPath),
    answerPosition = Some(answerPosition),
    responseText = None,
    expectedAnswer = None
  )

  /** Create a context for LLM-based grading (TokenBenchmark) */
  def forLLM(
    taskId: String,
    skill: String,
    responseText: String,
    expectedAnswer: String
  ): GraderContext = LLMGraderContext(
    taskId = taskId,
    skill = skill,
    caseNum = None,
    outputPath = None,
    answerPath = None,
    answerPosition = None,
    responseText = Some(responseText),
    expectedAnswer = Some(expectedAnswer)
  )

  /** Create a composite context for both grader types */
  def forComposite(
    taskId: String,
    skill: String,
    caseNum: Int,
    outputPath: Path,
    answerPath: Path,
    answerPosition: String,
    responseText: String,
    expectedAnswer: String
  ): GraderContext = CompositeGraderContext(
    taskId = taskId,
    skill = skill,
    caseNum = Some(caseNum),
    outputPath = Some(outputPath),
    answerPath = Some(answerPath),
    answerPosition = Some(answerPosition),
    responseText = Some(responseText),
    expectedAnswer = Some(expectedAnswer)
  )

/** Context implementation for file-based grading */
private case class FileGraderContext(
  taskId: String,
  skill: String,
  caseNum: Option[Int],
  outputPath: Option[Path],
  answerPath: Option[Path],
  answerPosition: Option[String],
  responseText: Option[String],
  expectedAnswer: Option[String]
) extends GraderContext

/** Context implementation for LLM grading */
private case class LLMGraderContext(
  taskId: String,
  skill: String,
  caseNum: Option[Int],
  outputPath: Option[Path],
  answerPath: Option[Path],
  answerPosition: Option[String],
  responseText: Option[String],
  expectedAnswer: Option[String]
) extends GraderContext

/** Context implementation for composite grading */
private case class CompositeGraderContext(
  taskId: String,
  skill: String,
  caseNum: Option[Int],
  outputPath: Option[Path],
  answerPath: Option[Path],
  answerPosition: Option[String],
  responseText: Option[String],
  expectedAnswer: Option[String]
) extends GraderContext

// ============================================================================
// Grade Result
// ============================================================================

/** Result of a grading operation */
case class GradeResult[+S <: Score](
  graderName: String,
  score: S,
  details: GradeDetails,
  metadata: Map[String, String] = Map.empty
):
  def passed: Boolean = details.passed
  def withMetadata(key: String, value: String): GradeResult[S] =
    copy(metadata = metadata + (key -> value))

object GradeResult:
  given [S <: Score: Encoder]: Encoder[GradeResult[S]] = deriveEncoder

/** Details about the grading decision */
case class GradeDetails(
  passed: Boolean,
  reasoning: Option[String] = None,
  mismatches: List[CellMismatch] = Nil
):
  def withReasoning(reason: String): GradeDetails =
    copy(reasoning = Some(reason))
  def withMismatches(ms: List[CellMismatch]): GradeDetails =
    copy(mismatches = ms)

object GradeDetails:
  val passed: GradeDetails = GradeDetails(passed = true)
  val failed: GradeDetails = GradeDetails(passed = false)

  def pass(reasoning: String): GradeDetails =
    GradeDetails(passed = true, reasoning = Some(reasoning))

  def fail(reasoning: String): GradeDetails =
    GradeDetails(passed = false, reasoning = Some(reasoning))

  def fail(reasoning: String, mismatches: List[CellMismatch]): GradeDetails =
    GradeDetails(passed = false, reasoning = Some(reasoning), mismatches = mismatches)

  given Encoder[GradeDetails] = deriveEncoder

/** A cell value mismatch in file comparison */
case class CellMismatch(
  ref: String,
  expected: String,
  actual: String
)

object CellMismatch:
  given Encoder[CellMismatch] = deriveEncoder

// ============================================================================
// Grader Trait
// ============================================================================

/** Base trait for all graders */
trait Grader[S <: Score]:
  /** Unique name for this grader */
  def name: String

  /** Grade a task given its context */
  def grade(ctx: GraderContext): IO[GradeResult[S]]

  /** Check if this grader can handle the given context */
  def canHandle(ctx: GraderContext): Boolean

object Grader:

  /** Create a grader from a function */
  def apply[S <: Score](
    graderName: String,
    canHandleFn: GraderContext => Boolean = _ => true
  )(gradeFn: GraderContext => IO[GradeResult[S]]): Grader[S] =
    new Grader[S]:
      def name: String = graderName
      def grade(ctx: GraderContext): IO[GradeResult[S]] = gradeFn(ctx)
      def canHandle(ctx: GraderContext): Boolean = canHandleFn(ctx)

  /** Combine multiple graders into a composite grader */
  def composite(graders: Vector[Grader[Score]]): CompositeGrader =
    new CompositeGrader(graders)

// ============================================================================
// Composite Grader
// ============================================================================

/** A grader that combines multiple graders and produces a composite score */
final class CompositeGrader(
  graders: Vector[Grader[Score]],
  weights: Option[Map[String, Double]] = None
) extends Grader[Score.CompositeScore]:

  def name: String = s"composite(${graders.map(_.name).mkString(",")})"

  def canHandle(ctx: GraderContext): Boolean =
    graders.exists(_.canHandle(ctx))

  def grade(ctx: GraderContext): IO[GradeResult[Score.CompositeScore]] =
    graders
      .filter(_.canHandle(ctx))
      .traverse(_.grade(ctx))
      .map { results =>
        val components = results.map(r => r.graderName -> r.score)
        val composite = Score.CompositeScore(components, weights)
        val allPassed = results.forall(_.details.passed)
        val combinedReasons = results.flatMap(_.details.reasoning).mkString("; ")
        val combinedMismatches = results.flatMap(_.details.mismatches).toList
        val combinedMetadata = results.flatMap(_.metadata).toMap

        GradeResult(
          graderName = name,
          score = composite,
          details = GradeDetails(
            passed = allPassed,
            reasoning = Option.when(combinedReasons.nonEmpty)(combinedReasons),
            mismatches = combinedMismatches
          ),
          metadata = combinedMetadata
        )
      }

  /** Add weights for combining scores */
  def withWeights(w: Map[String, Double]): CompositeGrader =
    new CompositeGrader(graders, Some(w))

  /** Add a grader to this composite */
  def add(grader: Grader[Score]): CompositeGrader =
    new CompositeGrader(graders :+ grader, weights)

object CompositeGrader:
  def fromGraders(graders: Grader[Score]*): CompositeGrader =
    new CompositeGrader(graders.toVector)

// ============================================================================
// Grader Type Enum
// ============================================================================

/** Supported grader types */
enum GraderType derives CanEqual:
  case File
  case LLM
  case FileAndLLM
  case Custom(name: String)

object GraderType:
  def fromString(s: String): Either[String, GraderType] =
    s.toLowerCase match
      case "file" => Right(File)
      case "llm" => Right(LLM)
      case "file,llm" | "fileandllm" | "both" => Right(FileAndLLM)
      case other => Right(Custom(other))

  given Encoder[GraderType] = Encoder.encodeString.contramap {
    case File => "file"
    case LLM => "llm"
    case FileAndLLM => "file,llm"
    case Custom(name) => name
  }

  given Decoder[GraderType] = Decoder.decodeString.emap(fromString)
