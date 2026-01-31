package com.tjclp.xl.agent.benchmark.grading

import cats.{Order, Show}
import io.circe.*
import io.circe.generic.semiauto.*

// ============================================================================
// Score Hierarchy with Opaque Types
// ============================================================================

/** Base trait for all score types with normalized comparison */
sealed trait Score derives CanEqual:
  /** Normalized value between 0.0 and 1.0 for cross-type comparison */
  def normalized: Double

object Score:

  // --------------------------------------------------------------------------
  // Opaque Type: NormalizedScore
  // --------------------------------------------------------------------------

  /** Zero-overhead normalized score (0.0 to 1.0) */
  opaque type NormalizedScore = Double

  object NormalizedScore:
    inline def apply(value: Double): NormalizedScore =
      math.max(0.0, math.min(1.0, value))

    def unsafe(value: Double): NormalizedScore = value

    extension (s: NormalizedScore)
      inline def value: Double = s
      inline def percent: Int = (s * 100).toInt
      inline def asPercent: String = f"${s * 100}%.1f%%"

    given Encoder[NormalizedScore] = Encoder.encodeDouble.contramap(_.value)
    given Decoder[NormalizedScore] = Decoder.decodeDouble.map(NormalizedScore.apply)
    given Order[NormalizedScore] = Order.by(_.value)
    given Show[NormalizedScore] = Show.show(_.asPercent)

  // --------------------------------------------------------------------------
  // BinaryScore Enum
  // --------------------------------------------------------------------------

  /** Pass/Fail scoring for deterministic evaluation */
  enum BinaryScore extends Score derives CanEqual:
    case Pass
    case Fail

    def normalized: Double = this match
      case Pass => 1.0
      case Fail => 0.0

    def passed: Boolean = this == Pass

  object BinaryScore:
    def fromBoolean(passed: Boolean): BinaryScore =
      if passed then Pass else Fail

    given Encoder[BinaryScore] = Encoder.encodeString.contramap(_.toString.toLowerCase)
    given Decoder[BinaryScore] = Decoder.decodeString.emap {
      case "pass" => Right(Pass)
      case "fail" => Right(Fail)
      case other => Left(s"Unknown BinaryScore: $other")
    }
    given Order[BinaryScore] = Order.by(_.normalized)
    given Show[BinaryScore] = Show.show(_.toString)

  // --------------------------------------------------------------------------
  // FractionalScore Case Class
  // --------------------------------------------------------------------------

  /** Fractional scoring (e.g., 2/3 test cases passed) */
  case class FractionalScore(passing: Int, total: Int) extends Score:
    require(passing >= 0, "passing must be non-negative")
    require(total >= 0, "total must be non-negative")
    require(passing <= total, "passing cannot exceed total")

    def normalized: Double = if total == 0 then 0.0 else passing.toDouble / total
    def allPassed: Boolean = total > 0 && passing == total
    def formatted: String = s"$passing/$total"

  object FractionalScore:
    val zero: FractionalScore = FractionalScore(0, 0)
    val perfect: FractionalScore = FractionalScore(1, 1)

    def fromResults[A](results: List[A])(passed: A => Boolean): FractionalScore =
      FractionalScore(results.count(passed), results.length)

    given Encoder[FractionalScore] = deriveEncoder
    given Decoder[FractionalScore] = deriveDecoder
    given Order[FractionalScore] = Order.by(_.normalized)
    given Show[FractionalScore] = Show.show(_.formatted)

  // --------------------------------------------------------------------------
  // LetterGrade Enum
  // --------------------------------------------------------------------------

  /** Letter grade for LLM-as-judge evaluation */
  enum LetterGrade extends Score derives CanEqual:
    case A, B, C, D, F

    def normalized: Double = this match
      case A => 1.0
      case B => 0.75
      case C => 0.5
      case D => 0.25
      case F => 0.0

    def toNumeric: Int = this match
      case A => 4
      case B => 3
      case C => 2
      case D => 1
      case F => 0

    def isPassing: Boolean = toNumeric >= 2 // C or better

  object LetterGrade:
    def fromNumeric(n: Int): LetterGrade = n match
      case x if x >= 4 => A
      case 3 => B
      case 2 => C
      case 1 => D
      case _ => F

    def fromNormalized(n: Double): LetterGrade =
      if n >= 0.875 then A
      else if n >= 0.625 then B
      else if n >= 0.375 then C
      else if n >= 0.125 then D
      else F

    def average(grades: List[LetterGrade]): Option[LetterGrade] =
      Option.when(grades.nonEmpty) {
        val avg = grades.map(_.toNumeric).sum.toDouble / grades.size
        if avg >= 3.5 then A
        else if avg >= 2.5 then B
        else if avg >= 1.5 then C
        else if avg >= 0.5 then D
        else F
      }

    def fromString(s: String): Either[String, LetterGrade] =
      s.toUpperCase match
        case "A" => Right(A)
        case "B" => Right(B)
        case "C" => Right(C)
        case "D" => Right(D)
        case "F" => Right(F)
        case other => Left(s"Unknown grade: $other")

    given Encoder[LetterGrade] = Encoder.encodeString.contramap(_.toString)
    given Decoder[LetterGrade] = Decoder.decodeString.emap(fromString)
    given Order[LetterGrade] = Order.by(_.toNumeric)
    given Show[LetterGrade] = Show.show(_.toString)

  // --------------------------------------------------------------------------
  // CompositeScore Case Class
  // --------------------------------------------------------------------------

  /** Composite score combining multiple graders */
  case class CompositeScore(
    components: Vector[(String, Score)],
    weights: Option[Map[String, Double]] = None
  ) extends Score:
    def normalized: Double =
      if components.isEmpty then 0.0
      else
        weights match
          case Some(w) =>
            val totalWeight = components.map { case (name, _) => w.getOrElse(name, 1.0) }.sum
            if totalWeight == 0 then 0.0
            else
              components.map { case (name, score) =>
                score.normalized * w.getOrElse(name, 1.0)
              }.sum / totalWeight
          case None =>
            components.map(_._2.normalized).sum / components.size

    def get(name: String): Option[Score] =
      components.find(_._1 == name).map(_._2)

    def passed: Boolean = normalized >= 0.5

  object CompositeScore:
    val empty: CompositeScore = CompositeScore(Vector.empty)

    def fromScores(scores: (String, Score)*): CompositeScore =
      CompositeScore(scores.toVector)

    given Encoder[CompositeScore] = Encoder.instance { cs =>
      Json.obj(
        "components" -> Json.obj(cs.components.map { case (name, score) =>
          name -> Json.fromDouble(score.normalized).getOrElse(Json.Null)
        }*),
        "normalized" -> Json.fromDouble(cs.normalized).getOrElse(Json.Null)
      )
    }
    given Show[CompositeScore] = Show.show { cs =>
      val parts = cs.components.map { case (name, score) =>
        s"$name=${f"${score.normalized}%.2f"}"
      }
      s"CompositeScore(${parts.mkString(", ")} -> ${f"${cs.normalized}%.2f"})"
    }

  // --------------------------------------------------------------------------
  // Type Class Instances for Score
  // --------------------------------------------------------------------------

  given Encoder[Score] = Encoder.instance {
    case b: BinaryScore => BinaryScore.given_Encoder_BinaryScore(b)
    case f: FractionalScore => FractionalScore.given_Encoder_FractionalScore(f)
    case g: LetterGrade => LetterGrade.given_Encoder_LetterGrade(g)
    case c: CompositeScore => CompositeScore.given_Encoder_CompositeScore(c)
  }

  given Order[Score] = Order.by(_.normalized)
  given Show[Score] = Show.show(s => f"${s.normalized}%.2f")
