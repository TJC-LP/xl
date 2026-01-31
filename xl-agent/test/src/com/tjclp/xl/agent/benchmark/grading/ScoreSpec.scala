package com.tjclp.xl.agent.benchmark.grading

import munit.FunSuite
import Score.*

class ScoreSpec extends FunSuite:

  // --------------------------------------------------------------------------
  // NormalizedScore Tests
  // --------------------------------------------------------------------------

  test("NormalizedScore clamps values to [0, 1]"):
    assertEquals(NormalizedScore(-0.5).value, 0.0)
    assertEquals(NormalizedScore(0.5).value, 0.5)
    assertEquals(NormalizedScore(1.5).value, 1.0)

  test("NormalizedScore.percent returns percentage"):
    assertEquals(NormalizedScore(0.75).percent, 75)
    assertEquals(NormalizedScore(0.333).percent, 33)
    assertEquals(NormalizedScore(1.0).percent, 100)

  test("NormalizedScore.asPercent returns percentage string"):
    assertEquals(NormalizedScore(0.75).asPercent, "75.0%")

  // --------------------------------------------------------------------------
  // BinaryScore Tests
  // --------------------------------------------------------------------------

  test("BinaryScore.Pass normalizes to 1.0"):
    assertEquals(BinaryScore.Pass.normalized, 1.0)

  test("BinaryScore.Fail normalizes to 0.0"):
    assertEquals(BinaryScore.Fail.normalized, 0.0)

  test("BinaryScore.fromBoolean creates correct score"):
    assertEquals(BinaryScore.fromBoolean(true), BinaryScore.Pass)
    assertEquals(BinaryScore.fromBoolean(false), BinaryScore.Fail)

  test("BinaryScore.passed returns correct boolean"):
    assert(BinaryScore.Pass.passed)
    assert(!BinaryScore.Fail.passed)

  // --------------------------------------------------------------------------
  // FractionalScore Tests
  // --------------------------------------------------------------------------

  test("FractionalScore normalizes correctly"):
    assertEquals(FractionalScore(2, 4).normalized, 0.5)
    assertEquals(FractionalScore(3, 3).normalized, 1.0)
    assertEquals(FractionalScore(0, 5).normalized, 0.0)
    assertEquals(FractionalScore(0, 0).normalized, 0.0)

  test("FractionalScore.allPassed checks if all cases passed"):
    assert(FractionalScore(3, 3).allPassed)
    assert(!FractionalScore(2, 3).allPassed)
    assert(!FractionalScore(0, 0).allPassed)

  test("FractionalScore.fromResults creates score from list"):
    val results = List(true, true, false, true)
    val score = FractionalScore.fromResults(results)(identity)
    assertEquals(score.passing, 3)
    assertEquals(score.total, 4)

  // --------------------------------------------------------------------------
  // LetterGrade Tests
  // --------------------------------------------------------------------------

  test("LetterGrade normalizes correctly"):
    assertEquals(LetterGrade.A.normalized, 1.0)
    assertEquals(LetterGrade.B.normalized, 0.75)
    assertEquals(LetterGrade.C.normalized, 0.5)
    assertEquals(LetterGrade.D.normalized, 0.25)
    assertEquals(LetterGrade.F.normalized, 0.0)

  test("LetterGrade.toNumeric returns correct values"):
    assertEquals(LetterGrade.A.toNumeric, 4)
    assertEquals(LetterGrade.B.toNumeric, 3)
    assertEquals(LetterGrade.C.toNumeric, 2)
    assertEquals(LetterGrade.D.toNumeric, 1)
    assertEquals(LetterGrade.F.toNumeric, 0)

  test("LetterGrade.isPassing identifies passing grades"):
    assert(LetterGrade.A.isPassing)
    assert(LetterGrade.B.isPassing)
    assert(LetterGrade.C.isPassing)
    assert(!LetterGrade.D.isPassing)
    assert(!LetterGrade.F.isPassing)

  test("LetterGrade.fromNumeric converts correctly"):
    assertEquals(LetterGrade.fromNumeric(4), LetterGrade.A)
    assertEquals(LetterGrade.fromNumeric(3), LetterGrade.B)
    assertEquals(LetterGrade.fromNumeric(2), LetterGrade.C)
    assertEquals(LetterGrade.fromNumeric(1), LetterGrade.D)
    assertEquals(LetterGrade.fromNumeric(0), LetterGrade.F)
    assertEquals(LetterGrade.fromNumeric(5), LetterGrade.A)
    assertEquals(LetterGrade.fromNumeric(-1), LetterGrade.F)

  test("LetterGrade.fromString parses correctly"):
    assertEquals(LetterGrade.fromString("A"), Right(LetterGrade.A))
    assertEquals(LetterGrade.fromString("b"), Right(LetterGrade.B))
    assertEquals(LetterGrade.fromString("F"), Right(LetterGrade.F))
    assert(LetterGrade.fromString("X").isLeft)

  test("LetterGrade.average computes correct average"):
    assertEquals(LetterGrade.average(List(LetterGrade.A, LetterGrade.A)), Some(LetterGrade.A))
    assertEquals(LetterGrade.average(List(LetterGrade.A, LetterGrade.B)), Some(LetterGrade.A)) // 3.5 -> A
    assertEquals(LetterGrade.average(List(LetterGrade.B, LetterGrade.C)), Some(LetterGrade.B)) // 2.5 -> B
    assertEquals(LetterGrade.average(List(LetterGrade.F, LetterGrade.F)), Some(LetterGrade.F))
    assertEquals(LetterGrade.average(Nil), None)

  // --------------------------------------------------------------------------
  // CompositeScore Tests
  // --------------------------------------------------------------------------

  test("CompositeScore averages component scores"):
    val composite = CompositeScore.fromScores(
      "file" -> BinaryScore.Pass,
      "llm" -> LetterGrade.C
    )
    assertEquals(composite.normalized, 0.75) // (1.0 + 0.5) / 2

  test("CompositeScore with weights"):
    val composite = CompositeScore(
      Vector("file" -> BinaryScore.Pass, "llm" -> LetterGrade.F),
      Some(Map("file" -> 2.0, "llm" -> 1.0))
    )
    // (1.0 * 2.0 + 0.0 * 1.0) / 3.0 = 0.667
    assert(math.abs(composite.normalized - 0.667) < 0.01)

  test("CompositeScore.get retrieves component"):
    val composite = CompositeScore.fromScores(
      "file" -> BinaryScore.Pass,
      "llm" -> LetterGrade.B
    )
    assertEquals(composite.get("file"), Some(BinaryScore.Pass))
    assertEquals(composite.get("llm"), Some(LetterGrade.B))
    assertEquals(composite.get("unknown"), None)

  test("CompositeScore.passed checks threshold"):
    val passing = CompositeScore.fromScores("a" -> BinaryScore.Pass)
    val failing = CompositeScore.fromScores("a" -> BinaryScore.Fail)
    val borderline = CompositeScore.fromScores("a" -> LetterGrade.C)

    assert(passing.passed)
    assert(!failing.passed)
    assert(borderline.passed) // 0.5 >= 0.5
