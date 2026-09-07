package com.tjclp.xl.text

import munit.FunSuite

class SuggestSpec extends FunSuite:

  test("closest ranks the near miss and drops the unrelated name") {
    assertEquals(Suggest.closest("Sale", Seq("Sales", "Data")), Vector("Sales"))
  }

  test("closest ranks by distance, then by candidate for a stable order") {
    assertEquals(
      Suggest.closest("Data", Seq("Date", "Dat", "Data2", "Summary")),
      Vector("Dat", "Data2", "Date")
    )
  }

  test("closest is empty when nothing is close") {
    assertEquals(Suggest.closest("Nope", Seq("Data", "Summary")), Vector.empty[String])
    assertEquals(Suggest.closest("Data", Nil), Vector.empty[String])
  }

  test("closest is case-insensitive but returns the candidate as spelled") {
    assertEquals(Suggest.closest("data", Seq("DATA", "Summary")), Vector("DATA"))
    assertEquals(Suggest.closest("SUMARY", Seq("Data", "Summary")), Vector("Summary"))
  }

  test("closest keeps distance <= max(2, input.length / 3) for inputs longer than three") {
    // four characters: two edits allowed, three refused
    assertEquals(Suggest.closest("abcd", Seq("abcdef", "abcdefg")), Vector("abcdef"))
    // long input (15 chars): length / 3 = 5 edits allowed, six refused
    assertEquals(Suggest.closest("IncomeStatement", Seq("IncomeStatementFY2025")), Vector.empty)
    assertEquals(
      Suggest.closest("IncomeStatement", Seq("IncomeStatement2025", "Income")),
      Vector("IncomeStatement2025")
    )
    assertEquals(
      Suggest.closest("IncomeStatement", Seq("Income Statemnt")),
      Vector("Income Statemnt")
    )
  }

  test("closest allows one edit for an input of three characters or fewer") {
    // `pu` is one edit from put; cf and putf (two edits each) are not near misses
    assertEquals(Suggest.closest("pu", Seq("put", "cf", "putf", "style")), Vector("put"))
    assertEquals(Suggest.closest("Dat", Seq("Data", "Date", "Summary")), Vector("Data", "Date"))
    assertEquals(Suggest.closest("ab", Seq("abcd", "abcde")), Vector.empty[String])
  }

  test("closest never accepts a distance equal to the input length (a full rewrite)") {
    // two-character input: `cf` is two substitutions away from `pu`
    assertEquals(Suggest.closest("pu", Seq("cf")), Vector.empty[String])
    // one character: only an exact (case-insensitive) match remains
    assertEquals(Suggest.closest("a", Seq("A", "b", "ab")), Vector("A"))
    // empty input suggests nothing
    assertEquals(Suggest.closest("", Seq("a", "put")), Vector.empty[String])
    assertEquals(Suggest.threshold("pu"), 1)
    assertEquals(Suggest.threshold("a"), 0)
    assertEquals(Suggest.threshold("abcd"), 2)
    assertEquals(Suggest.threshold("IncomeStatement"), 5)
  }

  test("closest honours the limit and defaults to three") {
    val many = Seq("Dat1", "Dat2", "Dat3", "Dat4", "Dat5")
    assertEquals(Suggest.closest("Data", many).size, 3)
    assertEquals(Suggest.closest("Data", many, limit = 5), many.toVector)
    assertEquals(Suggest.closest("Data", many, limit = 0), Vector.empty[String])
  }

  test("closest drops duplicate candidates") {
    assertEquals(Suggest.closest("Data", Seq("Date", "Date")), Vector("Date"))
  }
