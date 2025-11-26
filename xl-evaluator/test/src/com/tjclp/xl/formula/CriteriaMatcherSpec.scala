package com.tjclp.xl.formula

import munit.FunSuite
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.CriteriaMatcher.*

class CriteriaMatcherSpec extends FunSuite:

  // ===== Criterion Parsing Tests =====

  test("parse: exact string") {
    assertEquals(parse("Apple"), Exact("Apple"))
  }

  test("parse: exact number from BigDecimal") {
    assertEquals(parse(BigDecimal(42)), Exact(BigDecimal(42)))
  }

  test("parse: exact number from Int") {
    assertEquals(parse(100), Exact(BigDecimal(100)))
  }

  test("parse: exact boolean") {
    assertEquals(parse(true), Exact(true))
    assertEquals(parse(false), Exact(false))
  }

  test("parse: greater than") {
    assertEquals(parse(">100"), Compare(CompareOp.Gt, BigDecimal(100)))
  }

  test("parse: greater than or equal") {
    assertEquals(parse(">=50"), Compare(CompareOp.Gte, BigDecimal(50)))
  }

  test("parse: less than") {
    assertEquals(parse("<10"), Compare(CompareOp.Lt, BigDecimal(10)))
  }

  test("parse: less than or equal") {
    assertEquals(parse("<=5"), Compare(CompareOp.Lte, BigDecimal(5)))
  }

  test("parse: not equal") {
    assertEquals(parse("<>0"), Compare(CompareOp.Neq, BigDecimal(0)))
  }

  test("parse: negative number comparison") {
    assertEquals(parse(">-10"), Compare(CompareOp.Gt, BigDecimal(-10)))
    assertEquals(parse("<=-5.5"), Compare(CompareOp.Lte, BigDecimal("-5.5")))
  }

  test("parse: decimal comparison") {
    assertEquals(parse(">3.14"), Compare(CompareOp.Gt, BigDecimal("3.14")))
  }

  test("parse: comparison with spaces") {
    assertEquals(parse("> 100"), Compare(CompareOp.Gt, BigDecimal(100)))
    assertEquals(parse("<= 50"), Compare(CompareOp.Lte, BigDecimal(50)))
  }

  test("parse: invalid comparison falls back to exact") {
    assertEquals(parse(">abc"), Exact(">abc"))
    assertEquals(parse("<>xyz"), Exact("<>xyz"))
  }

  test("parse: wildcard with asterisk") {
    assertEquals(parse("A*"), Wildcard("A*"))
    assertEquals(parse("*pple"), Wildcard("*pple"))
    assertEquals(parse("*"), Wildcard("*"))
  }

  test("parse: wildcard with question mark") {
    assertEquals(parse("A?ple"), Wildcard("A?ple"))
    assertEquals(parse("???"), Wildcard("???"))
  }

  test("parse: escaped wildcards are not wildcards") {
    // Escaped wildcards become exact match with the literal character (unescaped)
    assertEquals(parse("~*"), Exact("*"))
    assertEquals(parse("~?"), Exact("?"))
    assertEquals(parse("test~*"), Exact("test*"))
  }

  test("parse: mixed escaped and unescaped wildcards") {
    assertEquals(parse("~**"), Wildcard("~**"))
    assertEquals(parse("*~*"), Wildcard("*~*"))
  }

  test("parse: equals prefix extracts number") {
    assertEquals(parse("=100"), Exact(BigDecimal(100)))
    assertEquals(parse("=text"), Exact("text"))
  }

  // ===== Exact Matching Tests =====

  test("matches: exact text case-insensitive") {
    val cell = CellValue.Text("Apple")
    assert(matches(cell, Exact("Apple")))
    assert(matches(cell, Exact("APPLE")))
    assert(matches(cell, Exact("apple")))
  }

  test("matches: exact text no match") {
    val cell = CellValue.Text("Apple")
    assert(!matches(cell, Exact("Banana")))
  }

  test("matches: exact number") {
    val cell = CellValue.Number(BigDecimal(42))
    assert(matches(cell, Exact(BigDecimal(42))))
    assert(!matches(cell, Exact(BigDecimal(43))))
  }

  test("matches: exact number from text criterion") {
    val cell = CellValue.Number(BigDecimal(42))
    assert(matches(cell, Exact("42")))
    assert(!matches(cell, Exact("43")))
  }

  test("matches: exact text from number criterion") {
    val cell = CellValue.Text("42")
    assert(matches(cell, Exact(BigDecimal(42))))
  }

  test("matches: exact boolean") {
    val trueCell = CellValue.Bool(true)
    val falseCell = CellValue.Bool(false)
    assert(matches(trueCell, Exact(true)))
    assert(matches(falseCell, Exact(false)))
    assert(!matches(trueCell, Exact(false)))
  }

  test("matches: exact boolean from text criterion") {
    val cell = CellValue.Bool(true)
    assert(matches(cell, Exact("TRUE")))
    assert(matches(cell, Exact("true")))
    assert(!matches(cell, Exact("FALSE")))
  }

  test("matches: empty cell matches empty string") {
    assert(matches(CellValue.Empty, Exact("")))
    assert(!matches(CellValue.Empty, Exact("something")))
  }

  test("matches: formula with cached value") {
    val cell = CellValue.Formula("=A1+B1", Some(CellValue.Number(BigDecimal(100))))
    assert(matches(cell, Exact(BigDecimal(100))))
    assert(!matches(cell, Exact(BigDecimal(50))))
  }

  test("matches: formula without cached value") {
    val cell = CellValue.Formula("=A1+B1", None)
    assert(!matches(cell, Exact(BigDecimal(100))))
  }

  test("matches: rich text") {
    val cell = CellValue.RichText(com.tjclp.xl.richtext.RichText.plain("Hello"))
    assert(matches(cell, Exact("Hello")))
    assert(matches(cell, Exact("HELLO")))
  }

  test("matches: error cells never match") {
    val cell = CellValue.Error(com.tjclp.xl.cells.CellError.Value)
    assert(!matches(cell, Exact("anything")))
  }

  // ===== Comparison Matching Tests =====

  test("matches: greater than") {
    val cell = CellValue.Number(BigDecimal(150))
    assert(matches(cell, Compare(CompareOp.Gt, BigDecimal(100))))
    assert(!matches(cell, Compare(CompareOp.Gt, BigDecimal(150))))
    assert(!matches(cell, Compare(CompareOp.Gt, BigDecimal(200))))
  }

  test("matches: greater than or equal") {
    val cell = CellValue.Number(BigDecimal(100))
    assert(matches(cell, Compare(CompareOp.Gte, BigDecimal(100))))
    assert(matches(cell, Compare(CompareOp.Gte, BigDecimal(50))))
    assert(!matches(cell, Compare(CompareOp.Gte, BigDecimal(150))))
  }

  test("matches: less than") {
    val cell = CellValue.Number(BigDecimal(50))
    assert(matches(cell, Compare(CompareOp.Lt, BigDecimal(100))))
    assert(!matches(cell, Compare(CompareOp.Lt, BigDecimal(50))))
    assert(!matches(cell, Compare(CompareOp.Lt, BigDecimal(25))))
  }

  test("matches: less than or equal") {
    val cell = CellValue.Number(BigDecimal(50))
    assert(matches(cell, Compare(CompareOp.Lte, BigDecimal(50))))
    assert(matches(cell, Compare(CompareOp.Lte, BigDecimal(100))))
    assert(!matches(cell, Compare(CompareOp.Lte, BigDecimal(25))))
  }

  test("matches: not equal") {
    val cell = CellValue.Number(BigDecimal(50))
    assert(matches(cell, Compare(CompareOp.Neq, BigDecimal(0))))
    assert(matches(cell, Compare(CompareOp.Neq, BigDecimal(100))))
    assert(!matches(cell, Compare(CompareOp.Neq, BigDecimal(50))))
  }

  test("matches: comparison with text containing number") {
    val cell = CellValue.Text("150")
    assert(matches(cell, Compare(CompareOp.Gt, BigDecimal(100))))
  }

  test("matches: comparison with non-numeric text fails") {
    val cell = CellValue.Text("abc")
    assert(!matches(cell, Compare(CompareOp.Gt, BigDecimal(100))))
  }

  test("matches: comparison with boolean (TRUE=1, FALSE=0)") {
    val trueCell = CellValue.Bool(true)
    val falseCell = CellValue.Bool(false)
    assert(matches(trueCell, Compare(CompareOp.Gt, BigDecimal(0))))
    assert(!matches(falseCell, Compare(CompareOp.Gt, BigDecimal(0))))
  }

  test("matches: comparison with empty cell fails") {
    assert(!matches(CellValue.Empty, Compare(CompareOp.Gt, BigDecimal(0))))
  }

  // ===== Wildcard Matching Tests =====

  test("matches: wildcard * matches any characters") {
    val cell = CellValue.Text("Apple")
    assert(matches(cell, Wildcard("A*")))
    assert(matches(cell, Wildcard("*pple")))
    assert(matches(cell, Wildcard("*pp*")))
    assert(matches(cell, Wildcard("*")))
    assert(!matches(cell, Wildcard("B*")))
  }

  test("matches: wildcard ? matches single character") {
    val cell = CellValue.Text("Apple")
    assert(matches(cell, Wildcard("A?ple")))
    assert(matches(cell, Wildcard("?????")))
    assert(!matches(cell, Wildcard("????")))
    assert(!matches(cell, Wildcard("??????")))
  }

  test("matches: wildcard case-insensitive") {
    val cell = CellValue.Text("APPLE")
    assert(matches(cell, Wildcard("a*")))
    assert(matches(cell, Wildcard("*PPLE")))
    assert(matches(cell, Wildcard("apple")))
  }

  test("matches: wildcard with escaped asterisk") {
    val cell = CellValue.Text("test*value")
    assert(matches(cell, Wildcard("test~*value")))
    assert(!matches(cell, Wildcard("test~*other")))
  }

  test("matches: wildcard with escaped question mark") {
    val cell = CellValue.Text("test?value")
    assert(matches(cell, Wildcard("test~?value")))
    assert(!matches(cell, Wildcard("test~?other")))
  }

  test("matches: wildcard with escaped tilde") {
    val cell = CellValue.Text("test~value")
    assert(matches(cell, Wildcard("test~~value")))
  }

  test("matches: wildcard on number cell") {
    val cell = CellValue.Number(BigDecimal(12345))
    assert(matches(cell, Wildcard("123*")))
    assert(matches(cell, Wildcard("*45")))
    assert(matches(cell, Wildcard("1234?")))
  }

  test("matches: wildcard on boolean cell") {
    val cell = CellValue.Bool(true)
    assert(matches(cell, Wildcard("TR*")))
    assert(matches(cell, Wildcard("*UE")))
  }

  test("matches: wildcard on empty cell fails") {
    assert(!matches(CellValue.Empty, Wildcard("*")))
  }

  test("matches: wildcard with regex metacharacters") {
    val cell = CellValue.Text("test.value")
    assert(matches(cell, Wildcard("test.value")))
    assert(matches(cell, Wildcard("test*value")))

    val cell2 = CellValue.Text("test[1]")
    assert(matches(cell2, Wildcard("test[1]")))
  }

  // ===== Integration Tests =====

  test("integration: SUMIF-style matching") {
    val fruits = List(
      CellValue.Text("Apple"),
      CellValue.Text("Banana"),
      CellValue.Text("Apple"),
      CellValue.Text("Cherry")
    )

    val criterion = parse("Apple")
    val matchingCount = fruits.count(cell => matches(cell, criterion))
    assertEquals(matchingCount, 2)
  }

  test("integration: COUNTIF-style numeric comparison") {
    val numbers = List(
      CellValue.Number(BigDecimal(50)),
      CellValue.Number(BigDecimal(150)),
      CellValue.Number(BigDecimal(200)),
      CellValue.Number(BigDecimal(75))
    )

    val criterion = parse(">100")
    val matchingCount = numbers.count(cell => matches(cell, criterion))
    assertEquals(matchingCount, 2)
  }

  test("integration: wildcard filtering") {
    val products = List(
      CellValue.Text("Apple iPhone"),
      CellValue.Text("Apple MacBook"),
      CellValue.Text("Samsung Galaxy"),
      CellValue.Text("Apple Watch")
    )

    val criterion = parse("Apple*")
    val matchingCount = products.count(cell => matches(cell, criterion))
    assertEquals(matchingCount, 3)
  }
