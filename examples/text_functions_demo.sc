#!/usr/bin/env -S scala-cli shebang
//> using file project.scala


// Demonstrates the 6 text functions added in TJC-1055 / GH-116:
//   TRIM, MID, FIND, SUBSTITUTE, VALUE, TEXT
//
// Each section: build a small workbook, apply realistic formulas, and print
// "formula = result   (expected: ...)" so any divergence pops out visually.
//
// Run with:
//   1. Publish locally: ./mill xl.publishLocal
//   2. Run script:      scala-cli run examples/text_functions_demo.sc

import com.tjclp.xl.{*, given}
import com.tjclp.xl.cells.CellValue

println("=== XL Text Functions Demo (TJC-1055 / GH-116) ===\n")

/** Evaluate a formula on the given sheet and stringify the result. */
def eval(formula: String, sheet: Sheet): String =
  sheet.evaluateFormula(formula) match
    case Right(CellValue.Text(s)) => s"\"$s\""
    case Right(CellValue.Number(n)) => n.toString
    case Right(CellValue.Bool(b)) => b.toString
    case Right(other) => other.toString
    case Left(err) => s"<ERROR: $err>"

/** Print a formula result alongside the expected value. Mismatches visually pop. */
def show(formula: String, sheet: Sheet, expected: String): Unit =
  val got = eval(formula, sheet)
  val mark = if got == expected then "✓" else "✗"
  println(f"  $mark%s  $formula%-50s = $got%-30s (expected: $expected)")


// =====================================================================
// 1. TRIM + SUBSTITUTE — clean messy CSV-imported data
// =====================================================================
println("\n--- 1. Cleanup pipeline (TRIM, SUBSTITUTE) ---")

val cleanup = Sheet("Cleanup")
  .put(ref"A1", CellValue.Text("  alice@example.com  "))
  .put(ref"A2", CellValue.Text("Name: Bob; Age: 42"))
  .put(ref"A3", CellValue.Text("a,b,,c,,,d"))

show("=TRIM(A1)", cleanup, "\"alice@example.com\"")
show("=SUBSTITUTE(A2, \"; \", \" | \")", cleanup, "\"Name: Bob | Age: 42\"")
show("=SUBSTITUTE(A3, \",,\", \",\")", cleanup, "\"a,b,c,,d\"")
show("=SUBSTITUTE(SUBSTITUTE(A3, \",,\", \",\"), \",,\", \",\")", cleanup, "\"a,b,c,d\"")


// =====================================================================
// 2. VALUE — parse currency / percent / accounting strings
// =====================================================================
println("\n--- 2. Numeric parsing (VALUE) ---")

val parsing = Sheet("Parsing")
  .put(ref"A1", CellValue.Text("$1,234.56"))
  .put(ref"A2", CellValue.Text("(500)"))
  .put(ref"A3", CellValue.Text("45.5%"))
  .put(ref"A4", CellValue.Text(" $-1,000 "))

show("=VALUE(A1)", parsing, "1234.56")
show("=VALUE(A2)", parsing, "-500")
show("=VALUE(A3)", parsing, "0.455")
show("=VALUE(A4)", parsing, "-1000")


// =====================================================================
// 3. TEXT — format numbers / dates for display
// =====================================================================
println("\n--- 3. Display formatting (TEXT) ---")

val formatting = Sheet("Formatting")
  .put(ref"A1", CellValue.Number(BigDecimal("1234567.89")))
  .put(ref"A2", CellValue.Number(BigDecimal("0.075")))
  .put(ref"A3", CellValue.Number(BigDecimal("-1234.5")))

show("=TEXT(A1, \"#,##0.00\")", formatting, "\"1,234,567.89\"")
show("=TEXT(A2, \"0.00%\")", formatting, "\"7.50%\"")
show("=TEXT(A3, \"#,##0.00;-#,##0.00\")", formatting, "\"-1,234.50\"")
show("=TEXT(A1, \"0\")", formatting, "\"1234568\"")


// =====================================================================
// 4. FIND + MID — extract email domain (function composition)
// =====================================================================
println("\n--- 4. Extract email domain (FIND + MID) ---")

val emails = Sheet("Emails")
  .put(ref"A1", CellValue.Text("alice@example.com"))
  .put(ref"A2", CellValue.Text("bob@tjclp.com"))
  .put(ref"A3", CellValue.Text("charlie+filter@gmail.co.uk"))

// =MID(A1, FIND("@", A1) + 1, 100)  — MID handles overflow by clamping
show("=MID(A1, FIND(\"@\", A1) + 1, 100)", emails, "\"example.com\"")
show("=MID(A2, FIND(\"@\", A2) + 1, 100)", emails, "\"tjclp.com\"")
show("=MID(A3, FIND(\"@\", A3) + 1, 100)", emails, "\"gmail.co.uk\"")


// =====================================================================
// 5. Round-trip: TEXT(VALUE(s)) — normalize messy currency input
// =====================================================================
println("\n--- 5. Round-trip: messy → number → canonical (TEXT(VALUE(...))) ---")

val roundtrip = Sheet("Roundtrip")
  .put(ref"A1", CellValue.Text("$1,234.56"))
  .put(ref"A2", CellValue.Text("(2,500)"))
  .put(ref"A3", CellValue.Text("78.9%"))

show("=TEXT(VALUE(A1), \"#,##0.00\")", roundtrip, "\"1,234.56\"")
show("=TEXT(VALUE(A2), \"#,##0.00;-#,##0.00\")", roundtrip, "\"-2,500.00\"")
show("=TEXT(VALUE(A3), \"0.00%\")", roundtrip, "\"78.90%\"")


println("\n=== Demo Complete ===")
println("Tip: change a formula above and re-run to explore behavior.")
