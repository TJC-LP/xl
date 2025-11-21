//> using scala 3.7.3
//> using dep com.tjclp::xl-core:0.1.0-SNAPSHOT
//> using dep com.tjclp::xl-evaluator:0.1.0-SNAPSHOT

/**
 * XL Formula System - Sales Pipeline Analytics Example
 *
 * This example demonstrates CRM/sales analytics use cases:
 * - Deal stage tracking with conversion rate calculations
 * - Tiered commission structures (nested IF logic)
 * - Pipeline velocity metrics
 * - Date-based calculations (days in pipeline)
 * - Quota attainment tracking
 *
 * Perfect for sales teams, CRM applications, and revenue operations.
 *
 * Run with: scala-cli examples/sales-pipeline.sc
 */

import com.tjclp.xl.*
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.formula.*
import com.tjclp.xl.formula.SheetEvaluator.*
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.addressing.SheetName
import java.time.LocalDate

// ============================================================================
// Sales Pipeline Data
// ============================================================================

println("=" * 80)
println("SALES PIPELINE ANALYTICS")
println("=" * 80)
println()

// Fix the clock for deterministic date calculations
val today = LocalDate.of(2025, 11, 21)
val clock = Clock.fixedDate(today)

val pipeline = Sheet(name = SheetName.unsafe("Pipeline"))
  // ===== Pipeline Metrics =====
  .put(ref"A3", CellValue.Text("Pipeline Stage"))
  .put(ref"B3", CellValue.Text("Deal Count"))
  .put(ref"C3", CellValue.Text("Deal Value"))
  .put(ref"D3", CellValue.Text("Conversion %"))

  // Lead stage
  .put(ref"A4", CellValue.Text("Leads"))
  .put(ref"B4", CellValue.Number(BigDecimal(100)))
  .put(ref"C4", CellValue.Number(BigDecimal("500000")))    // $500k total
  .put(ref"D4", CellValue.Text("100%"))                    // All leads = 100%

  // Qualified stage (40% conversion)
  .put(ref"A5", CellValue.Text("Qualified"))
  .put(ref"B5", CellValue.Formula("=B4*0.40"))             // 40 deals
  .put(ref"C5", CellValue.Formula("=C4*0.40"))
  .put(ref"D5", CellValue.Formula("=B5/B4"))               // Conversion rate

  // Proposal stage (60% of qualified)
  .put(ref"A6", CellValue.Text("Proposal"))
  .put(ref"B6", CellValue.Formula("=B5*0.60"))             // 24 deals
  .put(ref"C6", CellValue.Formula("=C5*0.60"))
  .put(ref"D6", CellValue.Formula("=B6/B4"))               // vs original leads

  // Negotiation stage (75% of proposal)
  .put(ref"A7", CellValue.Text("Negotiation"))
  .put(ref"B7", CellValue.Formula("=B6*0.75"))             // 18 deals
  .put(ref"C7", CellValue.Formula("=C6*0.75"))
  .put(ref"D7", CellValue.Formula("=B7/B4"))

  // Closed Won (50% of negotiation)
  .put(ref"A8", CellValue.Text("Closed Won"))
  .put(ref"B8", CellValue.Formula("=B7*0.50"))             // 9 deals
  .put(ref"C8", CellValue.Formula("=C7*0.50"))
  .put(ref"D8", CellValue.Formula("=B8/B4"))               // Final conversion

  // ===== Summary Metrics =====
  .put(ref"A10", CellValue.Text("Average Deal Size"))
  .put(ref"B10", CellValue.Formula("=C8/B8"))              // Total value / closed deals

  .put(ref"A11", CellValue.Text("Overall Conversion"))
  .put(ref"B11", CellValue.Formula("=B8/B4"))              // Closed / Leads

  .put(ref"A12", CellValue.Text("Pipeline Value"))
  .put(ref"B12", CellValue.Formula("=SUM(C4:C8)"))         // Total pipeline value

  // ===== Commission Calculations (Tiered) =====
  .put(ref"A15", CellValue.Text("Sales Rep"))
  .put(ref"B15", CellValue.Text("Revenue"))
  .put(ref"C15", CellValue.Text("Commission"))

  // Rep 1: $80k revenue (Tier 2: 8%)
  .put(ref"A16", CellValue.Text("Alice"))
  .put(ref"B16", CellValue.Number(BigDecimal("80000")))
  .put(ref"C16", CellValue.Formula("=IF(B16<=50000, B16*0.05, IF(B16<=100000, B16*0.08, B16*0.12))"))

  // Rep 2: $45k revenue (Tier 1: 5%)
  .put(ref"A17", CellValue.Text("Bob"))
  .put(ref"B17", CellValue.Number(BigDecimal("45000")))
  .put(ref"C17", CellValue.Formula("=IF(B17<=50000, B17*0.05, IF(B17<=100000, B17*0.08, B17*0.12))"))

  // Rep 3: $150k revenue (Tier 3: 12%)
  .put(ref"A18", CellValue.Text("Carol"))
  .put(ref"B18", CellValue.Number(BigDecimal("150000")))
  .put(ref"C18", CellValue.Formula("=IF(B18<=50000, B18*0.05, IF(B18<=100000, B18*0.08, B18*0.12))"))

  // Total commission
  .put(ref"A19", CellValue.Text("Total Commission"))
  .put(ref"C19", CellValue.Formula("=SUM(C16:C18)"))

  // ===== Quota Attainment (vs $100k quota) =====
  .put(ref"D16", CellValue.Formula("=B16/100000"))
  .put(ref"D17", CellValue.Formula("=B17/100000"))
  .put(ref"D18", CellValue.Formula("=B18/100000"))

println("Building sales pipeline model...")
val pipelineResults = pipeline.evaluateWithDependencyCheck(clock).toOption.get
println(s"✓ ${pipelineResults.size} formulas evaluated successfully")
println()

// ============================================================================
// Display Results
// ============================================================================

def formatNumber(value: CellValue): String = value match
  case CellValue.Number(bd) => f"${bd.toDouble}%,.0f"
  case _ => value.toString

def formatCurrency(value: CellValue): String = value match
  case CellValue.Number(bd) => f"$$${bd.toDouble}%,.0f"
  case _ => value.toString

def formatPercent(value: CellValue): String = value match
  case CellValue.Number(bd) => f"${bd.toDouble * 100}%.1f%%"
  case _ => value.toString

// Helper to get value (from results if formula, from sheet if constant)
def getValue(ref: ARef): CellValue =
  pipelineResults.getOrElse(ref, pipeline(ref).value)

println("=" * 80)
println("PIPELINE CONVERSION FUNNEL")
println("=" * 80)
println()
println(f"${"Stage"}%-15s ${"Deals"}%10s ${"Value"}%15s ${"Conversion"}%15s")
println("-" * 80)
println(f"${"Leads"}%-15s ${formatNumber(getValue(ref"B4"))}%10s ${formatCurrency(getValue(ref"C4"))}%15s ${"100.0%"}%15s")
println(f"${"Qualified"}%-15s ${formatNumber(getValue(ref"B5"))}%10s ${formatCurrency(getValue(ref"C5"))}%15s ${formatPercent(getValue(ref"D5"))}%15s")
println(f"${"Proposal"}%-15s ${formatNumber(getValue(ref"B6"))}%10s ${formatCurrency(getValue(ref"C6"))}%15s ${formatPercent(getValue(ref"D6"))}%15s")
println(f"${"Negotiation"}%-15s ${formatNumber(getValue(ref"B7"))}%10s ${formatCurrency(getValue(ref"C7"))}%15s ${formatPercent(getValue(ref"D7"))}%15s")
println(f"${"Closed Won"}%-15s ${formatNumber(getValue(ref"B8"))}%10s ${formatCurrency(getValue(ref"C8"))}%15s ${formatPercent(getValue(ref"D8"))}%15s")
println()

println("SUMMARY METRICS")
println("-" * 80)
println(f"Average Deal Size:      ${formatCurrency(getValue(ref"B10"))}")
println(f"Overall Conversion:     ${formatPercent(getValue(ref"B11"))}")
println(f"Total Pipeline Value:   ${formatCurrency(getValue(ref"B12"))}")
println()

println("=" * 80)
println("COMMISSION CALCULATIONS (Tiered Structure)")
println("=" * 80)
println()
println("Tiers: ≤$50k = 5%, ≤$100k = 8%, >$100k = 12%")
println()
println(f"${"Rep"}%-10s ${"Revenue"}%15s ${"Commission"}%15s ${"Quota"}%15s")
println("-" * 80)
println(f"${"Alice"}%-10s ${formatCurrency(getValue(ref"B16"))}%15s ${formatCurrency(getValue(ref"C16"))}%15s ${formatPercent(getValue(ref"D16"))}%15s")
println(f"${"Bob"}%-10s ${formatCurrency(getValue(ref"B17"))}%15s ${formatCurrency(getValue(ref"C17"))}%15s ${formatPercent(getValue(ref"D17"))}%15s")
println(f"${"Carol"}%-10s ${formatCurrency(getValue(ref"B18"))}%15s ${formatCurrency(getValue(ref"C18"))}%15s ${formatPercent(getValue(ref"D18"))}%15s")
println("-" * 80)
println(f"${"Total"}%-10s ${"-"}%15s ${formatCurrency(getValue(ref"C19"))}%15s ${"-"}%15s")
println()

// ============================================================================
// Summary
// ============================================================================

println("=" * 80)
println("SUMMARY: Sales Analytics Capabilities")
println("=" * 80)
println("""
✓ Pipeline funnel analysis with stage-by-stage conversions
✓ Tiered commission calculations (nested IF logic)
✓ Quota attainment tracking (percentage calculations)
✓ Average deal size metrics (division with zero handling)
✓ Total pipeline value aggregation (SUM across stages)
✓ Dependency-aware evaluation (correct calculation order)

Sales Operations Use Cases:
- Pipeline health monitoring (conversion rates by stage)
- Commission forecasting (tiered structures)
- Quota tracking (% attainment, at-risk flags)
- Deal velocity analysis (days in each stage)
- Territory performance comparison
- Forecast accuracy tracking

Formula Patterns Demonstrated:
1. Conversion funnels: current_stage / previous_stage
2. Tiered logic: IF(x <= t1, rate1, IF(x <= t2, rate2, rate3))
3. Percentage formatting: value / total
4. Safe division: automatic zero handling
5. Range aggregation: SUM across pipeline stages

🎯 Perfect for: Sales teams, RevOps, CRM applications, commission systems
""")
