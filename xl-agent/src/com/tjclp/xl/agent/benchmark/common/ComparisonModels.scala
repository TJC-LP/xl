package com.tjclp.xl.agent.benchmark.common

import io.circe.*
import io.circe.generic.semiauto.*

/** Full usage breakdown from API response */
case class UsageBreakdown(
  inputTokens: Long,
  outputTokens: Long,
  cacheCreationInputTokens: Long = 0,
  cacheReadInputTokens: Long = 0
):
  def totalTokens: Long = inputTokens + outputTokens

  /** Calculate cost using model pricing (per 1M tokens) */
  def estimatedCost(pricing: ModelPricing): Double =
    val inputCost = inputTokens * pricing.inputPerMillion / 1_000_000.0
    val outputCost = outputTokens * pricing.outputPerMillion / 1_000_000.0
    val cacheWriteCost = cacheCreationInputTokens * pricing.cacheWritePerMillion / 1_000_000.0
    val cacheReadCost = cacheReadInputTokens * pricing.cacheReadPerMillion / 1_000_000.0
    inputCost + outputCost + cacheWriteCost + cacheReadCost

object UsageBreakdown:
  given Encoder[UsageBreakdown] = deriveEncoder
  given Decoder[UsageBreakdown] = deriveDecoder

  def fromTokenUsage(
    inputTokens: Long,
    outputTokens: Long,
    cacheCreation: Option[Long] = None,
    cacheRead: Option[Long] = None
  ): UsageBreakdown =
    UsageBreakdown(
      inputTokens = inputTokens,
      outputTokens = outputTokens,
      cacheCreationInputTokens = cacheCreation.getOrElse(0L),
      cacheReadInputTokens = cacheRead.getOrElse(0L)
    )

/** Model pricing (per 1M tokens) */
case class ModelPricing(
  inputPerMillion: Double,
  outputPerMillion: Double,
  cacheWritePerMillion: Double,
  cacheReadPerMillion: Double
)

object ModelPricing:
  given Encoder[ModelPricing] = deriveEncoder
  given Decoder[ModelPricing] = deriveDecoder

  // Claude Opus 4.5 pricing (as of Jan 2026)
  val Opus4 = ModelPricing(
    inputPerMillion = 15.0,
    outputPerMillion = 75.0,
    cacheWritePerMillion = 18.75,
    cacheReadPerMillion = 1.50
  )

  // Claude Sonnet 4 pricing
  val Sonnet4 = ModelPricing(
    inputPerMillion = 3.0,
    outputPerMillion = 15.0,
    cacheWritePerMillion = 3.75,
    cacheReadPerMillion = 0.30
  )

  def forModel(modelName: String): ModelPricing =
    if modelName.contains("opus") then Opus4
    else Sonnet4

/** Result of running one approach on a task */
case class ApproachResult(
  approach: String,
  success: Boolean,
  passed: Boolean,
  usage: UsageBreakdown,
  latencyMs: Long,
  error: Option[String] = None
)

object ApproachResult:
  given Encoder[ApproachResult] = deriveEncoder
  given Decoder[ApproachResult] = deriveDecoder

/** Comparison of both approaches on same task */
case class TaskComparison(
  taskId: String,
  taskName: String,
  xl: Option[ApproachResult],
  xlsx: Option[ApproachResult]
):
  def inputTokenDiff: Long =
    xl.map(_.usage.inputTokens).getOrElse(0L) -
      xlsx.map(_.usage.inputTokens).getOrElse(0L)

  def outputTokenDiff: Long =
    xl.map(_.usage.outputTokens).getOrElse(0L) -
      xlsx.map(_.usage.outputTokens).getOrElse(0L)

  def costDiff(pricing: ModelPricing): Double =
    xl.map(_.usage.estimatedCost(pricing)).getOrElse(0.0) -
      xlsx.map(_.usage.estimatedCost(pricing)).getOrElse(0.0)

object TaskComparison:
  given Encoder[TaskComparison] = deriveEncoder
  given Decoder[TaskComparison] = deriveDecoder

/** Aggregate comparison stats */
case class ComparisonStats(
  totalTasks: Int,
  xlPassed: Int,
  xlsxPassed: Int,
  xlTotalInputTokens: Long,
  xlTotalOutputTokens: Long,
  xlsxTotalInputTokens: Long,
  xlsxTotalOutputTokens: Long,
  xlEstimatedCost: Double,
  xlsxEstimatedCost: Double
):
  def costSavingsPercent: Double =
    if xlsxEstimatedCost > 0 then (xlsxEstimatedCost - xlEstimatedCost) / xlsxEstimatedCost * 100
    else 0.0

object ComparisonStats:
  given Encoder[ComparisonStats] = deriveEncoder
  given Decoder[ComparisonStats] = deriveDecoder

  def fromResults(results: List[TaskComparison], pricing: ModelPricing): ComparisonStats =
    val xlResults = results.flatMap(_.xl)
    val xlsxResults = results.flatMap(_.xlsx)

    ComparisonStats(
      totalTasks = results.size,
      xlPassed = xlResults.count(_.passed),
      xlsxPassed = xlsxResults.count(_.passed),
      xlTotalInputTokens = xlResults.map(_.usage.inputTokens).sum,
      xlTotalOutputTokens = xlResults.map(_.usage.outputTokens).sum,
      xlsxTotalInputTokens = xlsxResults.map(_.usage.inputTokens).sum,
      xlsxTotalOutputTokens = xlsxResults.map(_.usage.outputTokens).sum,
      xlEstimatedCost = xlResults.map(_.usage.estimatedCost(pricing)).sum,
      xlsxEstimatedCost = xlsxResults.map(_.usage.estimatedCost(pricing)).sum
    )

/** Full comparison benchmark report */
case class ComparisonReport(
  timestamp: String,
  dataset: String,
  model: String,
  pricing: ModelPricing,
  stats: ComparisonStats,
  results: List[TaskComparison]
):
  def toJsonPretty: String =
    import io.circe.syntax.*
    this.asJson.spaces2

object ComparisonReport:
  given Encoder[ComparisonReport] = deriveEncoder
  given Decoder[ComparisonReport] = deriveDecoder
