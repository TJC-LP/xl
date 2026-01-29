package com.tjclp.xl.agent.benchmark.execution

import cats.kernel.Monoid
import io.circe.*
import io.circe.generic.semiauto.*

// ============================================================================
// Token Usage with Cache Tracking
// ============================================================================

/**
 * Token usage tracking including cache hits for prompt caching optimization.
 *
 * This is the unified token usage model used throughout the benchmark framework. It extends the
 * basic input/output tracking with cache creation and read tokens to support Anthropic's prompt
 * caching feature.
 *
 * @param inputTokens
 *   Number of input tokens (non-cached)
 * @param outputTokens
 *   Number of output tokens generated
 * @param cacheCreationTokens
 *   Tokens used to create new cache entries
 * @param cacheReadTokens
 *   Tokens read from existing cache
 */
case class TokenUsage(
  inputTokens: Long,
  outputTokens: Long,
  cacheCreationTokens: Long = 0,
  cacheReadTokens: Long = 0
):
  /** Total tokens (input + output, not including cache) */
  def total: Long = inputTokens + outputTokens

  /** Total tokens including cache creation */
  def totalWithCache: Long = inputTokens + outputTokens + cacheCreationTokens

  /** Effective input tokens (accounting for cache) */
  def effectiveInputTokens: Long = inputTokens + cacheReadTokens

  /** Cache hit rate (if there was any caching) */
  def cacheHitRate: Option[Double] =
    val totalCacheable = cacheCreationTokens + cacheReadTokens
    Option.when(totalCacheable > 0)(cacheReadTokens.toDouble / totalCacheable)

  /** Combine with another TokenUsage */
  def +(other: TokenUsage): TokenUsage = TokenUsage(
    inputTokens + other.inputTokens,
    outputTokens + other.outputTokens,
    cacheCreationTokens + other.cacheCreationTokens,
    cacheReadTokens + other.cacheReadTokens
  )

  /** Estimate cost in dollars based on model pricing */
  def estimatedCost(pricing: ModelPricing): BigDecimal =
    val inputCost = BigDecimal(inputTokens) * pricing.inputPerMillion / 1_000_000
    val outputCost = BigDecimal(outputTokens) * pricing.outputPerMillion / 1_000_000
    val cacheCost = BigDecimal(cacheCreationTokens) * pricing.cacheWritePerMillion / 1_000_000
    val cacheReadCost = BigDecimal(cacheReadTokens) * pricing.cacheReadPerMillion / 1_000_000
    inputCost + outputCost + cacheCost + cacheReadCost

object TokenUsage:
  /** Zero token usage */
  val zero: TokenUsage = TokenUsage(0, 0, 0, 0)

  /** Create from input/output only (no cache) */
  def apply(input: Long, output: Long): TokenUsage =
    TokenUsage(input, output, 0, 0)

  /** Create from the agent's TokenUsage type */
  def fromAgentUsage(usage: com.tjclp.xl.agent.TokenUsage): TokenUsage =
    TokenUsage(usage.inputTokens, usage.outputTokens, 0, 0)

  /** Monoid instance for combining token usage */
  given Monoid[TokenUsage] with
    def empty: TokenUsage = zero
    def combine(x: TokenUsage, y: TokenUsage): TokenUsage = x + y

  given Encoder[TokenUsage] = deriveEncoder
  given Decoder[TokenUsage] = Decoder.instance { c =>
    for
      input <- c.get[Long]("inputTokens")
      output <- c.get[Long]("outputTokens")
      cacheCreate <- c.getOrElse[Long]("cacheCreationTokens")(0)
      cacheRead <- c.getOrElse[Long]("cacheReadTokens")(0)
    yield TokenUsage(input, output, cacheCreate, cacheRead)
  }

// ============================================================================
// Model Pricing
// ============================================================================

/**
 * Pricing information for token cost estimation.
 *
 * All prices are in dollars per million tokens.
 */
case class ModelPricing(
  inputPerMillion: BigDecimal,
  outputPerMillion: BigDecimal,
  cacheWritePerMillion: BigDecimal = BigDecimal(0),
  cacheReadPerMillion: BigDecimal = BigDecimal(0)
)

object ModelPricing:
  /** Claude Opus 4.5 pricing (as of Jan 2025) */
  val opus45: ModelPricing = ModelPricing(
    inputPerMillion = BigDecimal("15.00"),
    outputPerMillion = BigDecimal("75.00"),
    cacheWritePerMillion = BigDecimal("18.75"),
    cacheReadPerMillion = BigDecimal("1.50")
  )

  /** Claude Sonnet 4 pricing */
  val sonnet4: ModelPricing = ModelPricing(
    inputPerMillion = BigDecimal("3.00"),
    outputPerMillion = BigDecimal("15.00"),
    cacheWritePerMillion = BigDecimal("3.75"),
    cacheReadPerMillion = BigDecimal("0.30")
  )

  /** Claude Haiku 3.5 pricing */
  val haiku35: ModelPricing = ModelPricing(
    inputPerMillion = BigDecimal("0.80"),
    outputPerMillion = BigDecimal("4.00"),
    cacheWritePerMillion = BigDecimal("1.00"),
    cacheReadPerMillion = BigDecimal("0.08")
  )

  /** Default pricing (Sonnet 4) */
  val default: ModelPricing = sonnet4

  /** Get pricing for a model by name */
  def forModel(model: String): ModelPricing =
    model.toLowerCase match
      case m if m.contains("opus") => opus45
      case m if m.contains("sonnet") => sonnet4
      case m if m.contains("haiku") => haiku35
      case _ => default
