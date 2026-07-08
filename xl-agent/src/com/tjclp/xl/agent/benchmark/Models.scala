package com.tjclp.xl.agent.benchmark

// ============================================================================
// Centralized Model Configuration
// ============================================================================

/**
 * Single source of truth for model identifiers and pricing.
 *
 * This consolidates model configuration that was previously scattered across multiple files:
 *   - TokenUsage.scala (ModelPricing)
 *   - ComparisonModels.scala (ModelPricing)
 *   - LLMGrader.scala (DefaultModel)
 *   - TokenGrader.scala (Model)
 *   - UnifiedRunner.scala (model default)
 *   - AgentConfig (model default)
 */
object Models:

  // --------------------------------------------------------------------------
  // Model Identifiers (Claude 5 family + current Opus/Haiku)
  //
  // Aliases (not dated snapshots) so new snapshots are picked up automatically.
  // --------------------------------------------------------------------------

  /** Claude Fable 5 - most capable model, for the hardest tasks (premium pricing) */
  val Fable5: String = "claude-fable-5"

  /** Claude Opus 4.8 - highest-capability Opus tier, best for grading */
  val Opus48: String = "claude-opus-4-8"

  /** Claude Sonnet 5 - near-Opus coding/agentic quality at Sonnet cost, default for agents */
  val Sonnet5: String = "claude-sonnet-5"

  /** Claude Haiku 4.5 - fastest and cheapest, good for simple tasks */
  val Haiku45: String = "claude-haiku-4-5"

  /** Default model for agent execution */
  val DefaultAgent: String = Sonnet5

  /** Default model for grading (uses Opus for accuracy) */
  val DefaultGrader: String = Opus48

  // --------------------------------------------------------------------------
  // Pricing
  // --------------------------------------------------------------------------

  /**
   * Pricing information for token cost estimation.
   *
   * All prices are in dollars per million tokens. Uses BigDecimal for financial precision.
   *
   * @param inputPerMillion
   *   Cost per million input tokens
   * @param outputPerMillion
   *   Cost per million output tokens
   * @param cacheWritePerMillion
   *   Cost per million tokens for cache creation
   * @param cacheReadPerMillion
   *   Cost per million tokens for cache reads
   */
  case class Pricing(
    inputPerMillion: BigDecimal,
    outputPerMillion: BigDecimal,
    cacheWritePerMillion: BigDecimal = BigDecimal(0),
    cacheReadPerMillion: BigDecimal = BigDecimal(0)
  ):
    /** Estimate cost for given token counts */
    def estimateCost(inputTokens: Long, outputTokens: Long): BigDecimal =
      (BigDecimal(inputTokens) * inputPerMillion / 1_000_000) +
        (BigDecimal(outputTokens) * outputPerMillion / 1_000_000)

    /** Estimate cost including cache tokens */
    def estimateCostWithCache(
      inputTokens: Long,
      outputTokens: Long,
      cacheCreateTokens: Long,
      cacheReadTokens: Long
    ): BigDecimal =
      estimateCost(inputTokens, outputTokens) +
        (BigDecimal(cacheCreateTokens) * cacheWritePerMillion / 1_000_000) +
        (BigDecimal(cacheReadTokens) * cacheReadPerMillion / 1_000_000)

  /** Claude Fable 5 / Mythos 5 pricing (as of Jun 2026) */
  val FablePricing: Pricing = Pricing(
    inputPerMillion = BigDecimal("10.00"),
    outputPerMillion = BigDecimal("50.00"),
    cacheWritePerMillion = BigDecimal("12.50"),
    cacheReadPerMillion = BigDecimal("1.00")
  )

  /** Claude Opus 4.x pricing (as of Jun 2026) */
  val OpusPricing: Pricing = Pricing(
    inputPerMillion = BigDecimal("5.00"),
    outputPerMillion = BigDecimal("25.00"),
    cacheWritePerMillion = BigDecimal("6.25"),
    cacheReadPerMillion = BigDecimal("0.50")
  )

  /** Claude Sonnet pricing (standard rate; Sonnet 5 has intro $2/$10 through 2026-08-31) */
  val SonnetPricing: Pricing = Pricing(
    inputPerMillion = BigDecimal("3.00"),
    outputPerMillion = BigDecimal("15.00"),
    cacheWritePerMillion = BigDecimal("3.75"),
    cacheReadPerMillion = BigDecimal("0.30")
  )

  /** Claude Haiku 4.5 pricing (as of Jun 2026) */
  val HaikuPricing: Pricing = Pricing(
    inputPerMillion = BigDecimal("1.00"),
    outputPerMillion = BigDecimal("5.00"),
    cacheWritePerMillion = BigDecimal("1.25"),
    cacheReadPerMillion = BigDecimal("0.10")
  )

  /** Default pricing (Sonnet) */
  val DefaultPricing: Pricing = SonnetPricing

  /** Get pricing for a model by name */
  def pricingFor(model: String): Pricing =
    model.toLowerCase match
      case m if m.contains("fable") || m.contains("mythos") => FablePricing
      case m if m.contains("opus") => OpusPricing
      case m if m.contains("haiku") => HaikuPricing
      case _ => SonnetPricing
