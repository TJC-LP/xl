package com.tjclp.xl.formula.eval

import munit.ScalaCheckSuite
import org.scalacheck.rng.Seed

/**
 * The law-fuzz seed convention (GH-308) for property suites: 500 successes (or
 * `xl.roundtrip.minSuccess` / `XL_ROUNDTRIP_MIN_SUCCESS`), seeded with `defaultSeed`; a random
 * seed, printed with `label`, under `XL_ROUNDTRIP_RANDOM_SEED=true`; replayable with
 * `-Dxl.roundtrip.seed=<printed seed>`.
 */
abstract class LawFuzzSuite(label: String, defaultSeed: Long) extends ScalaCheckSuite:

  override def scalaCheckTestParameters: org.scalacheck.Test.Parameters =
    val minSuccess = sys.props
      .get("xl.roundtrip.minSuccess")
      .orElse(sys.env.get("XL_ROUNDTRIP_MIN_SUCCESS"))
      .flatMap(_.toIntOption)
      .getOrElse(500)
    val explicitSeed = sys.props.get("xl.roundtrip.seed").flatMap { raw =>
      raw.toLongOption.map(Seed.apply).orElse(Seed.fromBase64(raw).toOption)
    }
    val base = super.scalaCheckTestParameters.withMinSuccessfulTests(minSuccess)
    val random =
      explicitSeed.isEmpty &&
        sys.env.get("XL_ROUNDTRIP_RANDOM_SEED").exists(_.equalsIgnoreCase("true"))
    if random then
      val seed = Seed.random()
      println(s"[law-fuzz] random seed for this run: ${seed.toBase64} ($label)")
      base.withInitialSeed(seed)
    else base.withInitialSeed(explicitSeed.getOrElse(Seed(defaultSeed)))
