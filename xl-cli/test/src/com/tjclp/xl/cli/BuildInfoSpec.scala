package com.tjclp.xl.cli

import munit.FunSuite

/**
 * BuildInfo.version is a Mill-generated compile-time constant (ADR-016: no
 * getResourceAsStream/Properties at runtime), so it must always carry the real build version —
 * there is no "dev" resource-missing fallback left to hit.
 */
class BuildInfoSpec extends FunSuite:

  test("BuildInfo.version is the build version, not a fallback"):
    assert(
      BuildInfo.version.matches("""\d+\.\d+\.\d+.*"""),
      s"expected a semver-shaped version, got: ${BuildInfo.version}"
    )
