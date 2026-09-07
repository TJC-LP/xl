---
name: xl-testing
description: How xl is tested - MUnit + ScalaCheck suite layout, the shared Generators, law patterns (round-trip, monoid, lens), the external-consumer prelude probes, real-file fixtures, the four release gates (tests, examples harness, skill snippets, probes), and the test-count bookkeeping. Use when adding or changing tests, deciding where a spec goes, reproducing a CI failure, or preparing a PR that touches the public surface.
---

# Testing in xl

Frameworks: MUnit 1.0.3 (`FunSuite`, `ScalaCheckSuite`), munit-scalacheck 1.2.0, ScalaCheck
1.18.1, munit-cats-effect 2.x and scalacheck-effect for `IO` suites in `xl-cats-effect`. Reference:
`docs/reference/testing-guide.md`. Count as of this writing: 5,455 tests, all deterministic, none
flaky.

## Where a test goes

`<module>/test/src/com/tjclp/xl/<same package as the code>/<Thing>Spec.scala`, package matching
the source under test. Shared generators: `xl-core/test/src/com/tjclp/xl/Generators.scala`
(`import com.tjclp.xl.Generators.given`). Another module needs test code or fixtures from a
sibling? Depend on the sibling's test module in `package.mill`, the way `xl-cats-effect` picks up
the `xl-ooxml` fixture corpus; never override `resources` in a nested test object (Mill bug, see
`mill-build`):

```scala
object test extends ScalaTests with build.XLTestModule {
  override def moduleDeps = super.moduleDeps ++ Seq(build.`xl-ooxml`.test)
}
```

```scala
package com.tjclp.xl.addressing

import com.tjclp.xl.Generators.given
import munit.ScalaCheckSuite
import org.scalacheck.Prop.*

class ColumnSpec extends ScalaCheckSuite:
  test("parse accepts a bare letter") {
    assertEquals(Column.parse("D"), Right(Column.from0(3)))
  }
  property("parse . toLetter = id") {
    forAll { (c: Column) => assertEquals(Column.parse(c.toLetter), Right(c)) }
  }
```

## Patterns that carry the guarantees

- **Round-trip laws** for every parser/printer pair: `parse(print(x)) == Right(x)`, and for
  OOXML `read(write(wb)) ≈ wb` (`OoxmlGenerativeRoundTripSpec`).
- **Monoid laws** for `Patch` and `StylePatch` (associativity, identity) with generated values.
- **Lens laws** for optics (get-put, put-get, put-put).
- **Determinism**: writing the same workbook twice yields identical bytes; tests compare hashes,
  not just structure.
- **Behavior pins**: when a fix changes observable behavior, add a test that names the issue
  (`GH-361`) so the pin is discoverable.
- **External tools are optional**: rasterizer backends (cairosvg, rsvg-convert, resvg, ImageMagick)
  are probed, never required; `RasterizerListSpec` mocks availability. Do not add tests that need
  LibreOffice or a system binary.

## Running

```bash
./mill xl-core.test                                                   # one module
./mill xl-core.test.testOnly com.tjclp.xl.addressing.ColumnSpec       # one suite (FQN)
./mill xl-core.test.testOnly com.tjclp.xl.addressing.ColumnSpec -- '*parse*'   # glob on test names
./mill -i __.test                                                     # everything, as CI does
```

Full runs take minutes: give the Bash tool a 600000 ms timeout. Mill prints
`N/N, SUCCESS` on a green run; a red run lists failures with `==> X`.

## The external-consumer probes

`xl/test/src/xlprelude/` (`xlprelude.ScriptingPreludeTest`, `BaseImportProbe`) lives outside
`com.tjclp.xl` on purpose: code inside the package inherits its exports and real types, so only a
probe from outside sees what a script or downstream project sees. Any change to exports, the
scripting prelude, opaque-type members, or public signatures must be exercised there (see the
landmines in `xl-scala-style`).

## Fixtures

`xl-ooxml/test/resources/fixtures/*.xlsx` is a synthetic corpus produced by
`scripts/generate-fixtures.py` (openpyxl, plus LibreOffice conversions as `*-lo.xlsx`). It is
deterministic (fixed timestamps, no randomness) and shared with `xl-cats-effect` tests through
`moduleDeps`. Do not hand-edit fixtures; regenerate with the script and note that the `-lo` files
embed LibreOffice metadata and are not byte-stable.

## CLI contract goldens

`xl-cli/test/resources/golden/<case>.golden` pins the `xl` binary's agent-visible contract — exit
code, stdout and stderr per invocation — through the in-process `CliHarness` (`Cli.run` with an
injected `CliIO`, in `xl-cli/test/src/com/tjclp/xl/cli/contract/`). A golden diff is a contract
change, not noise: review it, then re-record deliberately with
`XL_UPDATE_GOLDEN=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.GoldenSpec` and add a
CHANGELOG line. New CLI behaviour gets a new `.golden` (write `## args`/`## stdin`, record). Details
in `docs/reference/testing-guide.md` ("CLI contract goldens").

## Gates before a PR touching the public surface

Run all four, with `set -o pipefail` when piping:

1. `./mill __.test`
2. `scripts/test-examples.sh` (`--compile-only` to skip running): checks that
   `examples/project.scala` pins the same version as `build.mill`, publishes the build to
   `ivy2Local`, compiles every `examples/*.sc`, runs a curated subset. Needs `scala-cli`.
3. `scripts/verify-skill-snippets.sh --local`: compiles every `//> using` fenced snippet in the
   xl-scripting skill and `docs/reference/scripting.md` against the local build.
4. The prelude probes (part of `xl.test`, already in gate 1, but read their failures first: they
   are the consumer-facing ones).

Docs-only PRs need none of these.

## Bookkeeping when suites change

Per-module counts are recorded in three places and CI does not check them: `CLAUDE.md`
("Run all tests (N)" and the Testing section), `docs/STATUS.md`, and the table in
`docs/reference/testing-guide.md`. Update all three when the total moves; get the numbers from
`./mill <module>.test` output rather than estimating.
