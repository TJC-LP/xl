---
name: mill-build
description: Build, test, format, and package this repo with Mill 1.1 - targeted module and suite commands, the JVM pin, what CI actually runs, realistic timings, and the Zinc/macro/resources gotchas. Use whenever you are about to run ./mill, want a faster loop than __.compile / __.test, hit a strange compile error after editing FunctionSpecs*, need the CLI jar or native binary, or are editing build.mill / package.mill.
---

# Mill build for xl

Mill 1.1.5 (`.mill-version`), Scala 3.9.0, JDK Temurin 25 (`.mill-jvm-version`). The `./mill`
launcher downloads Mill and the JDK through Coursier, so the only prerequisites are `bash`, `curl`,
and network access to GitHub + Maven Central. Never install a JDK to satisfy the build and never
change `javacOptions` to match a local JDK.

## Modules

| Module | Test module | What it holds |
| --- | --- | --- |
| `xl-core` | `xl-core.test` | domain model, macros, DSL, codecs, styles, `Generators.scala` |
| `xl-ooxml` | `xl-ooxml.test` | OOXML read/write, fixtures under `xl-ooxml/test/resources/fixtures` |
| `xl-cats-effect` | `xl-cats-effect.test` | `Excel[F]`, `ExcelIO`, SAX streaming |
| `xl-evaluator` | `xl-evaluator.test` | formula parser/evaluator, `FunctionSpecs*`, dependency graph |
| `xl-cli` | `xl-cli.test` | the `xl` CLI (`mainClass` = `com.tjclp.xl.cli.Main`) |
| `xl-agent` | `xl-agent.test` | benchmark runner |
| `xl-benchmarks` | none | JMH (`contrib.jmh.JmhModule`) |
| `xl` | `xl.test` | aggregate + scripting prelude; external-consumer probes in `xl/test/src/xlprelude/` |

Shared compiler settings live in `build.mill` (`XLModuleBase`); per-module deps in
`<module>/package.mill`. Dependency syntax: `mvn"org::artifact:version"` (Scala cross),
`mvn"org:::artifact:version"` (full Scala-version cross, used for the WartRemover plugin).

## Everyday commands

```bash
./mill __.compile                                  # everything, main + test sources
./mill xl-core.compile                             # one module (its deps compile first)
./mill xl-core.test                                # one module's tests
./mill xl-core.test.testOnly com.tjclp.xl.addressing.ColumnSpec              # one suite (FQN required)
./mill xl-core.test.testOnly com.tjclp.xl.addressing.ColumnSpec -- '*parse*' # tests whose name matches the glob
./mill -i __.test                                  # all tests, the way CI runs them
./mill mill.scalalib.scalafmt.ScalafmtModule/reformatAll __.sources     # format main + test sources
./mill mill.scalalib.scalafmt.ScalafmtModule/checkFormatAll __.sources  # the CI format check
./mill resolve __.test                             # list tasks; `./mill show xl-core.compileClasspath` prints a task
./mill clean                                       # wipe out/ (or `./mill clean xl-evaluator.compile` for one task)
./mill __.publishLocal                             # publish to ivy2Local (examples harness does this)
./mill __.prepareOffline                           # download every dependency, no compile (cache warm-up)
./mill xl-cli.assembly                             # out/xl-cli/assembly.dest/out.jar
./mill xl-cli.run -f data.xlsx sheets              # run the CLI from source (args follow the task)
./mill xl-cli.nativeImage                          # GraalVM native binary (local only)
make install-jar                                   # assembly + wrapper -> ~/.local/bin/xl (any JDK 17+)
make install                                       # native binary -> ~/.local/bin/xl (GraalVM)
```

`__.reformat` and `__.checkFormat` exist but cover main sources only, because test modules do not
mix in `ScalafmtModule`. CI runs the `__.sources` forms above, so use those.

## What CI runs

`.github/workflows/ci.yml`: `checkFormatAll __.sources` -> `./mill __.compile` -> `./mill -i __.test`,
plus an `examples` job running `scripts/test-examples.sh` (needs `scala-cli`). Release gates on top
of that: `scripts/verify-skill-snippets.sh --local`. Reproduce locally before pushing.

## Timings (set Bash timeouts accordingly)

| Situation | Expect |
| --- | --- |
| Warm incremental `./mill __.compile` | under a minute |
| `./mill __.test`, warm | a few minutes locally, about six in CI including compile |
| Fresh checkout, first `./mill` anything | dependency + JDK download first; allow 600000 ms |
| `__.publishLocal` + examples harness | several minutes |

Pass an explicit `timeout` (600000 ms) to the Bash tool for full compiles and test runs; do not
conclude a hang before it elapses. A second `./mill` invocation while one is running waits on the
`out/` lock instead of failing.

## Gotchas

- **`FunctionSpecs*` edits and "Cyclic reference involving trait TExprReferenceOps"** (or
  `TExprLookupOps` / `TExprAggregateOps`), pointing at files you never touched: the
  `FunctionRegistry.all` macro (`FunctionRegistryMacro.collect[FunctionSpecs.type]`) left Zinc's
  incremental state inconsistent. Not your code. Run `./mill clean xl-evaluator.compile`, then
  recompile. Scala 3.9 tracks macro type arguments in Zinc, so this should now be rare. New specs
  register themselves from the trait; there is no list to update.
- **Inline + opaque investigations need clean builds.** Incremental compiles move error positions
  between runs when opaque-type members are involved; trust only `./mill clean` + full compile and
  isolate with minimal probe files (see the `xl-scala-style` skill).
- **Never override `resources` in a nested `object test`** of a dash-named module: task resolution
  fails ("Could not detect the parent class of task ...resources.super...") regardless of the body.
  Share files across test modules with `moduleDeps ++= Seq(other.test)` instead
  (`xl-cats-effect/package.mill` documents this).
- **Adding compiler plugins**: `scalacPluginMvnDeps` must append to `super`, or the existing
  plugin set (WartRemover) is dropped.
- **Bumping Scala** requires a WartRemover build for the exact new version
  (`org.wartremover:::wartremover`, full cross); check Maven Central first. The examples pin
  `//> using scala` in `examples/project.scala` and the skill docs, and must move together.
- **`out/` is never cached across Mill versions** (stale artifacts break builds); `./mill clean`
  when a build behaves inexplicably after a Mill or Scala change.
- **Containers / one-shot CI**: add `--no-daemon` (or `-i`) so no background server outlives the
  step. The rehearsal script does this.
- **Warnings are not errors**: `-deprecation -feature` are on, `-Werror` is not. WartRemover
  Tier 1 warts are the only compile-time failures beyond type errors (see `xl-scala-style`).
