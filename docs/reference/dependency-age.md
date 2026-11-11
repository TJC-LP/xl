# Dependency age and updates

`python3 scripts/dependency_age.py` checks the direct Maven declarations in tracked
Mill files, the Scala compiler, and the Mill-versioned JMH build plugin. CI runs
its offline regression tests and the live check in their own `dependency-age` job,
so a Maven Central outage cannot skip the release-version gates. It uses Maven Central POM
`Last-Modified` timestamps, not a vulnerability database or proof of maintenance.
A missing/invalid timestamp, missing artifact or failed network request fails the
check. A pin older than 365 days needs a dated reason on its declaration line:
`// pinned until YYYY-MM-DD: reason`. The reason counts through the review date and
stops counting after it, so every exemption is revisited at least yearly; a date
more than 366 days ahead, an undated `// pinned:` comment, a blank reason or an
invalid date fails the check. A latest stable version can itself be older than 365
days; record that as the reason and re-check it by the review date.

The Mill version in `.mill-version` (checked as `mill-contrib-jmh_3`) has no
declaration line to annotate, so its reason goes on one dedicated line in
`build.mill`: `// mill-version pinned until YYYY-MM-DD: <reason>`. Add it only when Mill passes
365 days and an upgrade has to wait; more than one such line fails the check.

The checker supports the repository's literal `mvn"group:artifact:version"`, `::`
and `:::` syntax and rejects unsupported interpolated Maven literals. It does not
interpret arbitrary Scala programs. Block comments are found before line comments
and outside `"..."` literals, so a URL in a doc comment is safe; nested block
comments and a dependency inside or beside a block comment fail closed. Changes to dependency declaration style or
build-header dependencies require extending and reviewing the checker rather
than moving declarations outside its inventory. It checks direct pins, not the
entire resolved transitive graph; the dependency lock work in #833 covers selected
resolved artifact bytes separately.

`renovate.json` uses Renovate's supported `custom.regex` manager for these Mill
literals, plus `github-actions` and explicitly enabled `pre-commit`. There is no
native `mill` manager in the verified Renovate 44.221.1 source. Custom extraction
uses Scala binary suffix `_3`. The full-cross (`:::`) compiler-plugin suffix is read
from `build.mill`'s `scala3Version` with a `combination` match, which supports one
full-cross dependency, in `build.mill` only; the age check enforces that. The Scala
compiler and the wartremover plugin update in one "Scala 3.9 toolchain" PR. Scala
updates are restricted to 3.9 LTS.
Mill version updates are proposed from its official releases. Automerge is off.

Retain SHA-pinned Actions and frozen hook revisions. Hooks must use the
`rev: <full SHA> # frozen: vX.Y.Z` convention for Renovate to preserve and update
frozen revisions (#835). A Mill update also requires refreshing the committed
launcher digests (#837). Dependency changes require regenerating/reviewing the lock (#833), then
running the repository's validation gates. Renovate does not silently regenerate
these security records in this configuration.

The major migrations for OkHttp, scala-csv, scalacheck-effect and benchmark POI
remain separate #839 follow-ups with explicit current pin reasons. The config can
propose their PRs for review; it does not approve their compatibility or merge.

Committing configuration does not install an app. Repository administrators must
authorize and enable Renovate, then confirm an actual update PR and CI run. Until
that happens, the first-update acceptance criterion in #839 remains open. No
hosted updater run is established by local schema or extraction tests.

### Benchmark JMH setting

The inventory also checks both `org.openjdk.jmh:jmh-core` and
`org.openjdk.jmh:jmh-generator-bytecode`. Mill 1.1.5 inherits the generator version
from `jmhCoreVersion`. The checker recognizes the literal declaration as the first
active statement in the existing `xl-benchmarks/package.mill` package object,
followed by the existing `moduleDeps` declaration (or the closing module brace);
expressions, duplicate declarations, explicit generator overrides in this file,
and mixed block-comment/code lines require a policy update. Standalone block
comments and line comments are ignored. This is a deliberately bounded check of
this benchmark module, not a Scala parser or a search for other JMH modules.

One Renovate manager updates the shared core version; the generator follows it,
and the age check validates both artifact timestamps. The existing 1.35 harness
is retained for benchmark baseline comparability, with an inline age exemption.
This does not assert that newer JMH is incompatible or unsafe. A future upgrade
should refresh the benchmark baselines. The reason exempts age only; missing,
future, or unavailable Maven timestamps still fail.
