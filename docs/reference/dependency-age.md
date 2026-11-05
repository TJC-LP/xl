# Dependency age and updates

`python3 scripts/dependency_age.py` checks the direct Maven declarations in tracked
Mill files, the Scala compiler, and the Mill-versioned JMH build plugin. CI runs
its offline regression tests and the live check. It uses Maven Central POM
`Last-Modified` timestamps, not a vulnerability database or proof of maintenance.
A missing/invalid timestamp, missing artifact or failed network request fails the
check. A pin older than 365 days needs a nonempty `// pinned: reason` comment on
its declaration line. Keep reasons specific and revisit them when dependencies
change. A latest stable version can itself be older than 365 days.

The checker supports the repository's literal `mvn"group:artifact:version"`, `::`
and `:::` syntax and rejects unsupported interpolated Maven literals. It does not
interpret arbitrary Scala programs. Changes to dependency declaration style or
build-header dependencies require extending and reviewing the checker rather
than moving declarations outside its inventory. It checks direct pins, not the
entire resolved transitive graph; the dependency lock work in #833 covers selected
resolved artifact bytes separately.

`renovate.json` uses Renovate's supported `custom.regex` manager for these Mill
literals, plus `github-actions` and explicitly enabled `pre-commit`. There is no
native `mill` manager in the verified Renovate 44.221.1 source. Custom extraction
uses Scala binary suffix `_3` and compiler-plugin suffix `_3.9.0`; when changing
Scala, review those templates together. Scala updates are restricted to 3.9 LTS.
Mill version updates are proposed from its official releases. Automerge is off.

Retain SHA-pinned Actions and frozen hook revisions. Hooks must use the
`rev: <full SHA> # frozen: vX.Y.Z` convention for Renovate to preserve and update
frozen revisions. Integrate the hook-pinning contribution (#835) before enabling
the updater. A Mill update also requires refreshing committed launcher digests
(#837). Dependency changes require regenerating/reviewing the lock (#833), then
running the repository's validation gates. Renovate does not silently regenerate
these security records in this configuration.

The major migrations for OkHttp, scala-csv, scalacheck-effect and benchmark POI
remain separate #839 follow-ups with explicit current pin reasons. The config can
propose their PRs for review; it does not approve their compatibility or merge.

Committing configuration does not install an app. Repository administrators must
authorize and enable Renovate, then confirm an actual update PR and CI run. Until
that happens, the first-update acceptance criterion in #839 remains open. No
hosted updater run is established by local schema or extraction tests.
