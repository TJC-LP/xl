# Workflow rules (always loaded)

Process conventions for this repo that the code cannot express. Build/test detail lives in the
`mill-build` skill, testing patterns in `xl-testing`, Scala idioms in `xl-scala-style`.

- **Gates before a PR that touches the public surface** (run all four, with `set -o pipefail` so a
  piped `tail` cannot hide a failure): `./mill __.compile`, `./mill __.test`,
  `scripts/test-examples.sh`, `scripts/verify-skill-snippets.sh --local`. Docs-only changes need
  none of them.
- **Formatting is checked over ALL sources**: CI runs
  `./mill mill.scalalib.scalafmt.ScalafmtModule/checkFormatAll __.sources`. The shorter
  `./mill __.reformat` / `__.checkFormat` skip test sources; format with `reformatAll __.sources`
  before committing.
- **Stage paths explicitly. Never `git add -A` or `git add .`**: checkouts carry untracked personal
  files (an `AGENTS.md` from Codex, scratch `.xlsx` output, `.claude/settings.local.json`).
- **Commits**: conventional subject (`feat(core): …`, `fix(cli): …`, `perf(evaluator): …`,
  `docs: …`, `chore(release): …`), body says why. Reference issues with `Refs #n` in commit bodies;
  closing keywords belong in the PR body, one keyword per issue (`Closes #1, closes #2` — GitHub
  honors only the first reference after a keyword).
- **Merges are squash merges** whose subject is the PR title followed by ` (#PR)`.
- **CHANGELOG.md**: add entries under `## [Unreleased]` in the PR that makes the change;
  `/release-prep` moves them under the version heading at release time.
- **Long builds need explicit timeouts**: `./mill __.test` takes minutes and a cold
  `./mill __.compile` in a fresh checkout downloads dependencies first. Pass a 600000 ms timeout
  to the Bash tool for either; do not conclude a hang before that elapses.
- **Never rewrite shared history**: no force-push or rebase of a pushed branch without the user
  saying so.
- **Cloud sessions** (`CLAUDE_CODE_REMOTE=true`): no `gtr`, no GraalVM, no personal memory. The
  sandbox clone is already isolated, so skip worktrees; use `make install-jar` instead of
  `make install`. `scripts/remote-setup.sh` has already provisioned JDK 25 and scala-cli when the
  session started; its summary is in your context.
