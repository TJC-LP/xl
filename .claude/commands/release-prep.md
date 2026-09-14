# Release Preparation

Prepare and cut release version: $ARGUMENTS

If `$ARGUMENTS` is empty, do not guess silently: read the fallback in `build.mill`
(`sys.env.getOrElse("PUBLISH_VERSION", "<current>")`) — that is the LAST release — and propose the
next version from the `## [Unreleased]` section of `CHANGELOG.md` (any `**Breaking` entry → bump
the minor while we are pre-1.0; otherwise the patch). State the version you are using in your
first message and proceed. Below, `NEW` is that version and `OLD` the `build.mill` fallback.

The work order matters: **branch → bump → verify → commit → PR → squash-merge → tag the merge
commit → push the tag**. `main` only accepts pull requests (repository rule GH013) and merges are
squash merges, so the release commit's SHA changes on merge and the tag can only be created
afterwards. Every release tag sits on `main`.

## 1. Branch

```bash
git checkout main && git pull --ff-only origin main
git checkout -b release-NEW
```

## 2. Bump every version location

> **Line numbers drift between releases.** The authoritative list is the grep, not this file:
> ```bash
> grep -rn "OLD" --include="*.scala" --include="*.mill" --include="*.json" --include="*.md" . \
>   | grep -v "out/" | grep -v ".scala-build" | grep -v CHANGELOG.md | grep -v docs/RELEASING.md
> ```
> Two kinds of hit come back. **Pins** (dependency coordinates, version fields, the skill floor)
> all move to `NEW`. **History** ("since 0.22.0", "0.20.0–0.22.0 printed bare arrays", "the rule
> before 0.22.0") stays as it is — it describes a past release. Read each hit before editing.

The dependency coordinates are one find-and-replace across the tree (excluding `CHANGELOG.md` and
`docs/RELEASING.md`): `com.tjclp::xl:OLD`, `com.tjclp::xl-core:OLD` … `xl-evaluator:OLD`, and
sbt's `"com.tjclp" %% "xl" % "OLD"`. That covers `README.md`, `docs/QUICK-START.md`,
`docs/reference/scripting.md`, `examples/project.scala`, `examples/README.md`,
`plugin/skills/xl-scripting/SKILL.md`, `plugin/skills/xl-scripting/reference/RECIPES.md` (its
recipe headers are intentionally byte-identical), `plugin/skills/xl-scripting/reference/INSTALL.md`
(inside a `printf`) and the scaladoc of `xl/src/com/tjclp/xl/scripting.scala`. The release
workflow **fails the release** if the xl-scripting pins do not match the tag.

Then the individual fields:

1. **`build.mill`** — `val version: String = sys.env.getOrElse("PUBLISH_VERSION", "NEW")`.
2. **`xl-core/src/com/tjclp/xl/workbooks/WorkbookMetadata.scala`** — `appVersion: Option[String] = Some("NEW")`.
3. **`plugin/.claude-plugin/plugin.json`** — `"version": "NEW"`. **Do not skip** — it drifted to
   0.7.0 once because it was missing from this list; CI's `plugin-version` gate now compares it
   with `build.mill`.
4. **`xl-cli/src/com/tjclp/xl/cli/contract/Render.scala`** — the example envelope in the scaladoc
   (`"version": "OLD"`), so the documented sample matches what the binary prints.
5. **`plugin/skills/xl-cli/SKILL.md`** — the `**Requires xl >= …**` line (and the
   `xl --version must print …` bullet under Environment) becomes `NEW`. If the skill carries the
   `<!-- unreleased-contract -->` marker with its blockquote ("X has not shipped yet … install from
   source … the last release lacks …"), **remove both** now: this release ships that contract. A
   marker left behind, or a `Requires` version that differs from `plugin.json` without the marker,
   fails CI's `plugin-version` gate. The install snippets auto-detect the latest release and need
   no edit.
6. **`CHANGELOG.md`** — insert `## [NEW] - <YYYY-MM-DD>` directly under `## [Unreleased]` so every
   entry moves under the new heading and `[Unreleased]` is left empty. The tag message is
   extracted from this heading (step 7): a missing heading gives an empty GitHub release. If the
   release deserves a one-paragraph summary above its `### Added`, write it here.
7. **`docs/STATUS.md`** — `**Last Updated**: <date> (NEW)`, a `**New in NEW** (<date>)` paragraph
   with the release's headline bullets above the previous release's paragraph, and the test-count
   line/table if they are stale (`./mill __.test` in step 3 gives the numbers).
8. **`docs/plan/roadmap.md`** — `**Current Version**: **NEW** (…)` and the release's section
   heading `(Unreleased)` → `(Released <date>)`.

Skip **`docs/RELEASING.md`** (example version strings) and the historical headings in
`CHANGELOG.md`.

## 3. Verify (all on the release branch; every `./mill` call gets a 600000 ms timeout)

1. `./mill __.compile`
2. `./mill __.test` — note the per-module counts for STATUS/CLAUDE.md/testing-guide if they moved
3. `./mill mill.scalalib.scalafmt.ScalafmtModule/checkFormatAll __.sources`
4. `./scripts/test-examples.sh` — version-drift guard plus the examples against the local build
5. `./scripts/verify-skill-snippets.sh --local` — the xl-scripting snippets compile against `NEW`
6. No SNAPSHOT references:
   ```bash
   grep -r "SNAPSHOT" --include="*.scala" --include="*.mill" --include="*.md" . | grep -v RELEASING.md | grep -v ".scala-build" | grep -v "out/" | grep -v ".claude/commands"
   ```
7. **Doc drift into the skills.** Every public API the release's CHANGELOG names must appear in
   `plugin/skills/xl-scripting/reference/API.md` and `docs/reference/scripting.md`; drift here
   ships to every skill user. Make it mechanical: pull the backticked identifiers out of the new
   heading's `### Added` / `### Changed` entries (`withDefinedName`, `StructuralCachePolicy`,
   `Fill.pattern`, `TotalsRowFunction`, …) and `grep -c` each in both files; anything at 0 that a
   script author would call gets a table row or a sentence — a row in the right table with the
   version and issue link is the house style (`| \`wb.withDefinedName(…)\` | \`Workbook\` | 0.23.0
   ([#538](…)): … |`). 0.23.0 found five such gaps at this step.
8. The `plugin-version` gate's inputs agree: `jq -r .version plugin/.claude-plugin/plugin.json`,
   `grep -oE 'Requires xl >= [0-9.]+' plugin/skills/xl-cli/SKILL.md`, and no
   `unreleased-contract` marker.

## 4. Commit

Stage the touched paths explicitly (never `git add -A`) and commit:

```
chore(release): Bump version to NEW
```

with a body saying what moved (pins, CHANGELOG heading, STATUS/roadmap, any skill-doc additions)
and which gates ran.

## 5. Pull request and squash merge

```bash
git push -u origin release-NEW
gh pr create --base main --head release-NEW --title "chore(release): Bump version to NEW" --body "<what moved, gates>"
gh pr checks <PR> --watch --fail-fast
gh pr merge <PR> --squash --delete-branch --subject "chore(release): Bump version to NEW (#<PR>)"
```

Do **not** `git push origin main` (GH013 rejects it) and do **not** create the tag yet.

## 6. Realign `main`

```bash
git fetch origin && git checkout -B main origin/main   # the squash commit; your local release commit is gone from main by design
git log --oneline -1                                   # chore(release): Bump version to NEW (#<PR>)
```

## 7. Tag the merge commit (annotated, release notes from the CHANGELOG heading)

```bash
# Extract the notes (literal version — a $VERSION inside the awk pattern trips zsh)
awk '/^## \[NEW\]/{flag=1; next} /^## \[/{flag=0} flag' CHANGELOG.md | sed '/^$/d' > /tmp/notes-NEW.md
wc -l /tmp/notes-NEW.md                                # non-zero, or the heading is wrong
{ printf 'xl NEW — <one-line theme>\n\n'; cat /tmp/notes-NEW.md; } > /tmp/tag-NEW.md

git tag -a "vNEW" -F /tmp/tag-NEW.md
git cat-file -t "vNEW"                                 # must print "tag" (annotated), never "commit"
git branch -r --contains "vNEW"                        # must list origin/main
git push origin "vNEW"
```

Never `git tag vNEW` without `-a`: a lightweight tag has no message and the GitHub release shows
the commit message instead of the notes.

## 8. Watch the release

```bash
gh run list --limit 3                                  # "Release" on vNEW should be in_progress
gh run watch <run-id> --exit-status --interval 30      # ~25 min: five native builds, then publish
gh release view vNEW --json name,isDraft,assets --jq '{name,isDraft,assets:[.assets[].name]}'
```

Expected assets: `xl-NEW-{darwin-amd64,darwin-arm64,linux-amd64,linux-arm64}`,
`xl-NEW-windows-amd64.exe`, `xl-cli-NEW.tar.gz`, `xl-scripting-skill-NEW.zip`, `xl-skill-NEW.zip`.
Locally, `make install` rebuilds the native CLI at the released version.

## Recovery: a tag was pushed before the merge

The tag points at a commit that never reaches `main`, and the Release workflow is already
running from it. Before it publishes to Maven Central (irreversible — a second publish of the same
version is rejected, so the real release would then fail):

```bash
gh run list --limit 3                                  # find the Release run on vNEW
gh run cancel <run-id>
git push origin :refs/tags/vNEW                        # delete the remote tag
git tag -d vNEW                                        # and the local one
```

Then continue from step 5. The cancelled run's "Publish to Maven Central" job shows as `fail`
in `gh pr checks` afterwards; `gh run view <run-id> --json jobs` must say `cancelled` for it, with
no `startedAt` before the cancel. (0.23.0 went through exactly this.)
