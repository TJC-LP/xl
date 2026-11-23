---
name: xl-verify
description: Verify an xl change the way a user meets it - build the branch's real `xl` CLI and main's, drive the same made-up workbook flow through both, and check the output files, XML, values and timings against what the change claims. Setup, drive and check stages, a diff-triggered checklist from past bugs, and an evidence table for the PR's "Verified" section. Use after the gates and before requesting review on any PR that changes behavior or performance, when asked to "verify #N" or "/xl-verify", or to reproduce a reported bug at the CLI.
---

# xl-verify: check the change where a user meets it

Unit specs prove the code does what the test author thought of. This skill proves that the
**built `xl` CLI** behaves the way the PR says, on a workbook a user could have, with nothing else
in the file disturbed. It does not replace the gates in `.claude/rules/workflow.md`; run it after
them and before asking for review. Docs-only, build-only and test-only changes skip it.

**Made-up data only:** every workbook, sheet name, URL and number is invented (`Summary`, `Data`,
`example.com`, `example.net`, seeded random values).

## 0. Decide what to verify

Read the issue, the PR body and `git diff origin/main...HEAD --stat`. Write down, in
`$RUN/claims.md`:

1. **The claim:** what a user can now do, or no longer sees, in one or two sentences.
2. **The surface:** the exact `xl` commands a user would run to meet it (read each `--help`).
   For a library-only change, use the CLI command that reaches that code path. If none does,
   say so and verify with a scala-cli script against the public API instead.
3. **The flows:** the happy path, the edge cases from the issue, and every item from
   `references/checklist.md` whose trigger matches the diff. Aim for 3–10 flows, not 1.
4. **What must not change:** the parts of the file and the outputs the PR should leave alone.

## 1. Setup

From the repository root:

```bash
RUN=$(mktemp -d "${TMPDIR:-/tmp}/xl-verify.XXXX")       # scratch, never inside the repo
S=.claude/skills/xl-verify/scripts
bash $S/build-xl.sh "$PWD" "$RUN/xl-branch"              # this checkout's head
git worktree add --detach "$RUN/main-src" origin/main    # main, read-only
bash $S/build-xl.sh "$RUN/main-src" "$RUN/xl-main"
```

- **Stacked PR:** build the base from the parent PR's head instead of `origin/main`.
- **Perf PR, or a change downstream users pick up through a release:** also verify against the
  latest release as a third side, using an installed `xl` (`xl --version`) or a launcher built
  from its tag.
- **Fixtures:** put them in `$RUN/flows/fixtures/`. Build them **with xl itself** (`xl new`,
  `put`, `putf`, `batch`), so they are files a user could make.
  - For shapes xl cannot write (range hyperlinks, shared formulas from other tools, legacy
    parts), take one from `xl-ooxml/test/resources/fixtures/` (read its `PROVENANCE.md`). If none
    fits, add the minimal XML by hand and say so in the report.
  - For perf books, use the made-up generators in the `xl-evaluator` perf module
    (`LoadsBookMain shape=...`), at a size where the base takes more than 2 seconds.
- **Cleanup:** remove the main worktree afterwards (`git worktree remove --force "$RUN/main-src"`).

## 2. Drive

One bash script per flow, in `$RUN/flows/<name>.sh`. Inside a flow:

- **Use only `"$XL" ...` commands**, copying inputs from `$IN` into `$OUT` first. No Scala, no
  unzip-and-edit, no `./mill`: a flow is what a user types.
- **Save every output a user would look at:** the written `.xlsx`; the `--json` output of
  `view`, `cell`, `eval` and `audit`; `--strict lint`; and stderr on failures.
- **Re-read after every write** with the same launcher (`view --json`, `--strict lint`). A file
  xl cannot read back is a failure even if the write "succeeded".

Example flow:

```bash
# a range link beside a cell link survives an unrelated edit (#912)
cp "$IN/cell-and-range.xlsx" "$OUT/in.xlsx"
"$XL" -f "$OUT/in.xlsx" -s Summary -o "$OUT/out.xlsx" put A1 edit > "$OUT/put.txt"
"$XL" -f "$OUT/out.xlsx" -s Summary --json cell B7 > "$OUT/cell-B7.json"
"$XL" -f "$OUT/out.xlsx" --strict lint > "$OUT/lint.txt" 2>&1 || true
```

Run each flow through both sides:

```bash
bash $S/compare.sh "$RUN/flows/<name>.sh" "$RUN/xl-main" "$RUN/xl-branch" "$RUN/out/<name>" [runs]
```

- **`runs`:** leave it at 1, except for timing. For timing use 5 and report the median, on the
  same machine with nothing else heavy running.
- **What compare.sh writes:** it unpacks every `.xlsx`, pretty-prints the XML, and replaces each
  side's own output path with `<OUT>`. It then writes `diff.txt`, `changed.txt` and `times.tsv`.
  It does not judge.

## 3. Check

Act as a separate checker, and assume the driver might have cheated or stopped early. For each
flow:

1. **Did it run the real thing?**
   - `build-xl.sh` printed the expected sha and version for both launchers.
   - The flow used only `$XL`.
   - The exit codes in `times.tsv` are what a user would expect.
2. **Main shows the problem; the branch fixes it.** For a bug fix, the failing behavior must
   appear on main. If main already behaves correctly, stop and report: the issue may already be
   fixed, or misread.
3. **Only the claimed difference changed.** Read `diff.txt` line by line; every hunk must be
   explained by the claim. Anything else is a finding until explained: a relationship id
   renumbered, a style index moved, a part dropped, a cached value changed.
4. **The file is sound.**
   - `--strict lint` is clean on both sides, or reports the same on both.
   - The branch's output re-reads.
   - If LibreOffice is available (`soffice --headless --convert-to csv`), the file opens and its
     values agree with xl's for the cells the flow touched.
   - Opening in Microsoft Excel is a manual check: list the files worth opening rather than
     automating Excel.
5. **Perf claims:** report the median speedup and the spread, base vs branch, and confirm the
   outputs are identical. Identical results matter more than the speedup.
6. **Small defects:** note anything odd, even if the flow passed: a misleading message, a warning
   on stderr, a needless rewrite of an untouched part.

## 4. Report

Write `$RUN/report.md` and show it to the user before posting anything:

```markdown
### Verified at the CLI (xl-verify) on <branch sha> vs main <sha>

| Flow | Main | Branch | Verdict |
|---|---|---|---|
| range link survives an edit next to a cell link | link dropped | kept, rels intact | fixed |
| untouched parts | - | byte-identical except sheet1.xml hyperlinks | as claimed |

Timings (median of 5): recalc 12.4 s -> 3.1 s; outputs identical.
Findings: <none | numbered list with flow, evidence path, why it matters>
Not verified: <what and why, e.g. Excel open check left to the reviewer>
```

- **Findings in the PR's scope:** fix them in the PR, then re-run only the affected flows.
- **Findings outside its scope:** file a GitHub issue with a made-up repro.
- **The table** goes in the PR's "Verified" section. Flows and fixtures stay out of the repo
  unless the PR also turns them into specs.

## 5. Keep the verifier honest

- **Every escaped bug teaches the checklist.** When a bug that this skill should have caught
  reaches main or a release, add a trigger and a flow to `references/checklist.md`, with the
  issue number.
- **Seeded-bug check, about monthly:**
  1. On a scratch branch, revert the fix of a recently merged PR, keeping its tests out.
  2. Run this skill as if it were a new PR.
  3. Record in `references/honesty-log.md` whether it caught the bug. A miss means the skill or
     the checklist needs work.
