export const meta = {
  name: 'issue-wave',
  description: 'Parameterized backlog wave: baseline gate, optional design panel, worktree-isolated TDD implementers, pipelined adversarial review with one rework round',
  whenToUse: 'Run a wave of GitHub issues as parallel worktree clusters. Pass {wave, branch, clusters:[{key, issues, modules, brief, serialNote?}], designFirst:[{issue, question}]} as args. Integration/PR/merge stays in the main loop.',
  phases: [
    { title: 'Baseline', detail: 'compile+test green on the wave branch, capture ground truth' },
    { title: 'Design', detail: 'judge panel for designFirst issues (skipped when none)' },
    { title: 'Implement', detail: 'one TDD agent per cluster in an isolated worktree' },
    { title: 'Review', detail: 'adversarial reviewer per cluster, one rework round' },
  ],
}

const ARGS = typeof args === 'string' ? JSON.parse(args) : (args || {})
const REPO = ARGS.repo || '/Users/rcaputo3/git/xl'
const WAVE = ARGS.wave || 0
const BRANCH = ARGS.branch || `wave-${WAVE}`
const CLUSTERS = ARGS.clusters || []
const DESIGN_FIRST = ARGS.designFirst || []
if (!CLUSTERS.length) throw new Error('issue-wave requires args.clusters')

const BASELINE_SCHEMA = {
  type: 'object',
  required: ['green', 'testCount', 'preExisting', 'notes'],
  properties: {
    green: { type: 'boolean' },
    testCount: { type: 'integer' },
    preExisting: { type: 'array', items: { type: 'string' } },
    notes: { type: 'string' },
  },
}

const DESIGN_SCHEMA = {
  type: 'object',
  required: ['design', 'tradeoffs'],
  properties: { design: { type: 'string' }, tradeoffs: { type: 'string' } },
}

const JUDGE_SCHEMA = {
  type: 'object',
  required: ['brief', 'rationale'],
  properties: { brief: { type: 'string' }, rationale: { type: 'string' } },
}

const IMPL_SCHEMA = {
  type: 'object',
  required: ['outcome', 'sha', 'branch', 'worktreePath', 'summary', 'filesChanged', 'testsAdded', 'gateEvidence', 'gaps'],
  properties: {
    outcome: { type: 'string', enum: ['fixed', 'not-reproducible', 'partial', 'blocked'] },
    sha: { type: 'string' },
    branch: { type: 'string' },
    worktreePath: { type: 'string' },
    summary: { type: 'string' },
    filesChanged: { type: 'array', items: { type: 'string' } },
    testsAdded: { type: 'array', items: { type: 'string' } },
    gateEvidence: { type: 'string' },
    gaps: { type: 'array', items: { type: 'string' } },
  },
}

const REVIEW_SCHEMA = {
  type: 'object',
  required: ['verdict', 'findings', 'empiricalEvidence'],
  properties: {
    verdict: { type: 'string', enum: ['approve', 'rework', 'reject'] },
    findings: { type: 'array', items: { type: 'string' } },
    empiricalEvidence: { type: 'string' },
    reworkInstructions: { type: 'string' },
  },
}

const HOUSE_RULES = `HOUSE RULES (xl repo, ${REPO}):
- Purity charter: no null, no thrown exceptions in pure code, total functions returning Either/Option (XLResult), opaque types, enums with derives CanEqual. WartRemover enforces; .head/.tail and Option.get are warnings, var/while acceptable only in tests/macros.
- NEVER re-inline extension methods or factories on opaque types (compiler landmines — see NOTE in xl-core addressing/ARef.scala). No default params on exported extension methods (use @targetName overloads).
- Do NOT edit CHANGELOG.md or docs/plan/roadmap.md (the integrator owns them — avoids 6-way conflicts).
- Style: ./mill <module>.reformat before committing; match surrounding code idiom; comments only for constraints code cannot express.
- Commit with --no-verify (you run the format/test gates explicitly instead of the slow pre-commit hook), conventional message ending with "Refs #<issue>" per issue (NOT "Closes" — the wave PR owns closing).`

phase('Baseline')
log(`Wave ${WAVE}: baseline gate on branch ${BRANCH}...`)
const baseline = await agent(`You are the baseline agent for wave ${WAVE} in ${REPO} (branch ${BRANCH} checked out). READ-ONLY plus build commands; no file edits, no commits.
1. Confirm git branch --show-current prints ${BRANCH} and git status is clean.
2. Run ./mill __.compile (timeout 600000 ms). Then ./mill __.test (timeout 600000 ms).
3. Report: green (both succeeded), testCount (total test results from the run output), preExisting (any failures/warnings that exist BEFORE this wave touches anything — exact test names), notes (anything implementers should know, e.g. flaky tests, long-running suites).
If the suite is red, report green=false with the failing names — the wave will abort rather than build on a broken base.`, { label: 'baseline', phase: 'Baseline', schema: BASELINE_SCHEMA })

if (!baseline || !baseline.green) {
  return { aborted: 'baseline-red', baseline }
}
log(`Baseline green: ${baseline.testCount} tests. ${baseline.preExisting.length} pre-existing notes.`)

const designBriefs = {}
if (DESIGN_FIRST.length) {
  phase('Design')
  log(`Design panel for ${DESIGN_FIRST.length} issue(s)...`)
  const LENSES = ['Excel-parity-first: match Excel observable behavior exactly, cite real Excel semantics', 'purity-first: totality, determinism, law-governed — the design must not weaken the purity charter', 'minimal-diff-first: smallest change that ships, defer everything deferrable, name what is deferred']
  for (const df of DESIGN_FIRST) {
    const designs = await parallel(LENSES.map((lens, i) => () =>
      agent(`You are design panelist ${i + 1} for xl issue #${df.issue} in ${REPO} (read-only — explore code, read the issue with gh issue view ${df.issue}, but edit nothing).
Your lens: ${lens}.
Design question: ${df.question}
Produce a concrete design: API signatures, file-level change list, semantics decisions, test strategy, what is explicitly out of scope. Ground every claim in the actual code (cite paths).`, { label: `design:${df.issue}:${i + 1}`, phase: 'Design', schema: DESIGN_SCHEMA })))
    const valid = designs.filter(Boolean)
    const judged = await agent(`You are the adversarial design judge for xl issue #${df.issue}. Three panelists designed independently under different lenses. Score each for correctness against the actual codebase (verify claims — read the code), Excel parity, purity-charter fit, and implementability. Then synthesize THE brief the implementer will follow: take the winning skeleton, graft superior ideas from the others, list explicit decisions (no open questions), file-level change list, and test plan.
PANELIST DESIGNS:
${valid.map((d, i) => `--- Design ${i + 1} ---\n${d.design}\nTradeoffs: ${d.tradeoffs}`).join('\n')}`, { label: `judge:${df.issue}`, phase: 'Design', schema: JUDGE_SCHEMA })
    if (judged) designBriefs[df.issue] = judged.brief
  }
}

phase('Implement')
log(`Fanning out ${CLUSTERS.length} worktree implementers...`)

const implPrompt = (c) => {
  const briefExtras = (c.issues || []).filter(n => designBriefs[n]).map(n => `\nDESIGN BRIEF for #${n} (binding — follow it):\n${designBriefs[n]}`).join('')
  return `You are the implementation agent for cluster "${c.key}" of wave ${WAVE} in the xl repo. You are running in an ISOLATED GIT WORKTREE — confirm with pwd and git rev-parse --show-toplevel (it will NOT be ${REPO}; that is correct). Base branch: ${BRANCH}. Modules in scope: ${c.modules}.

ISSUES TO RESOLVE (read each fully first: gh issue view <n> --json title,body,comments): ${(c.issues || []).map(n => '#' + n).join(', ')}
${c.serialNote ? `SERIALIZATION ORDER (mandatory): ${c.serialNote}` : ''}

CLUSTER BRIEF (exploration-verified seeds — re-verify line numbers before editing, code may have drifted):
${c.brief}${briefExtras}

${HOUSE_RULES}

METHOD (mandatory):
1. TDD: for each issue, FIRST write the failing test that pins the bug/feature from the issue's repro, run it, SEE IT FAIL, then implement until green. Tests live in the module's existing spec files where the brief names them.
2. Touch only files within this cluster's footprint. If a fix genuinely requires a file another cluster owns, STOP that part and report it as a gap instead.
3. Local gates before declaring done: ./mill <module>.test for every module you touched (timeout 600000), ./mill __.checkFormat after ./mill <module>.reformat. All green.
4. Commit (git commit --no-verify) with a conventional message; one commit per issue is ideal, a single clean cluster commit acceptable. Message body: "Refs #<n>" lines.
5. Report: outcome (fixed | not-reproducible [only if the brief says verify-first and you proved it] | partial | blocked), sha (git rev-parse HEAD), branch (git branch --show-current), worktreePath (pwd), summary, filesChanged, testsAdded (test names), gateEvidence (the actual gate commands + tail of their output), gaps (anything discovered worth filing as a new issue — be specific).
Your final structured output is consumed by an adversarial reviewer who will try to refute your work — include enough evidence that an honest reviewer can retrace every claim.`
}

const reviewPrompt = (c, impl, isReReview) => `You are the ADVERSARIAL reviewer for cluster "${c.key}" of wave ${WAVE} (xl repo). ${isReReview ? 'This is a RE-REVIEW after one rework round — previous findings should now be addressed.' : ''}Default stance: the implementation is wrong until you fail to refute it. Empirical evidence only — claims you did not execute do not count.

THE IMPLEMENTER REPORTED: ${JSON.stringify(impl, null, 2)}

ISSUES: ${(c.issues || []).map(n => '#' + n).join(', ')} (read them: gh issue view <n>)
CLUSTER BRIEF (what was supposed to happen):
${c.brief}

PROTOCOL:
1. Get the code: work in the implementer's worktree if it still exists (cd ${impl.worktreePath}); otherwise create your own: git -C ${REPO} worktree add /tmp/review-${c.key}-w${WAVE} ${impl.sha} and work there (remove it when done).
2. Refutation pass: run the issue's ORIGINAL repro — it must now behave correctly. Then prove the new tests pin the bug: temporarily restore the pre-fix source (git checkout ${BRANCH} -- <changed src files>, NOT the test files), run the new tests, they MUST FAIL; then restore (git checkout ${impl.sha} -- <files>). A test that passes without the fix is a rubber stamp — that is a rework finding.
3. Break it: probe edge cases the issue implies but the tests skip (empty/boundary/unicode/negative/huge inputs; for parser work, round-trip law parse-print-parse; for OOXML work, write-read round-trip AND streaming-vs-in-memory parity where both paths exist; for refactors, behavioral equivalence on the old test suite).
4. Run the module test suite yourself (./mill <module>.test) — do not trust gateEvidence.
5. Scope check: git show --stat ${impl.sha} — flag files outside the cluster footprint, CHANGELOG/roadmap edits (forbidden), or unrelated drive-by changes.
VERDICT: approve (everything held) | rework (fixable findings — give precise reworkInstructions: file, what is wrong, what correct looks like) | reject (fundamentally wrong approach; explain). Report empiricalEvidence: the commands you ran and their actual outcomes. Do NOT edit the implementation yourself.`

const reworkPrompt = (c, impl, review) => `You are the rework agent for cluster "${c.key}" of wave ${WAVE} (xl repo). An adversarial reviewer found concrete problems in the implementation. Fix exactly these findings — no scope creep.

WORKTREE: cd ${impl.worktreePath} — if it no longer exists, recreate it: git -C ${REPO} worktree add /tmp/rework-${c.key}-w${WAVE} ${impl.branch} && cd there.
PRIOR IMPLEMENTATION REPORT: ${JSON.stringify(impl, null, 2)}
REVIEWER FINDINGS (fix all of these):
${(review.findings || []).map(f => '- ' + f).join('\n')}
REWORK INSTRUCTIONS: ${review.reworkInstructions || '(see findings)'}

${HOUSE_RULES}

Fix, keep TDD discipline (failing test first for any finding that lacks one), rerun local gates (./mill <module>.test + ./mill __.checkFormat after reformat), commit on top (--no-verify), and report the same structured output as an implementer (outcome/sha/branch/worktreePath/summary/filesChanged/testsAdded/gateEvidence/gaps). The reviewer re-reviews after you.`

const results = await pipeline(
  CLUSTERS,
  c => agent(implPrompt(c), { label: `impl:${c.key}`, phase: 'Implement', isolation: 'worktree', schema: IMPL_SCHEMA }),
  async (impl, c) => {
    if (!impl) return { cluster: c.key, status: 'impl-failed' }
    if (impl.outcome === 'blocked') return { cluster: c.key, status: 'blocked', impl }
    let review = await agent(reviewPrompt(c, impl, false), { label: `review:${c.key}`, phase: 'Review', schema: REVIEW_SCHEMA })
    let finalImpl = impl
    let reworked = false
    if (review && review.verdict === 'rework') {
      log(`Cluster ${c.key}: rework round (${review.findings.length} findings)`)
      const fixed = await agent(reworkPrompt(c, impl, review), { label: `rework:${c.key}`, phase: 'Review', schema: IMPL_SCHEMA })
      if (fixed) {
        finalImpl = fixed
        reworked = true
        review = await agent(reviewPrompt(c, fixed, true), { label: `re-review:${c.key}`, phase: 'Review', schema: REVIEW_SCHEMA })
      }
    }
    return { cluster: c.key, status: review ? review.verdict : 'review-failed', impl: finalImpl, review, reworked }
  }
)

const summary = results.filter(Boolean)
const approved = summary.filter(r => r.status === 'approve')
const problems = summary.filter(r => r.status !== 'approve')
log(`Wave ${WAVE} agents done: ${approved.length}/${CLUSTERS.length} clusters approved${problems.length ? `; needs integrator attention: ${problems.map(p => p.cluster + ':' + p.status).join(', ')}` : ''}`)

return {
  wave: WAVE,
  baseline: { testCount: baseline.testCount, preExisting: baseline.preExisting },
  designBriefs: Object.keys(designBriefs),
  clusters: summary,
  integrationOrder: summary.filter(r => r.impl && r.impl.sha).map(r => ({ cluster: r.cluster, sha: r.impl.sha, branch: r.impl.branch, status: r.status, outcome: r.impl.outcome })),
  gaps: summary.flatMap(r => (r.impl && r.impl.gaps) || []),
}
