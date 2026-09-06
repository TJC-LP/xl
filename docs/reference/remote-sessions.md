# Remote Claude Code Sessions

How this repo supports Claude Code when it is not running on a maintainer's laptop: cloud sessions
on [claude.ai/code](https://claude.ai/code) (also `claude --cloud`, the mobile and desktop apps),
scheduled routines, and the `@claude` GitHub Action. All three start from a fresh clone, so only
committed files shape the session.

## What a remote session gets from the repo

| Layer | File(s) | Loaded by |
| --- | --- | --- |
| Instructions | `CLAUDE.md` | every session |
| Always-on rules | `.claude/rules/workflow.md` | every session |
| Skills (on demand) | `.claude/skills/{mill-build,xl-scala-style,xl-testing}/SKILL.md` | when relevant or invoked |
| Commands | `.claude/commands/*.md` (`/release-prep`, `/docs-cleanup-xl`) | on invocation |
| Permissions + hooks | `.claude/settings.json` | every session |
| Toolchain pins | `.mill-version` (Mill 1.1.5), `.mill-jvm-version` (Temurin 25) | the `./mill` launcher |
| Provisioning | `scripts/remote-setup.sh` (SessionStart hook) | cloud sessions, routines, Actions |

Not available remotely: `~/.claude/` (user settings, personal rules, auto-memory), the parent
`~/git/CLAUDE.md`, `.claude/settings.local.json`, the `gtr` worktree helper, GraalVM, LibreOffice
(fixture regeneration). Anything a remote session must know has to live in the files above.

## The sandbox and what the repo does about it

Anthropic's hosted environment is a fresh Ubuntu 24.04 VM on x86_64 with OpenJDK 21, Node, Python,
Docker, `git`, `gh`, and `jq` pre-installed. The repo needs JDK 25 (`javacOptions --release 25`),
Mill 1.1.5, and `scala-cli` for the examples harness. Three things close the gap:

1. **`.mill-jvm-version`** — Mill's launcher reads it and downloads Temurin 25 through Coursier
   before evaluating the build, so `./mill __.compile` works with no setup at all, on any machine.
   (Mill's launcher is fetched from GitHub releases; dependencies come from Maven Central; the JDK
   comes from GitHub's Adoptium releases. All three hosts are on the default *Trusted* network
   allowlist of cloud environments.)
2. **`scripts/remote-setup.sh`** — a SessionStart hook declared in `.claude/settings.json`. It exits
   immediately unless `CLAUDE_CODE_REMOTE=true`. Remotely it fetches the Coursier launcher,
   provisions the same JDK (shared cache with Mill), installs `scala-cli`, appends `JAVA_HOME` and
   `PATH` to `$CLAUDE_ENV_FILE` (sourced before every Bash call), and prints a summary that lands in
   Claude's context. Idempotent; seconds once the environment cache holds the toolchain.

   **TLS-intercepting proxies.** Some sandboxes route HTTPS through a proxy that re-terminates TLS
   and inject its CA into the JVM via `JAVA_TOOL_OPTIONS=-Djavax.net.ssl.trustStore=…`. GraalVM
   native launchers ignore that option: Mill's default `1.1.5-native-*` binary dies with a PKIX
   error while fetching `mill-runner-daemon`, and coursier's `cs` and scala-cli's native image fail
   the same way, so out of the box such a session can neither build nor run the scala-cli gates.
   When the hook sees that truststore (or `XL_REMOTE_JVM_LAUNCHERS=true`) it switches every
   launcher to the JVM, which does honour `JAVA_TOOL_OPTIONS`: it exports
   `MILL_VERSION=<.mill-version>-jvm` (the `./mill` script's opt-in to the JVM launcher; the file
   pin itself is untouched), runs coursier from its JAR, and installs scala-cli as a coursier
   bootstrap launcher. `curl` needs nothing (it reads `SSL_CERT_FILE`). Doing it by hand:
   `export MILL_VERSION=1.1.5-jvm` before `./mill`.
3. **`.claude/settings.json`** — allowlists the build, test, format, gate, and read-only `git`/`gh`
   commands so a session is not blocked on approvals for routine work. Pushing and opening PRs still
   ask.

### The `xl` cloud environment

Cloud environments are configured per organization at claude.ai/admin-settings/cloud-environments.
The TJC organization has an Anthropic-hosted environment named **`xl`** (Trusted network access, no
environment variables) whose setup script is the snippet below; pick it when starting a cloud session
or routine for this repo. The setup script runs once as root before Claude Code launches, must exit
zero within about five minutes, and its result is snapshotted and reused by later sessions, so the
first `./mill` in each session finds the toolchain and dependency cache already on disk.

```bash
#!/bin/bash
# xl: provision the Scala/Mill toolchain once; Anthropic snapshots the result for later sessions.
# Every step is guarded, so this is harmless in a checkout that is not xl.
cd "$(git rev-parse --show-toplevel 2>/dev/null || pwd)"
[ -f scripts/remote-setup.sh ] && CLAUDE_CODE_REMOTE=true bash scripts/remote-setup.sh || true
# Pre-fetch Mill, the compiler bridge, and every dependency (network only, no compile).
[ -f build.mill ] && timeout 180 ./mill --no-daemon __.prepareOffline || true
exit 0
```

Keep this block and the environment's script in sync when either changes. Without the environment
everything still works; the first build of a session just downloads dependencies first, which is why
`CLAUDE.md` tells sessions to give `./mill` a 600000 ms timeout.

## GitHub Actions

`.github/workflows/claude.yml` answers `@claude` mentions on issues, PRs, and reviews. It installs
the same toolchain as `ci.yml` (Temurin 25 + `scala-cli` via `coursier/setup-action`, Coursier and
Mill caches) and grants `contents`, `pull-requests`, and `issues` write permission, so Claude can
compile, run tests, push commits, and open PRs from the workflow. Build, test, and gate commands
are pre-approved through `--allowedTools`; the repo's `.claude/settings.json`, `CLAUDE.md`, rules,
and skills are read from the checkout.

`.github/workflows/claude-code-review.yml` reviews every non-draft PR on open and on each push,
skipping PRs that only touch `docs/**` or `CHANGELOG.md`. It is read-only (`gh` commands only).

Both use the `ANTHROPIC_API_KEY` repository secret.

## Rehearsing the sandbox locally

`scripts/remote-rehearsal.sh` runs the whole story in Docker: Ubuntu 24.04 with only OpenJDK 21,
the committed tree (`git archive HEAD`), then Mill's JDK self-provisioning, the SessionStart hook,
and a module compile plus one suite (`--full` runs everything). Use it after changing
`.mill-jvm-version`, `build.mill` toolchain settings, `scripts/remote-setup.sh`, or the hook wiring.

```bash
scripts/remote-rehearsal.sh                        # native arch, xl-core compile + ColumnSpec
scripts/remote-rehearsal.sh --platform linux/amd64 # emulate the real sandbox architecture
scripts/remote-rehearsal.sh --full                 # __.compile + __.test
```

## Working conventions that differ remotely

- No worktrees: the sandbox clone is already isolated. Work on the current branch; cloud sessions
  push to `claude/*` branches and PRs are opened from the web UI or by asking the session.
- `make install-jar` instead of `make install` (no GraalVM). `xl` then lives in `~/.local/bin`,
  which the hook puts on `PATH`.
- Commits made in cloud sessions carry a `Claude-Session: <url>` trailer automatically.
- Personal memory is absent. Decisions worth keeping across remote sessions belong in `CLAUDE.md`
  or `.claude/rules/`, not in conversation.
