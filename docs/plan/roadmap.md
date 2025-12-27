# XL Roadmap

> **Track Progress**: [Linear Project](https://linear.app/tjc-technologies/project/xl) | [GitHub Issues](https://github.com/TJC-LP/xl/issues)

**Last Updated**: 2025-12-27

---

## TL;DR (For AI Agents)

**Current Status**: Production-ready with **81 formula functions**, SAX streaming (36% faster than POI), Excel tables, and full OOXML round-trip. 800+ tests passing.

**Next Priority**: Security hardening (TJC-480 through TJC-483) before v1.0.0 release.

**Quick Start**: Check Linear for available issues, then use worktree workflow below.

---

## Current Focus

### v1.0.0 Requirements

| Priority | Issue | Title | Status |
|----------|-------|-------|--------|
| Urgent | [TJC-480](https://linear.app/tjc-technologies/issue/TJC-480) | Security: ZIP Bomb Detection | Backlog |
| High | [TJC-481](https://linear.app/tjc-technologies/issue/TJC-481) | Security: Formula Injection Guards | Backlog |
| High | [TJC-482](https://linear.app/tjc-technologies/issue/TJC-482) | Security: File Size Limits | Backlog |
| Medium | [TJC-483](https://linear.app/tjc-technologies/issue/TJC-483) | Security: XLSM Macro Handling | Backlog |

### Backlog (Post-1.0)

| Priority | Issue | Title | Status |
|----------|-------|-------|--------|
| Medium | [TJC-484](https://linear.app/tjc-technologies/issue/TJC-484) | Drawing Layer (Images, Shapes) | Backlog |
| Medium | [TJC-486](https://linear.app/tjc-technologies/issue/TJC-486) | Merged Cells in Streaming Write | Backlog |
| Medium | [TJC-487](https://linear.app/tjc-technologies/issue/TJC-487) | Query API | Backlog |
| Medium | [TJC-488](https://linear.app/tjc-technologies/issue/TJC-488) | Named Ranges | Backlog |
| Low | [TJC-485](https://linear.app/tjc-technologies/issue/TJC-485) | Pivot Tables | Backlog |
| — | [TJC-318](https://linear.app/tjc-technologies/issue/TJC-318) | Chart Model | Backlog |

---

## Worktree Workflow

Use git worktrees for isolated development. See `~/git/CLAUDE.md` for full `gtr` command reference.

### Starting Work

```bash
# 1. Check for existing worktrees
cd ~/git/xl
gtr list

# 2. Create worktree for Linear issue
gtr create TJC-480-zip-bomb-detection

# 3. Work in isolation
cd ~/git/worktrees/xl/TJC-480-zip-bomb-detection
# ... implement ...

# 4. Create PR and clean up
gh pr create
gtr rm TJC-480-zip-bomb-detection
```

### Parallel Work Guidelines

**Safe** (different modules):
- Security (xl-ooxml) + Query API (xl-core)
- Charts (xl-ooxml) + Formula enhancements (xl-evaluator)

**Coordinate** (same module):
- Multiple xl-ooxml changes → merge sequentially

### Module Conflict Matrix

| Module | Risk | Reason |
|--------|------|--------|
| `xl-core/Sheet.scala` | High | Central domain model |
| `xl-ooxml/Worksheet.scala` | Medium | OOXML serialization hub |
| `xl-evaluator/` | Low | Isolated formula module |

---

## Completed Work

All completed phases are documented in git history. Key milestones:

- **P0-P8**: Foundation, OOXML, streaming, codecs, macros
- **WI-07/08/09**: Formula parser, evaluator, 81 functions
- **WI-10**: Excel table support
- **WI-17**: SAX streaming write (36% faster than POI)
- **WI-19**: Row/column property serialization

For historical details: `git log --oneline docs/plan/`

---

## Related Documentation

| Doc | Purpose |
|-----|---------|
| [STATUS.md](../STATUS.md) | Current capabilities |
| [LIMITATIONS.md](../LIMITATIONS.md) | Known limitations |
| [QUICK-START.md](../QUICK-START.md) | Get started in 5 minutes |
| [reference/cli.md](../reference/cli.md) | CLI command reference |
| [reference/performance-guide.md](../reference/performance-guide.md) | Optimization guide |

---

## Adding New Work

1. Create Linear issue with description and acceptance criteria
2. Add `enhancement` or `security` label
3. Link to `xl` project
4. Reference issue ID in commits: `fix(ooxml): implement ZIP bomb detection (TJC-480)`
