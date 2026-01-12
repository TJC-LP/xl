# XL Roadmap

> **Track Progress**: [GitHub Issues](https://github.com/TJC-LP/xl/issues)

**Last Updated**: 2025-12-27

---

## TL;DR

**Current Status**: Production-ready with **81 formula functions**, SAX streaming (36% faster than POI), Excel tables, and full OOXML round-trip. 800+ tests passing.

**Current Version**: 0.5.0

---

## Release Roadmap

### v0.5.0 (Current)

Core security hardening complete:

| Feature | Status |
|---------|--------|
| ZIP Bomb Detection | ✅ Done |
| Formula Injection Guards | ✅ Done |
| XXE Prevention | ✅ Done |

### v0.6.0 (Security Polish)

| Feature | Status |
|---------|--------|
| File Size Limits (enforcement) | Planned |
| XLSM Macro Handling | Planned |

### v0.7.0 (Features)

| Feature | Status |
|---------|--------|
| Drawing Layer (Images, Shapes) | Planned |
| Chart Model | Planned |
| Two-phase streaming (SST + styles) | Planned |

### Future

| Feature | Status |
|---------|--------|
| Merged Cells in Streaming Write | Backlog |
| Query API | Backlog |
| Named Ranges | Backlog |
| Pivot Tables | Backlog |

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

## Contributing

1. Check [GitHub Issues](https://github.com/TJC-LP/xl/issues) for open tasks
2. See [CONTRIBUTING.md](../CONTRIBUTING.md) for code guidelines
3. Reference issue number in commits: `fix(ooxml): implement feature (#123)`
