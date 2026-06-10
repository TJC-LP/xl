# XL Documentation Index

Welcome to the XL documentation. This index helps you find what you need quickly.

## Quick Links

| I want to... | Go to |
|--------------|-------|
| Get started in 5 minutes | [QUICK-START.md](QUICK-START.md) |
| Write Excel scripts with scala-cli | [reference/scripting.md](reference/scripting.md) |
| See code examples | [reference/examples.md](reference/examples.md) |
| Use the CLI tool | [reference/cli.md](reference/cli.md) |
| Contribute code | [CONTRIBUTING.md](CONTRIBUTING.md) |
| Check current status | [STATUS.md](STATUS.md) |

---

## Issue Tracking

- **[GitHub Issues](https://github.com/TJC-LP/xl/issues)** - Issues and discussions
- **[plan/roadmap.md](plan/roadmap.md)** - Release roadmap

---

## By Audience

### New Users
- **[QUICK-START.md](QUICK-START.md)** - Installation, first spreadsheet, reading files
- **[FAQ-AND-GLOSSARY.md](FAQ-AND-GLOSSARY.md)** - Common questions and terminology

### Library Users
- **[reference/scripting.md](reference/scripting.md)** - Scripting guide: scala-cli + the one-import prelude (`com.tjclp.xl.scripting`)
- **[../examples/scripting_tour.sc](../examples/scripting_tour.sc)** - Canonical runnable tour of the scripting prelude
- **[reference/examples.md](reference/examples.md)** - End-to-end code examples
- **[reference/performance-guide.md](reference/performance-guide.md)** - In-memory vs streaming, optimization
- **[reference/migration-from-poi.md](reference/migration-from-poi.md)** - Coming from Apache POI

### CLI Users
- **[reference/cli.md](reference/cli.md)** - Command reference, examples, LLM integration

### Claude Code Users (Plugin Skills)
- **[../plugin/skills/xl-cli/](../plugin/skills/xl-cli/)** - xl-cli skill: quick reads, single edits, search, visual exports
- **[../plugin/skills/xl-scripting/](../plugin/skills/xl-scripting/)** - xl-scripting skill: type-safe Scala scripts (bulk transforms, pipelines, recalculation)

### Contributors
- **[CONTRIBUTING.md](CONTRIBUTING.md)** - Code quality, pre-commit hooks
- **[design/style-guide.md](design/style-guide.md)** - Code style rules (opaque types, totality, formatting)
- **[reference/testing-guide.md](reference/testing-guide.md)** - MUnit, ScalaCheck, property tests
- **[reference/implementation-scaffolds.md](reference/implementation-scaffolds.md)** - Code templates
- **[reference/ai-contracts-guide.md](reference/ai-contracts-guide.md)** - Scaladoc contract format for parsers/serializers

---

## By Topic

### Architecture & Design
- **[design/architecture.md](design/architecture.md)** - Module dependencies, I/O flow
- **[design/domain-model.md](design/domain-model.md)** - Core types (Cell, Sheet, Workbook)
- **[design/purity-charter.md](design/purity-charter.md)** - Pure FP principles
- **[design/decisions.md](design/decisions.md)** - Architectural decision records (ADRs)
- **[design/io-modes.md](design/io-modes.md)** - In-memory vs streaming comparison
- **[design/query-api.md](design/query-api.md)** - Streaming query API (design-only, not implemented)

### Performance
- **[reference/performance-guide.md](reference/performance-guide.md)** - Mode selection, benchmarks
- **[design/io-modes.md](design/io-modes.md)** - Streaming architecture

### OOXML Internals
- **[reference/ooxml-research.md](reference/ooxml-research.md)** - Excel file format details

### Limitations & Status
- **[STATUS.md](STATUS.md)** - Current capabilities, test coverage
- **[LIMITATIONS.md](LIMITATIONS.md)** - Known limitations, workarounds

---

## Releases

- **[RELEASING.md](RELEASING.md)** - Maven Central publishing process
- **[../CHANGELOG.md](../CHANGELOG.md)** - Version history

---

## Historical / Archive

- **[archive/plan/](archive/plan/)** - Superseded plan documents (v0.10.0 execution & triage); current scheduling lives in [plan/roadmap.md](plan/roadmap.md)
- **[reviews/](reviews/)** - Dated review artifacts ([gpt5-review-template.md](reviews/gpt5-review-template.md), [style-review.md](reviews/style-review.md)), kept for the record
