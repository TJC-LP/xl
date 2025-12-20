# XL Documentation Index

Welcome to the XL documentation. This index helps you find what you need quickly.

## Quick Links

| I want to... | Go to |
|--------------|-------|
| Get started in 5 minutes | [QUICK-START.md](QUICK-START.md) |
| See code examples | [reference/examples.md](reference/examples.md) |
| Use the CLI tool | [plan/xl-cli.md](plan/xl-cli.md) |
| Contribute code | [CONTRIBUTING.md](CONTRIBUTING.md) |
| Check current status | [STATUS.md](STATUS.md) |

---

## By Audience

### New Users
- **[QUICK-START.md](QUICK-START.md)** - Installation, first spreadsheet, reading files
- **[FAQ-AND-GLOSSARY.md](FAQ-AND-GLOSSARY.md)** - Common questions and terminology

### Library Users
- **[reference/examples.md](reference/examples.md)** - End-to-end code examples
- **[reference/performance-guide.md](reference/performance-guide.md)** - In-memory vs streaming, optimization
- **[reference/migration-from-poi.md](reference/migration-from-poi.md)** - Coming from Apache POI

### CLI Users
- **[plan/xl-cli.md](plan/xl-cli.md)** - Command reference, examples, LLM integration

### Contributors
- **[CONTRIBUTING.md](CONTRIBUTING.md)** - Code quality, pre-commit hooks
- **[reference/testing-guide.md](reference/testing-guide.md)** - MUnit, ScalaCheck, property tests
- **[reference/implementation-scaffolds.md](reference/implementation-scaffolds.md)** - Code templates

---

## By Topic

### Architecture & Design
- **[design/architecture.md](design/architecture.md)** - Module dependencies, I/O flow
- **[design/domain-model.md](design/domain-model.md)** - Core types (Cell, Sheet, Workbook)
- **[design/purity-charter.md](design/purity-charter.md)** - Pure FP principles
- **[design/decisions.md](design/decisions.md)** - Architectural decision records (ADRs)
- **[design/io-modes.md](design/io-modes.md)** - In-memory vs streaming comparison

### Formula System
- **[reference/examples.md](reference/examples.md)** - Formula evaluation examples
- **[plan/conditional-functions.md](plan/conditional-functions.md)** - SUMIF, COUNTIF specs
- **[plan/sumproduct-xlookup.md](plan/sumproduct-xlookup.md)** - Advanced function specs

### Performance
- **[reference/performance-guide.md](reference/performance-guide.md)** - Mode selection, benchmarks
- **[design/io-modes.md](design/io-modes.md)** - Streaming architecture
- **[plan/streaming-improvements.md](plan/streaming-improvements.md)** - Future optimizations

### OOXML Internals
- **[reference/ooxml-research.md](reference/ooxml-research.md)** - Excel file format details

### Limitations & Status
- **[STATUS.md](STATUS.md)** - Current capabilities, test coverage
- **[LIMITATIONS.md](LIMITATIONS.md)** - Known limitations, workarounds

---

## Roadmap & Planning

- **[plan/roadmap.md](plan/roadmap.md)** - Work prioritization, available tasks
- **[plan/strategic-implementation-plan.md](plan/strategic-implementation-plan.md)** - Long-term vision
- **[plan/future-improvements.md](plan/future-improvements.md)** - Enhancement ideas

---

## Releases

- **[RELEASING.md](RELEASING.md)** - Maven Central publishing process
- **[../CHANGELOG.md](../CHANGELOG.md)** - Version history

---

## Archive

Completed feature plans (historical reference):
- [archive/plan/formula-system.md](archive/plan/formula-system.md)
- [archive/plan/sax-streaming-write.md](archive/plan/sax-streaming-write.md)
