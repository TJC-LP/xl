# Linting & Formatting — Scalafmt for Mill

## Setup

This project uses **Scalafmt** for code formatting, integrated into Mill via `ScalafmtModule`.

- **Scalafmt** version: `3.10.1` (configured in `.scalafmt.conf`)
- **Scala dialect**: `scala3` (Scala 3.7.3)
- **Note**: Scalafix setup deferred to avoid plugin complexity; imports organized manually

## Configuration

`.scalafmt.conf` settings:

- Max column: 100 characters
- Indentation: 2 spaces (main, callSite, defnSite)
- Docstrings: Asterisk style with wrapping
- No alignment presets (align.preset = none)
- Chain-breaking on first method dot
- Dangling parentheses enabled
- No import curly brace spaces

## Local Commands

### Formatting

```bash
# Format all modules
./mill __.reformat

# Check formatting (CI mode - fails if not formatted)
./mill __.checkFormat

# Format specific module
./mill xl-core.reformat
```

### Per-Module Commands

```bash
./mill xl-core.reformat       # Format core
./mill xl-ooxml.reformat      # Format OOXML
./mill xl-macros.reformat     # Format macros
```

## CI Integration

GitHub Actions CI (`.github/workflows/ci.yml`) runs:

1. **Format check** — `./mill __.checkFormat`
2. **Compile** — `./mill __.compile`
3. **Test** — `./mill __.test`

Format check runs **before** compilation to fail fast on style violations.

## Pre-commit Hook (Optional)

Save to `.git/hooks/pre-commit` and `chmod +x`:

```bash
#!/usr/bin/env bash
set -euo pipefail

echo "Running Scalafmt..."
./mill __.reformat

echo "Staging changes..."
git add -u
```

## Style Guide Alignment

Formatting follows `docs/plan/25-style-guide.md`:

- Opaque types, enums with `derives CanEqual`
- Extension methods over implicit classes
- Total functions with ADT errors
- Clear macro diagnostics

## Troubleshooting

### Mill cache issues

Clear Mill cache if you see stale formatting:

```bash
./mill clean
rm -rf out/
```

### Format check failures in CI

Run locally before pushing:

```bash
./mill __.checkFormat
```

If it fails, run:

```bash
./mill __.reformat
git add -u
git commit --amend --no-edit
```

## Future: Scalafix Integration

Scalafix (OrganizeImports, etc.) can be added later via:

- `mill-scalafix` plugin
- SemanticDB for unused import removal
- `.scalafix.conf` for rule configuration

For now, we keep it simple with Scalafmt only.

## Useful References

- [Scalafmt docs](https://scalameta.org/scalafmt/)
- [Mill Scalafmt module](https://mill-build.org/mill/javalib/Scala_Module_Config.html#_formatting_with_scalafmt)
- [Scalafmt configuration](https://scalameta.org/scalafmt/docs/configuration.html)
