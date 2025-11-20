# Contributing Guide

## Code Quality

- Use Scala 3.7; run `./mill __.reformat` before committing.
- All code must pass `./mill __.compile` (includes WartRemover checks).
- Run `./mill __.test` to verify all tests pass.
- Pre-commit hooks will automatically check formatting and compilation.

## Pre-commit Hooks

XL uses pre-commit hooks to automatically verify code quality before commits.

### Setup (One-Time)

1. Install pre-commit framework:
```bash
pip install pre-commit
# or: brew install pre-commit
```

2. Install the git hook scripts:
```bash
cd /path/to/xl
pre-commit install
```

3. (Optional) Run against all files to verify setup:
```bash
pre-commit run --all-files
```

### What Gets Checked

The hooks automatically run on `git commit` and verify:

- **Scalafmt formatting** (`./mill __.checkFormat`)
- **Compilation** (`./mill __.compile` - includes WartRemover)
- **Trailing whitespace** removal
- **End-of-file** newlines
- **YAML validity** (.yml/.yaml files)
- **Merge conflicts** detection
- **Large files** prevention (>1MB)

If any check fails, the commit is aborted with a clear error message.

### Manual Execution

Run hooks manually without committing:
```bash
# Run all hooks
pre-commit run --all-files

# Run specific hook
pre-commit run scala-compile --all-files
pre-commit run scalafmt-check --all-files
```

### Troubleshooting

**"Command not found: pre-commit"**
- Install via `pip install pre-commit` or `brew install pre-commit`

**"Hook failed but I want to commit anyway"**
- Not recommended, but use `git commit --no-verify` to skip hooks
- Fix the issues first - hooks prevent broken code from entering the repository

**"Compilation is slow in pre-commit"**
- Hooks use Mill's incremental compilation (fast for small changes)
- Initial run may be slower (~10s), subsequent runs are <2s

## WartRemover

XL uses WartRemover to enforce purity and totality. Your code must pass all Tier 1 warts (errors):

- No `null`, `.head`, `.tail`, `.get` on Try/Either
- See `docs/design/wartremover-policy.md` for complete policy

Tier 2 warts (warnings) are monitored but don't fail builds:
- `.get` on Option is acceptable in test code
- `var`/`while`/`return` are acceptable in performance-critical macros

## Testing

- Add property tests for new laws; expand golden corpus.
- Test coverage goal: 90%+ with property-based tests for all algebras.
- Use MUnit + ScalaCheck for property tests.

## Code Style

XL follows strict coding conventions to maintain purity, totality, and law-governed semantics:

- Prefer **opaque types** for domain quantities (Column, Row, ARef, etc.)
- Use **enums** for closed sets; `derives CanEqual` everywhere
- Use **final case class** for all data model types (prevents subclassing, enables JVM optimizations)
- Keep public functions **total**; return ADTs (Either/Option) for errors
- Prefer **extension methods** over implicit classes
- Macros must emit **clear diagnostics**; avoid surprises in desugaring
- Keep printers deterministic; update canonicalization when structures change
- Document every new public type in `/docs` with examples

For detailed style rationale and examples, see `docs/design/purity-charter.md` and `CLAUDE.md`.
