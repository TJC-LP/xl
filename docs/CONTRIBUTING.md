# Contributing Guide

## Code Quality

- Use Scala 3.7; run `./mill __.reformat` before committing.
- All code must pass `./mill __.compile` (includes WartRemover checks).
- Run `./mill __.test` to verify all tests pass.
- Pre-commit hooks will automatically check formatting and compilation.

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

- Keep printers deterministic; update canonicalization when structures change.
- Document every new public type in `/docs` with examples.
- Prefer **opaque types** for domain quantities.
- Use **enums** for closed sets; `derives CanEqual` everywhere.
- Keep public functions **total**; return Either/Option for errors.
