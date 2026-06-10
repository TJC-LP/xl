# Code Style Guide

House style for all XL modules. Two tools enforce it mechanically — Scalafmt (formatting) and
WartRemover (purity warts) — and everything else here is convention that reviewers hold the line
on. Philosophy lives in [purity-charter.md](purity-charter.md); this document is about how the
code is written.

All code must pass `./mill __.checkFormat` and compile clean under the Tier 1 warts before merge.

## Formatting (Scalafmt)

Configured in `.scalafmt.conf` at the repo root — **Scalafmt 3.10.1**, `runner.dialect = scala3`,
`maxColumn = 100`, 2-space indentation everywhere (`indent.main/callSite/defnSite = 2`),
`docstrings.style = Asterisk` with wrapping.

```bash
./mill __.reformat       # Format all modules
./mill __.checkFormat    # CI / pre-commit check (fails on drift)
```

Never hand-format around the tool; if a construct formats badly, restructure the code.

## Type Discipline

- **Opaque types for domain quantities**: `Column`, `Row`, `ARef`, `SheetName`, and the style
  units (`Pt`, `Px`, `Emu`, `StyleId`) are opaque wrappers over primitives — zero runtime
  overhead, no accidental `Int`/`Long` mixing.

  ```scala
  opaque type Column = Int   // Cannot pass a raw Int where a Column is expected
  opaque type ARef = Long    // Packed: (row << 32) | col
  ```

- **Enums for closed sums** (e.g. `Patch`, `StylePatch`, `CellValue`, `TExpr`) with exhaustive
  `match` — no wildcard-default escape hatches on domain enums. Add `derives CanEqual` where
  typed equality matters (e.g. `TExpr`, `Anchor`, `RefType`).
- **`final case class` for data**; smart constructors on the companion when invariants exist
  (`SheetName.apply` returns `Either`, `CellRange` normalizes on construction).

## Totality & Error Handling

- All fallible public APIs return `XLResult[A]`:

  ```scala
  type XLResult[A] = Either[XLError, A]   // xl-core/src/com/tjclp/xl/error/XLError.scala
  ```

- No `null` (use `Option`), no partial functions, no thrown exceptions as control flow in the
  pure modules (`xl-core`, `xl-ooxml`, `xl-evaluator`). Effects live only behind `F[_]` in
  `xl-cats-effect`.
- Validation is first-class: return `Either`/`ValidatedNec`, never validate by throwing.

## Extension Methods over Implicit Classes

Use Scala 3 `extension` blocks, not Scala 2 implicit classes. Same-named extension methods on
different receivers need distinct JVM names — disambiguate with `@annotation.targetName`:

```scala
// xl-core/src/com/tjclp/xl/unsafe.scala — `.unsafe` exists for several receiver types,
// so each overload carries its own JVM name:
extension (sheet: Sheet)
  @annotation.targetName("unsafeSheet")
  def unsafe: Sheet = sheet
```

(`xl-core/src/com/tjclp/xl/extensions.scala` uses the same pattern for the `style`/`put`
overload families.)

## Opaque-Type Members Must Stay Non-`inline`

Companion factories and extension methods that touch an opaque type's underlying representation
must **not** be declared `inline`. Inline bodies fail to re-elaborate at call sites outside the
defining package, where the representation is hidden — external consumers of the published
artifacts get compile errors (issue #252). The authoritative warning lives in
`xl-core/src/com/tjclp/xl/addressing/ARef.scala`:

> every member here (and in Column/Row/SheetName/style units) that touches the opaque
> representation must stay NON-inline: inline bodies fail to re-elaborate at call sites outside
> this package […] Do not "optimize" these back to `inline` — the JIT/AOT inlines trivial static
> methods anyway.

This is guarded by external-consumer probes in `xl/test/src/xlprelude/`; a "performance" PR that
re-adds `inline` will break them.

## Import Order

Group imports top-down, separated by blank lines:

1. Java / javax
2. Scala stdlib
3. Cats / Cats Effect / fs2
4. Project (`com.tjclp.xl.*`)
5. Test frameworks (test sources only)

## WartRemover

WartRemover **3.5.6** runs as a compiler plugin on every module (configured in `build.mill`,
trait `XLModuleBase`). Two tiers:

- **Tier 1 (errors, fail the build)**: `Null`, `TryPartial`, `EitherProjectionPartial`,
  `TripleQuestionMark`, `ArrayEquals`, `JavaConversions`, `Option2Iterable`.
- **Tier 2 (warnings, monitored)**: `IterableOps`, `OptionPartial`, `Var`, `Return`, `While`,
  `AsInstanceOf`, `IsInstanceOf` — acceptable in tests, macros, and performance-critical
  parser internals.

Tier rationale, suppression rules, and the process for adding warts are in
[wartremover-policy.md](wartremover-policy.md) — do not duplicate them here.

## Known Gotchas

- **Monoid syntax needs type ascription** on enum cases (the enum case's precise type is not the
  enum type, so `Monoid[Patch]` is not found):

  ```scala
  val p = (Patch.Put(ref, value): Patch) |+| (Patch.SetStyle(ref, 1): Patch)
  ```

  The DSL's `++` on patches avoids this; prefer it in examples and scripts.
- **`{*, given}` imports**: Scala 3's `*` does not pull in given instances — public API examples
  must write `import com.tjclp.xl.{*, given}` (or `com.tjclp.xl.scripting.{*, given}` in
  scripts, never both in one file).
- **`var`/`while`/`return`** are tolerated (Tier 2 warnings) only in macros, zero-allocation
  parsers, and tests — always with a comment saying why.

## Enforcement Pipeline

```bash
./mill __.checkFormat    # Scalafmt drift → CI failure
./mill __.compile        # WartRemover Tier 1 → compile error
./mill __.test           # 3005+ tests
```

Pre-commit hooks (`.pre-commit-config.yaml`) run the same `checkFormat` and `compile` steps;
GitHub Actions runs all three.
