---
name: xl-scala-style
description: Scala 3 as written in xl - opaque types, enums, extension methods, XLResult totality, WartRemover tiers and how to react to a wart, import order, the export/opaque/inline compiler landmines that break the published API for external consumers, and what changed in Scala 3.9 LTS. Use when writing or reviewing any .scala source, adding public API, touching exports or the scripting prelude, resolving a WartRemover error, or deciding whether something may be `inline`.
---

# Scala style for xl

Philosophy: purity, totality, determinism, law-governed semantics (`docs/design/purity-charter.md`).
Mechanics: `docs/design/style-guide.md` (formatting, type discipline) and
`docs/design/wartremover-policy.md` (wart tiers, suppression rules). This skill is the working
checklist plus the traps those documents only hint at.

## Non-negotiables

- **Total functions**: fallible public APIs return `XLResult[A] = Either[XLError, A]`. No `null`
  (use `Option`), no thrown exceptions as control flow in `xl-core`, `xl-ooxml`, `xl-evaluator`.
  Effects only behind `F[_]` in `xl-cats-effect`.
- **Opaque types for domain quantities** (`Column`, `Row`, `ARef`, `SheetName`, `Pt`, `Px`,
  `Emu`, `StyleId`). Smart constructors return `Either`; `CellRange` normalizes on construction.
- **Enums for closed sums** with exhaustive `match`, `derives CanEqual` where typed equality
  matters. **`final case class`** for data.
- **Extension methods, not implicit classes.** Same-named extensions on different receivers need
  `@annotation.targetName("distinctJvmName")`.
- **Deterministic output**: canonical ordering for XML and styles; re-serialization is
  byte-identical. Never introduce wall-clock, randomness, or hash-order dependence into output.
- **Import order**: java/javax, scala stdlib, cats/cats-effect/fs2, `com.tjclp.xl.*`, test
  frameworks (test sources only), separated by blank lines.
- **Imports of the public API**: `import com.tjclp.xl.{*, given}` (the `*` alone skips givens).
  Scripts use `import com.tjclp.xl.scripting.{*, given}` and never both in one file.

## WartRemover: reacting to a wart

Tier 1 (compile errors): `Null`, `TryPartial`, `EitherProjectionPartial`, `TripleQuestionMark`,
`ArrayEquals`, `JavaConversions`, `Option2Iterable`. Tier 2 (warnings): `IterableOps`
(`.head`/`.tail`), `OptionPartial` (`.get`), `Var`, `Return`, `While`, `AsInstanceOf`,
`IsInstanceOf`.

1. Fix the code first: `headOption`, pattern match, `fold`, `getOrElse`, recursion or
   `foldLeft` instead of `var`/`while`.
2. Suppress only in macros, zero-allocation parsers, opaque-type internals, or tests, always with a
   comment saying why: `@SuppressWarnings(Array("org.wartremover.warts.Var")) // hot parser loop`.
3. Tier 2 warnings are acceptable in tests; keep them out of new production code.

## Compiler landmines (external consumers break while the repo compiles)

Tests inside `com.tjclp.xl` inherit the package's exports and see real types, so they never catch
these. The gate is the external-consumer probe suite in `xl/test/src/xlprelude/`
(`xlprelude.ScriptingPreludeTest`); extend it whenever you change exports or public signatures.

1. **Opaque-type members must stay non-`inline`.** Companion factories and extension methods that
   touch an opaque representation fail to re-elaborate at call sites outside the package
   ("expression does not take parameters", errors that move between lines across incremental
   builds). The JIT inlines trivial static methods anyway. The authoritative note is in
   `xl-core/src/com/tjclp/xl/addressing/ARef.scala`.
2. **Do not export extension methods on opaque types.** Export forwarders produce path-dependent
   proxy types (`$proxy.ARef`) that never unify at use sites. Companion implicit scope resolves
   them naturally; delete the forwarder instead.
3. **Do not export opaque-type companions as term forwarders.** Factory results lose extension
   lookup (`Column.from0(0).toLetter` fails). Use explicit pairs:
   `type X = pkg.X; val X: pkg.X.type = pkg.X`.
4. **`export X.*` skips givens; `export X.given` carries inherited givens too**, which flattens
   low-priority-trait tricks into ambiguity. Exclude the low-priority given by name on every hop:
   `export X.{default as _, given}`.
5. **No default parameters on extension methods that are merged through wildcard-exported
   objects**: the compiler crashes in Namer on `$default$N` getters. Use explicit `@targetName`
   overloads (see `WorkbookEvaluator.evaluateFormula`).
6. **`export java.time.{LocalDate}` forwards only the type**, not Java statics; scripts must import
   `java.time` directly.
7. **Debugging protocol**: Zinc incremental builds are untrustworthy for inline + opaque
   investigations. Reproduce with `./mill clean` + full compile and isolate with minimal A/B probe
   files (alias-typed vs inferred vs real-typed receivers).

## Everyday gotchas

- **Monoid syntax needs type ascription on enum cases**:
  `(Patch.Put(ref, v): Patch) |+| (Patch.SetStyle(ref, 1): Patch)`. Prefer the DSL's `++`.
- **Macro literals** validate at compile time: `ref"A1"`, `fx"=SUM(A1:B10)"`, `money"$$1,234.56"`
  (`$` anchors need `$$`). Pure `fx""` literals return `CellValue`; runtime-interpolated ones
  return `Either`.
- **Adding a formula function**: add a `FunctionSpec` to the right `FunctionSpecs*` trait in
  `xl-evaluator`; the registry macro collects it. If the build then reports a cyclic reference in
  files you did not touch, `./mill clean xl-evaluator.compile` (see `mill-build`).
- **Product strings** (`WorkbookMetadata.application`, POM description) name "Scala 3", never a
  minor version.

## Scala 3.9 LTS notes (3.8.3 -> 3.9.0, upgraded 2026-09)

- 3.9 is the long-term-support line (maintained at least three years; succeeds 3.3 LTS as the
  library baseline). Artifacts built here are readable only by Scala 3.9.0+ compilers (TASTy is
  forward-compatible only), so scala-cli snippets and docs pin `//> using scala 3.9.0`.
- New warnings to fix rather than ignore (build has `-deprecation -feature`, no `-Werror`):
  implicits defined in an inaccessible companion (3.10 stops finding them), `$` in definition
  names unless backticked, `- 42` / `-42.abs` negative-literal ambiguities, non-variable typed
  patterns in `val` definitions (`x @ (_: T)`), and `SourceFile.jpath` in quotes reflection.
- Semantics: implicit conversion to `AnyRef` is disallowed again; equivalent implicit candidates
  count as divergent; `case x @ (_: A, _: B)` binds the refined tuple type; nested objects
  compile to static inner classes and `Unit`-alias results erase to `void` (binary layout, matters
  only to Java callers and mixed-version classpaths).
- `into` is a soft keyword in type and modifier positions (SIP-71). Zinc now records macro type
  arguments as dependencies. WartRemover 3.6.1 is the plugin build for 3.9.0.

## Pre-PR checklist

1. `./mill mill.scalalib.scalafmt.ScalafmtModule/reformatAll __.sources`
2. `./mill __.compile` (Tier 1 warts) and `./mill __.test`
3. Public surface or export changed? Extend `xl/test/src/xlprelude/` and run
   `scripts/test-examples.sh` + `scripts/verify-skill-snippets.sh --local`.
4. New capability? Exercise it in `examples/scripting_tour.sc` and the xl-scripting skill.
5. `CHANGELOG.md` entry under `## [Unreleased]`.
