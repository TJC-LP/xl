# Error propagation: Excel error VALUES as first-class results

**Status**: shipped in 0.14.0 (GH-344 items 1–6). Builds on GH-337 (elementwise carriage),
GH-339 (branch carriage), GH-338 (logical folds).

Excel distinguishes two failure worlds. An **error value** (`#DIV/0!`, `#N/A`, `#NUM!`, …) is
*data*: it lives in a cell, cascades to dependents, is caught by `IFERROR`/`ISERROR`, and writes
to disk as a `t="e"` cell. A **host failure** (unparseable formula, missing sheet, unknown
function, circular reference) is *infrastructure*: nothing computes, and nothing may pretend it
did. xl models both without giving up totality.

## The channel convention

`EvalError.ErrorValue(err: CellError, context: Option[String])` is an Excel error value **in
flight**. It travels the `Left` channel of `Either[EvalError, A]` — Either's short-circuit *is*
Excel's strict-position absorption — so it stays `IFERROR`/`ISERROR`-catchable like every other
failure, and `FunctionSpec` signatures did not change (the 69 `FunctionSpec[BigDecimal]` specs
are untouched; no value-channel join exists at `call.spec.eval`, which would reintroduce the
deferred-ClassCastException class of GH-193/302/306).

## The two-table law

Two total tables in the `EvalError` companion, **distinct judgments**:

| table | judgment | arms |
|---|---|---|
| `toCellError` | permissive **element demotion** (GH-337): what error ELEMENT does a per-element failure become inside an array? | `ErrorValue(e) → e`; `DivByZero → #DIV/0!`; `RefError → #REF!`; `TypeMismatch/CodecFailed/EvalFailed → #VALUE!`; `CircularRef → fatal` |
| `toErrorValue` | strict **boundary promotion** (GH-344): what error VALUE does a Left become at a CellValue result boundary? | `ErrorValue(e) → e`; `DivByZero → #DIV/0!`; **everything else stays a loud Left** |

`EvalFailed → #VALUE!` in the element table keeps carriage total; the same arm is deliberately
absent from the boundary table so infrastructure failures never launder into `#VALUE!` cells.
`DivByZero` promotes because every raise site is a genuine Excel `#DIV/0!` and scalar `=1/0`
must agree with the array path's element carriage.

## Promotion sites (exhaustive)

1. `SheetEvaluator.evaluateFormulaWith` — funnels every `evaluateFormula` overload,
   `evaluateCell`, `evaluateWithDependencyCheck`, `wb.evaluateFormula`, `recalculate`, and the
   CLI. Headline: `sheet.evaluateFormula("=1/0") == Right(CellValue.Error(Div0))`.
2. `SheetEvaluator.evaluateArrayFormulaImpl` — an error result spills a 1×1 `CellValue.Error`
   at the origin.
3. `Evaluator.evalCrossSheetFormula` — an error-computing precedent delivers `CellValue.Error`
   to its readers (the cell-mediated cascade), memoized per pass like any value.

Plus one policy site: **LET binds error values as VALUES** — a binding whose expression
classifies as an error value binds `CellValue.Error(code)` and evaluation continues
(`=LET(x,1/0,x)` → `#DIV/0!`, `=LET(x,1/0,IFERROR(x,5))` → `5`, `=LET(bad,1/0,42)` → `42`);
host failures keep the loud `LET binding 'x': …` wrap.

## Where ErrorValue is raised

- `ScalarCoercion.coerce` Error arm — every typed argument position, IF/IFS scalar conditions,
  `toIntArg`, LET/`Coerced` positions, and the scalar-entry top-left collapse.
- `ArrayArithmetic.compareCellValues` error arms (left operand first) — scalar `=1<#REF!`.
- The scalar equality fast-path pre-check — `cellValueEquals` itself stays total fold-to-false
  (CriteriaMatcher and the lookups depend on error cells never matching).
- `decodeOrCarried` at the four `Ref`/`SheetRef` decode sites — a decode refusal over a cell
  that carries an error raises the error value (`=A1+1` over `#REF!` is `#REF!`), while decode
  refusals over plain values keep their loud `CodecFailed`.
- `Concat` operand guards and `SWITCH` target/case pre-checks.
- `conditionTruthy`'s error arm — the AND/OR folds short-circuit with the coded error;
  `broadcastNot` mirrors `broadcastIf`'s positional carriage.
- Aggregates: `propagatedElementError` for consumed error elements (expression arrays AND
  raw-range cells after item 6's resolve-then-police), per `Aggregator.propagatesErrors`
  (COUNT skips; COUNTA/COUNTBLANK count by emptiness).
- Op-level `#NUM!`/`#DIV/0!`/`#N/A` classifications at source: pow overflow, SQRT/LOG/LN
  domains, MOD zero divisor, LARGE/SMALL/PERCENTILE/QUARTILE domains, RANK not-found,
  RANDBETWEEN inverted bounds, RATE/IRR/XIRR non-convergence.
- SUMPRODUCT dimension mismatch → `ErrorValue(Value)` (exact-dimension enforcement, Excel).

## Broadcast totality (item 4b)

The shared broadcast machinery is **total in Right**: unequal non-1 axes extend to
`max(l, r)` and positions beyond an operand's extent read as `#N/A` elements
(`={1;2;3}*{1;2}` → `{1; 4; #N/A}`), uniformly across arithmetic, ordered compare, equality,
and `broadcastIf`. SUMPRODUCT deliberately does not ride this law.

## Eager logicals (item 3)

Excel does not short-circuit logical functions: AND/OR evaluate **every** argument
left-to-right; the first failure (error value or refusal) wins; only then does the decisive
fold apply. `=AND(FALSE,1/0)` → `#DIV/0!`.

## TRUE/FALSE text (item 5)

Exactly the text `"TRUE"`/`"FALSE"` — case-insensitive, **no trim** — coerces in condition
positions, through one shared recognition table (`ScalarCoercion.boolTextValue`) used by all
three condition tables (`coerceBool`, `decodeBool`, `conditionTruthy`; the L6 parity law).
Accepted micro-divergence: cell-sourced boolean text coerces too (provenance is invisible at
the decode layer; the direct/bound parity law wins).

## recalculate() contract

Promotion happens inside `evaluateCell`, so error results reach the recalc fold as `Right`s:
they cache as `Formula(expr, Some(Error(code)))`, write as `t="e"` cells (round-trip stable),
and cascade to dependents. `RecalcResult.errors` = could-not-evaluate host failures only;
`isClean` still means `errors.isEmpty` — sharpened to "every formula computed a value
(possibly an error value)"; the new `RecalcResult.excelErrors` lists error-valued cells,
sorted by (sheet, ref). Iterative recalculation converges error-valued members by the existing
exact-equality arm.

## Pins that did NOT move

IFERROR/ISERROR catch both channels; ISERR excludes `#N/A` on both channels; CircularRef stays
structural/fatal in both tables; COUNT-family semantics; `cellValueEquals` totality;
`FunctionSpec`/`EvalContext`/`ArgSpec` signatures; GH-388 NonFatal containment; streaming and
writer determinism. AND/OR **text** refusals and bare-range logical arguments stay loud Lefts
until the named follow-ups land.

## Named follow-ups (deliberately out of scope)

`TypeMismatch → #VALUE!` boundary demotion (`="abc"+1`); `#NAME?` for unknown functions (needs
a structural case); lookup no-match → `ErrorValue(NA)` migration; CriteriaMatcher
error-criteria semantics; AND/OR raw-range text/blank skip parity (GH-338); parser
error-literal tokens (`=#N/A`); `AGGREGATE()`/`NA()`/`ERROR.TYPE`; xl-agent items 7–9 of #344.

## Laws (ErrorValueLawsSpec)

- **L1 boundary**: every error producer yields `Right(CellValue.Error(_))` at
  `evaluateFormula`; parse/missing-workbook/missing-sheet/unknown-function/cycle/LET-name
  canaries stay Left.
- **L2 catchability**: ∀ producers f: `IFERROR(f,42) = 42` ∧ `ISERROR(f) = TRUE`.
- **L3 absorption**: left-operand precedence; first-error-in-argument-order determinism.
- **L4**: `Aggregate(id, r) ≡ Call(spec, r)`, including error policy.
- **L5 broadcast totality**: never Left; beyond-extent positions are exactly the `#N/A` pads.
- **L6**: `decodeBool ≡ coerceBool ≡ conditionTruthy` on text inputs.
