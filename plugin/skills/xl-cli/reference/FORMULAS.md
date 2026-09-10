# Supported Formula Functions

`xl eval "=F(...)"` evaluates any function below (`xl -f book.xlsx -s Sheet eval` against a
workbook, no `-f` for constants); `xl evala` displays or spills an array result. The table is
generated from the binary — `xl functions --json` prints the same rows with their arity, argument
slots and flags — so it never drifts from what the installed `xl` evaluates. The prose notes below
it are hand-written semantics worth knowing.

- `args`: accepted argument count; `n+` = at least n, no upper bound.
- `arguments`: each slot as the parser describes it (`optional …` may be omitted, `…...` repeats).
- `flags`: `date`/`time` — the result is a date or time and infers a number format;
  `dynamic deps` — the cells read are decided at evaluation time (`INDIRECT`, `OFFSET`), so such
  cells are always recalculated; `volatile` — the value can change between two recalculations
  with no input changing (`TODAY`, `NOW`, `RAND`, `RANDBETWEEN`; `xl audit` lists the cells);
  `special form` — parsed by the formula parser itself (`LET`).

<!-- generated:functions:begin -->
| function | args | arguments | flags |
| --- | --- | --- | --- |
| `ABS` | 1 | number | — |
| `ADDRESS` | 2–5 | number, number, optional number, optional boolean, optional text | — |
| `AND` | 1+ | boolean... | — |
| `AVERAGE` | 1+ | number or range... | — |
| `AVERAGEIF` | 2–3 | range, value, optional range | — |
| `AVERAGEIFS` | 3+ | range, range, value... | — |
| `CEILING` | 2 | number, number | — |
| `CELL` | 1–2 | text, optional value | — |
| `CHOOSE` | 2+ | value... | — |
| `COLUMN` | 0–1 | optional value | — |
| `COLUMNS` | 1 | value | — |
| `CONCATENATE` | 1+ | text... | — |
| `COUNT` | 1+ | number or range... | — |
| `COUNTA` | 1+ | number or range... | — |
| `COUNTBLANK` | 1+ | number or range... | — |
| `COUNTIF` | 2 | range, value | — |
| `COUNTIFS` | 2+ | range, value... | — |
| `DATE` | 3 | integer, integer, integer | date |
| `DATEDIF` | 3 | date, date, text | — |
| `DAY` | 1 | date | — |
| `EDATE` | 2 | date, integer | date |
| `EOMONTH` | 2 | date, integer | date |
| `ERROR.TYPE` | 1 | cell | — |
| `EXP` | 1 | number | — |
| `FILTER` | 2–3 | range, array or range, optional value | — |
| `FIND` | 2–3 | text, text, optional integer | — |
| `FLOOR` | 2 | number, number | — |
| `FV` | 3–5 | number, number, number, optional number, optional number | — |
| `HLOOKUP` | 3–4 | cell, range, integer, optional boolean | — |
| `HYPERLINK` | 1–2 | text, optional text | — |
| `IF` | 3 | boolean, value, value | — |
| `IFERROR` | 2 | cell, cell | — |
| `IFNA` | 2 | cell, cell | — |
| `IFS` | 2+ | value... | — |
| `INDEX` | 2–3 | range, number, optional number | — |
| `INDIRECT` | 1–2 | text, optional boolean | dynamic deps |
| `INT` | 1 | number | — |
| `IRR` | 1–2 | range, optional number | — |
| `ISBLANK` | 1 | cell | — |
| `ISERR` | 1 | cell | — |
| `ISERROR` | 1 | cell | — |
| `ISNA` | 1 | cell | — |
| `ISNUMBER` | 1 | cell | — |
| `ISTEXT` | 1 | cell | — |
| `LARGE` | 2 | range, integer | — |
| `LEFT` | 2 | text, integer | — |
| `LEN` | 1 | text | — |
| `LN` | 1 | number | — |
| `LOG` | 1–2 | number, optional number | — |
| `LOWER` | 1 | text | — |
| `MATCH` | 2–3 | value, range, optional number | — |
| `MAX` | 1+ | number or range... | — |
| `MAXIFS` | 3+ | range, range, value... | — |
| `MEDIAN` | 1+ | number or range... | — |
| `MID` | 3 | text, integer, integer | — |
| `MIN` | 1+ | number or range... | — |
| `MINIFS` | 3+ | range, range, value... | — |
| `MOD` | 2 | number, number | — |
| `MONTH` | 1 | date | — |
| `MROUND` | 2 | number, number | — |
| `N` | 1 | cell | — |
| `NA` | 0 | — | — |
| `NETWORKDAYS` | 2–3 | date, date, optional range | — |
| `NOT` | 1 | boolean | — |
| `NOW` | 0 | — | time, volatile |
| `NPER` | 3–5 | number, number, number, optional number, optional number | — |
| `NPV` | 2 | number, range | — |
| `OFFSET` | 3–5 | value, integer, integer, optional integer, optional integer | dynamic deps |
| `OR` | 1+ | boolean... | — |
| `PERCENTILE` | 2 | range, number | — |
| `PI` | 0 | — | — |
| `PMT` | 3–5 | number, number, number, optional number, optional number | — |
| `POWER` | 2 | number, number | — |
| `PV` | 3–5 | number, number, number, optional number, optional number | — |
| `QUARTILE` | 2 | range, integer | — |
| `RAND` | 0 | — | volatile |
| `RANDBETWEEN` | 2 | number, number | volatile |
| `RANK` | 2–3 | number, range, optional integer | — |
| `RATE` | 3–6 | number, number, number, optional number, optional number, optional number | — |
| `RIGHT` | 2 | text, integer | — |
| `ROUND` | 2 | number, number | — |
| `ROUNDDOWN` | 2 | number, number | — |
| `ROUNDUP` | 2 | number, number | — |
| `ROW` | 0–1 | optional value | — |
| `ROWS` | 1 | value | — |
| `RRI` | 3 | number, number, number | — |
| `SEARCH` | 2–3 | text, text, optional integer | — |
| `SEQUENCE` | 1–4 | integer, optional integer, optional number, optional number | — |
| `SIGN` | 1 | number | — |
| `SINGLE` | 1 | array or range | — |
| `SMALL` | 2 | range, integer | — |
| `SORT` | 1–3 | range, optional integer, optional integer | — |
| `SQRT` | 1 | number | — |
| `STDEV` | 1+ | number or range... | — |
| `STDEVP` | 1+ | number or range... | — |
| `SUBSTITUTE` | 3–4 | text, text, text, optional integer | — |
| `SUM` | 1+ | number or range... | — |
| `SUMIF` | 2–3 | range, value, optional range | — |
| `SUMIFS` | 3+ | range, range, value... | — |
| `SUMPRODUCT` | 1+ | array or range... | — |
| `SWITCH` | 3+ | value... | — |
| `TEXT` | 2 | value, text | — |
| `TODAY` | 0 | — | date, volatile |
| `TRANSPOSE` | 1 | range | — |
| `TRIM` | 1 | text | — |
| `TRUNC` | 1–2 | number, optional number | — |
| `UNIQUE` | 1–3 | range, optional boolean, optional boolean | — |
| `UPPER` | 1 | text | — |
| `VALUE` | 1 | text | — |
| `VAR` | 1+ | number or range... | — |
| `VARP` | 1+ | number or range... | — |
| `VLOOKUP` | 3–4 | cell, range, integer, optional boolean | — |
| `WORKDAY` | 2–3 | date, integer, optional range | date |
| `XIRR` | 2–3 | range, range, optional number | — |
| `XLOOKUP` | 3–6 | value, range, range, optional value, optional integer, optional integer | — |
| `XNPV` | 3 | number, range, range | — |
| `YEAR` | 1 | date | — |
| `YEARFRAC` | 2–3 | date, date, optional integer | — |
| `LET` | 3+ | name, value, name, value..., calculation | special form |
<!-- generated:functions:end -->

## Notes on semantics

**Aggregates and criteria.** `SUM`, `COUNT`, `AVERAGE`, `MIN`, `MAX`, `MEDIAN`, `STDEV`, `VAR` and
their variants are variadic: `=SUM(1,2,3)`, `=SUM(A1:A5, B1:B5)`, `=SUM(A1, 5, B1:B3)`. Criteria for
`SUMIF`/`COUNTIF`/`AVERAGEIF`/`SUMIFS`/`COUNTIFS`/`AVERAGEIFS`/`MAXIFS`/`MINIFS` follow Excel:
`">100"`, `"<>0"`, `"Yes"`, and `*`/`?` wildcards (`"Widget*"`). `MAXIFS`/`MINIFS` return `0` when
nothing matches. `SUMPRODUCT` accepts array expressions such as `(A1:A5="Yes")*B1:B5`.

**Statistics.** `STDEV`/`VAR` are the sample forms (n−1); `STDEVP`/`VARP` the population forms
(n). Sample forms need at least two values, population forms at least one; both use Welford's
algorithm for numerical stability. `LARGE`/`SMALL` take a 1-based `k`; `RANK` orders descending
unless `order` is non-zero; `PERCENTILE` takes `p` in `[0, 1]` with inclusive interpolation;
`QUARTILE` takes `quart` 0–4 (0 = min, 2 = median, 4 = max).

**Logic.** `IFS` returns the first TRUE condition's value, else `#N/A`. `SWITCH` takes an optional
trailing default. `CHOOSE` is 1-based; an index out of range is `#VALUE!`. `IFNA` catches only
`#N/A`; `IFERROR` catches every error. `ISERR` is TRUE for every error except `#N/A`; `ISERROR`
includes it. `ISBLANK` is FALSE for a cell holding an empty string. `N` converts to a number
(TRUE → 1, text → 0), `NA()` yields `#N/A`.

**Text.** `TRIM` strips ASCII spaces and collapses internal runs. `MID`/`FIND`/`SEARCH` are 1-indexed;
`FIND` is case-sensitive, `SEARCH` is not and accepts wildcards. `SUBSTITUTE` replaces every
occurrence unless `instance` names one. `VALUE` parses currency, percent and accounting
parentheses (`=VALUE("$1,234.56")` → 1234.56). `TEXT` supports `0`, `#`, `,`, `%`, `$` and date
tokens (`=TEXT(A1, "#,##0.00")`).

**Dates.** `TODAY`/`NOW` are volatile (`xl audit` lists the cells that hold them). `DATEDIF` units:
`"Y"` years, `"M"` months, `"D"` days, `"MD"` days ignoring months and years, `"YM"` months ignoring
years, `"YD"` days ignoring years. `YEARFRAC` basis: 0 = US 30/360, 1 = actual/actual,
2 = actual/360, 3 = actual/365, 4 = European 30/360. `NETWORKDAYS`/`WORKDAY` exclude weekends and
the optional holiday range.

**Financial.** `PMT`, `FV`, `PV`, `RATE`, `NPER` use the standard time-value-of-money identities:
`rate` per period, `nper` periods, `pmt` per period (negative = outflow), `pv` present value
(negative = outflow), `fv` future value (default 0), `type` 0 = end of period (default), 1 =
beginning. `NPV` discounts `value1…` from period 1; `IRR`/`XIRR` accept an optional `guess`;
`XNPV`/`XIRR` pair values with dates.

**Lookup and reference.** `VLOOKUP`/`HLOOKUP` take an optional exact/approximate flag (`FALSE` =
exact); `XLOOKUP` takes lookup, lookup array and return array; `MATCH` takes `match_type` 0 for
exact. `OFFSET` returns a range and composes with aggregates (`=SUM(OFFSET(A1, 1, 0, 5, 1))`);
`INDIRECT` reads the cell its text names. Both are dynamic: the dependency graph cannot see their
targets, so cells holding them are always recalculated. `ADDRESS` `abs_num`: 1 = `$A$1`, 2 = `A$1`,
3 = `$A1`, 4 = `A1`. `ROW()`/`COLUMN()` without an argument refer to the cell being evaluated.
`CELL("filename")` reports the workbook's saved path, `CELL("address", ref)`, `CELL("row", ref)` and
`CELL("col", ref)` the reference parts.

**Arrays (dynamic arrays).** `TRANSPOSE`, `SEQUENCE`, `SORT`, `UNIQUE` and `FILTER` return a grid.
Use `xl evala "=..."` to display it or `--at <ref>` to spill it into the sheet; a 1×1 result
collapses to a scalar. `SEQUENCE(rows, [cols], [start], [step])` defaults to one column starting at
1 with step 1; `SORT(array, [sort_index], [sort_order])` sorts rows by a 1-based column, 1 =
ascending (default), −1 = descending; `UNIQUE` keeps first-seen order; `FILTER(array, include,
[if_empty])` keeps the rows where `include` is truthy and returns `if_empty` (else `#N/A`) when
none match.

**Randomness.** `RAND()` and `RANDBETWEEN(lo, hi)` are volatile; `xl recalc` and `--eval` draw fresh
values on every run.

**LET (special form).** `=LET(name1, value1, [name2, value2, ...], calculation)` binds names for
the calculation. Scope is lexical with `let*` semantics: binding N is visible to bindings N+1… and
to the body, exactly as in Excel. At least one name/value pair and a final calculation are
required; a trailing pair without a body is a parse error. Range-shaped bindings
(`A1:B10`, `Sheet2!A1:A10`) are substituted at their use sites, so `=LET(r, A1:A10, SUM(r))` works.
