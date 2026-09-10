<!-- GENERATED from `xl functions --json` by DocsGenSpec; do not edit by hand.
     A diff here is a contract change. Regenerate: XL_UPDATE_DOCS=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.DocsGenSpec -->

# Formula functions

118 functions in the evaluator's registry, plus `LET` — a parser-level special form
(lexical bindings) that the registry does not hold. `args` is the accepted argument count
(`n+` = at least n, no upper bound); `arguments` names each slot as the parser describes it
(`optional …` may be omitted, `…...` repeats). `flags`: `date`/`time` — the result is a date
or time (drives number-format inference); `dynamic deps` — the cells read are decided at
evaluation time (INDIRECT/OFFSET), so such cells are always recalculated; `volatile` — the
value can change between two recalculations with no input changing (TODAY/NOW/RAND/
RANDBETWEEN), which `xl audit` reports. Run `xl functions --json` for this table as JSON,
`xl eval "=F(...)"` to try one.

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
