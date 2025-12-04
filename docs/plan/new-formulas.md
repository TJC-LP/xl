# xl-cli Formula Expansion Priorities

## Currently Supported (30 functions)

**Math:** SUM, AVERAGE, MIN, MAX, COUNT
**Logic:** IF, AND, OR, NOT
**Text:** CONCATENATE, LEFT, RIGHT, LEN, UPPER, LOWER
**Date:** TODAY, NOW, DATE, YEAR, MONTH, DAY
**Financial:** NPV, IRR
**Lookup:** VLOOKUP, XLOOKUP
**Conditional:** SUMIF, COUNTIF, SUMIFS, COUNTIFS, SUMPRODUCT

---

## Priority Tiers

### 🔴 P0: Critical for PE/Financial Models

These are used constantly and block real workflows without them.

| Function | Use Case | Example |
|----------|----------|---------|
| **IFERROR** | Error handling (everywhere) | `=IFERROR(A1/B1, 0)` |
| **ISERROR** | Error checking in IF | `=IF(ISERROR(X), 0, X)` |
| **ROUND** | Presentation / rounding | `=ROUND(A1, 2)` |
| **ROUNDUP** | Conservative rounding | `=ROUNDUP(A1, 0)` |
| **ROUNDDOWN** | Floor rounding | `=ROUNDDOWN(A1, 0)` |
| **ABS** | Absolute value | `=ABS(A1-B1)` |
| **INDEX** | Flexible lookup | `=INDEX(B:B, MATCH(A1, A:A, 0))` |
| **MATCH** | Position lookup | `=MATCH("Revenue", A:A, 0)` |
| **XIRR** | IRR with dates (PE returns) | `=XIRR(cashflows, dates)` |
| **XNPV** | NPV with dates | `=XNPV(rate, cashflows, dates)` |

**Rationale:** Saw `ISERROR` in Syndigo valuation. XIRR/XNPV are essential for PE fund return calculations. INDEX/MATCH is preferred over VLOOKUP in serious models.

---

### 🟡 P1: High Value for Financial Models

| Function | Use Case | Example |
|----------|----------|---------|
| **EOMONTH** | End of month (projections) | `=EOMONTH(A1, 3)` |
| **EDATE** | Add months to date | `=EDATE(A1, 12)` |
| **YEARFRAC** | Fraction of year (day count) | `=YEARFRAC(A1, B1)` |
| **PMT** | Loan payment | `=PMT(rate, nper, pv)` |
| **PV** | Present value | `=PV(rate, nper, pmt)` |
| **FV** | Future value | `=FV(rate, nper, pmt)` |
| **RATE** | Implied rate | `=RATE(nper, pmt, pv)` |
| **NPER** | Number of periods | `=NPER(rate, pmt, pv)` |
| **MEDIAN** | Median (comps analysis) | `=MEDIAN(A1:A10)` |
| **STDEV** | Volatility | `=STDEV(A1:A100)` |
| **COUNTA** | Count non-empty | `=COUNTA(A:A)` |
| **ISBLANK** | Check empty | `=IF(ISBLANK(A1), "N/A", A1)` |
| **ISNUMBER** | Type check | `=ISNUMBER(A1)` |
| **ISTEXT** | Type check | `=ISTEXT(A1)` |

**Rationale:** Date functions critical for projection models. TVM functions for debt modeling. Stats for comps analysis.

---

### 🟢 P2: Nice to Have

| Function | Use Case | Example |
|----------|----------|---------|
| **TRIM** | Clean whitespace | `=TRIM(A1)` |
| **TEXT** | Format as text | `=TEXT(A1, "$#,##0")` |
| **VALUE** | Text to number | `=VALUE(A1)` |
| **MID** | Extract substring | `=MID(A1, 2, 5)` |
| **FIND** | Find position | `=FIND("Q", A1)` |
| **SUBSTITUTE** | Replace text | `=SUBSTITUTE(A1, "old", "new")` |
| **MOD** | Modulo | `=MOD(A1, 4)` |
| **INT** | Integer part | `=INT(A1)` |
| **CEILING** | Round up to multiple | `=CEILING(A1, 0.05)` |
| **FLOOR** | Round down to multiple | `=FLOOR(A1, 1000)` |
| **POWER** | Exponent | `=POWER(1.1, 5)` |
| **SQRT** | Square root | `=SQRT(A1)` |
| **LN** | Natural log | `=LN(A1)` |
| **EXP** | e^x | `=EXP(A1)` |
| **LOG** | Logarithm | `=LOG(A1, 10)` |
| **ROW** | Current row | `=ROW()` |
| **COLUMN** | Current column | `=COLUMN()` |
| **ROWS** | Count rows | `=ROWS(A1:A10)` |
| **COLUMNS** | Count columns | `=COLUMNS(A1:E1)` |
| **OFFSET** | Dynamic range | `=OFFSET(A1, 1, 0)` |
| **INDIRECT** | Dynamic reference | `=INDIRECT("A"&B1)` |
| **CHOOSE** | Select from list | `=CHOOSE(A1, "Q1", "Q2", "Q3", "Q4")` |

---

### 🔵 P3: Modern Excel (Lower Priority)

| Function | Use Case | Example |
|----------|----------|---------|
| **IFS** | Multiple conditions | `=IFS(A1>90,"A", A1>80,"B")` |
| **SWITCH** | Case statement | `=SWITCH(A1, 1,"Jan", 2,"Feb")` |
| **MAXIFS** | Conditional max | `=MAXIFS(B:B, A:A, "2024")` |
| **MINIFS** | Conditional min | `=MINIFS(B:B, A:A, ">0")` |
| **FILTER** | Dynamic filter | `=FILTER(A:B, B:B>100)` |
| **SORT** | Dynamic sort | `=SORT(A1:B10, 2, -1)` |
| **UNIQUE** | Deduplicate | `=UNIQUE(A:A)` |
| **SEQUENCE** | Generate sequence | `=SEQUENCE(10, 1, 2020)` |
| **LET** | Named calculations | `=LET(x, A1*2, x+1)` |
| **LAMBDA** | Custom functions | Advanced |

---

## Implementation Recommendation

### Phase 1: P0 (Critical)
```
IFERROR, ISERROR, ROUND, ROUNDUP, ROUNDDOWN, ABS, INDEX, MATCH, XIRR, XNPV
```
**+10 functions → 40 total**

This unlocks:
- Error-safe formulas (IFERROR wraps everything in real models)
- Proper rounding for display
- Flexible lookups (INDEX/MATCH > VLOOKUP)
- PE fund return calculations (XIRR/XNPV are table stakes)

### Phase 2: P1 (High Value)
```
EOMONTH, EDATE, YEARFRAC, PMT, PV, FV, RATE, NPER, MEDIAN, STDEV, COUNTA, ISBLANK, ISNUMBER, ISTEXT
```
**+14 functions → 54 total**

### Phase 3: P2 + P3 (Complete)
Remaining functions as needed.

---

## What NOT to Prioritize

These are rarely used in financial models:

- **DGET, DSUM, etc.** — Database functions (use pivot tables instead)
- **FREQUENCY** — Histogram (rare in models)
- **RANK** — Ranking (usually done in pandas)
- **HYPERLINK** — Links (not calc-related)
- **GETPIVOTDATA** — Pivot extraction (complex)
- **Cube functions** — OLAP (enterprise only)
- **INFO, CELL** — Metadata (edge cases)

---

## Summary

| Phase | Functions | Total | Coverage |
|-------|-----------|-------|----------|
| Current | 30 | 30 | Basic models |
| +P0 | +10 | 40 | **Most PE valuations** |
| +P1 | +14 | 54 | Full financial modeling |
| +P2/P3 | +30 | 84 | Excel power users |

**Recommendation:** Ship P0 first. That gets you from "useful for exploration" to "can eval most real formulas."

