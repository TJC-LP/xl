# Supported Formula Functions

The `eval` command supports 72 Excel functions.

## Math Functions

| Function | Syntax | Example |
|----------|--------|---------|
| SUM | `=SUM(range)` | `=SUM(A1:A10)` |
| AVERAGE | `=AVERAGE(range)` | `=AVERAGE(B1:B5)` |
| MIN | `=MIN(range)` | `=MIN(C1:C100)` |
| MAX | `=MAX(range)` | `=MAX(D1:D50)` |
| COUNT | `=COUNT(range)` | `=COUNT(A:A)` (counts numeric cells) |
| COUNTA | `=COUNTA(range)` | `=COUNTA(A:A)` (counts non-empty cells) |
| COUNTBLANK | `=COUNTBLANK(range)` | `=COUNTBLANK(A:A)` (counts empty cells) |
| ABS | `=ABS(number)` | `=ABS(-5)` |
| SQRT | `=SQRT(number)` | `=SQRT(16)` → 4 |
| MOD | `=MOD(number, divisor)` | `=MOD(7, 3)` → 1 |
| POWER | `=POWER(number, power)` | `=POWER(2, 10)` → 1024 |
| LOG | `=LOG(number, [base])` | `=LOG(100)` → 2 (base 10) |
| LN | `=LN(number)` | `=LN(2.718)` → 1 |
| EXP | `=EXP(number)` | `=EXP(1)` → 2.718... |
| FLOOR | `=FLOOR(number, significance)` | `=FLOOR(2.5, 1)` → 2 |
| CEILING | `=CEILING(number, significance)` | `=CEILING(2.1, 1)` → 3 |
| TRUNC | `=TRUNC(number, [num_digits])` | `=TRUNC(8.9)` → 8 |
| SIGN | `=SIGN(number)` | `=SIGN(-5)` → -1 |
| INT | `=INT(number)` | `=INT(-8.9)` → -9 |
| PI | `=PI()` | `=PI()` → 3.14159... |
| ROUND | `=ROUND(number, digits)` | `=ROUND(3.14159, 2)` |
| ROUNDUP | `=ROUNDUP(number, digits)` | `=ROUNDUP(3.14159, 2)` |
| ROUNDDOWN | `=ROUNDDOWN(number, digits)` | `=ROUNDDOWN(3.99, 0)` |

## Statistical Functions

| Function | Syntax | Example |
|----------|--------|---------|
| MEDIAN | `=MEDIAN(range)` | `=MEDIAN(A1:A10)` → middle value |
| STDEV | `=STDEV(range)` | `=STDEV(A1:A10)` → sample std dev (n-1) |
| STDEVP | `=STDEVP(range)` | `=STDEVP(A1:A10)` → population std dev (n) |
| VAR | `=VAR(range)` | `=VAR(A1:A10)` → sample variance (n-1) |
| VARP | `=VARP(range)` | `=VARP(A1:A10)` → population variance (n) |

**Note**: STDEV/VAR use Welford's algorithm for numerical stability. Sample variants require at least 2 values; population variants require at least 1.

## Logic Functions

| Function | Syntax | Example |
|----------|--------|---------|
| IF | `=IF(condition, true_val, false_val)` | `=IF(A1>100,"High","Low")` |
| AND | `=AND(cond1, cond2, ...)` | `=AND(A1>0, B1>0)` |
| OR | `=OR(cond1, cond2, ...)` | `=OR(A1="Yes", B1="Yes")` |
| NOT | `=NOT(condition)` | `=NOT(A1=0)` |

## Text Functions

| Function | Syntax | Example |
|----------|--------|---------|
| CONCATENATE | `=CONCATENATE(text1, text2, ...)` | `=CONCATENATE(A1," ",B1)` |
| LEFT | `=LEFT(text, num_chars)` | `=LEFT(A1, 3)` |
| RIGHT | `=RIGHT(text, num_chars)` | `=RIGHT(A1, 4)` |
| LEN | `=LEN(text)` | `=LEN(A1)` |
| UPPER | `=UPPER(text)` | `=UPPER(A1)` |
| LOWER | `=LOWER(text)` | `=LOWER(A1)` |

## Date Functions

| Function | Syntax | Example |
|----------|--------|---------|
| TODAY | `=TODAY()` | Returns current date |
| NOW | `=NOW()` | Returns current datetime |
| DATE | `=DATE(year, month, day)` | `=DATE(2024, 1, 15)` |
| YEAR | `=YEAR(date)` | `=YEAR(A1)` |
| MONTH | `=MONTH(date)` | `=MONTH(A1)` |
| DAY | `=DAY(date)` | `=DAY(A1)` |

## Date Calculation Functions

| Function | Syntax | Example |
|----------|--------|---------|
| EOMONTH | `=EOMONTH(start_date, months)` | `=EOMONTH(A1, 3)` |
| EDATE | `=EDATE(start_date, months)` | `=EDATE(A1, -1)` |
| DATEDIF | `=DATEDIF(start, end, unit)` | `=DATEDIF(A1, B1, "M")` |
| NETWORKDAYS | `=NETWORKDAYS(start, end, [holidays])` | `=NETWORKDAYS(A1, B1)` |
| WORKDAY | `=WORKDAY(start, days, [holidays])` | `=WORKDAY(A1, 10)` |
| YEARFRAC | `=YEARFRAC(start, end, [basis])` | `=YEARFRAC(A1, B1, 1)` |

**DATEDIF units**: `"Y"` (years), `"M"` (months), `"D"` (days), `"MD"` (days ignoring months/years), `"YM"` (months ignoring years), `"YD"` (days ignoring years)

**YEARFRAC basis**: 0=US 30/360, 1=Actual/actual, 2=Actual/360, 3=Actual/365, 4=European 30/360

## Financial Functions

| Function | Syntax | Example |
|----------|--------|---------|
| NPV | `=NPV(rate, value1, ...)` | `=NPV(0.1, A1:A5)` |
| IRR | `=IRR(values, [guess])` | `=IRR(A1:A10)` |
| XNPV | `=XNPV(rate, values, dates)` | `=XNPV(0.1, A1:A5, B1:B5)` |
| XIRR | `=XIRR(values, dates, [guess])` | `=XIRR(A1:A5, B1:B5)` |

## Lookup Functions

| Function | Syntax | Example |
|----------|--------|---------|
| VLOOKUP | `=VLOOKUP(value, range, col, [match])` | `=VLOOKUP(A1, B:D, 2, FALSE)` |
| XLOOKUP | `=XLOOKUP(lookup, lookup_arr, return_arr)` | `=XLOOKUP(A1, B:B, C:C)` |
| INDEX | `=INDEX(array, row, [col])` | `=INDEX(A1:C10, 2, 3)` |
| MATCH | `=MATCH(value, range, [match_type])` | `=MATCH(A1, B:B, 0)` |

## Conditional Functions

| Function | Syntax | Example |
|----------|--------|---------|
| SUMIF | `=SUMIF(range, criteria, [sum_range])` | `=SUMIF(A:A, ">100", B:B)` |
| COUNTIF | `=COUNTIF(range, criteria)` | `=COUNTIF(A:A, "Yes")` |
| AVERAGEIF | `=AVERAGEIF(range, criteria, [avg_range])` | `=AVERAGEIF(A:A, ">100", B:B)` |
| SUMIFS | `=SUMIFS(sum_range, crit_range1, crit1, ...)` | `=SUMIFS(C:C, A:A, "Q1", B:B, ">0")` |
| COUNTIFS | `=COUNTIFS(range1, crit1, range2, crit2, ...)` | `=COUNTIFS(A:A, "Active", B:B, ">100")` |
| AVERAGEIFS | `=AVERAGEIFS(avg_range, crit_range1, crit1, ...)` | `=AVERAGEIFS(C:C, A:A, "Q1", B:B, ">0")` |
| SUMPRODUCT | `=SUMPRODUCT(array1, [array2], ...)` | `=SUMPRODUCT(A1:A10, B1:B10)` |

## Error Handling Functions

| Function | Syntax | Example |
|----------|--------|---------|
| IFERROR | `=IFERROR(value, value_if_error)` | `=IFERROR(A1/B1, 0)` |
| ISERROR | `=ISERROR(value)` | `=ISERROR(A1/0)` |

## Reference Functions

| Function | Syntax | Example |
|----------|--------|---------|
| ROW | `=ROW(ref)` | `=ROW(A5)` → 5 |
| COLUMN | `=COLUMN(ref)` | `=COLUMN(C1)` → 3 |
| ROWS | `=ROWS(range)` | `=ROWS(A1:A10)` → 10 |
| COLUMNS | `=COLUMNS(range)` | `=COLUMNS(A1:D1)` → 4 |
| ADDRESS | `=ADDRESS(row, col, [abs], [a1], [sheet])` | `=ADDRESS(1, 1, 1, TRUE)` → "$A$1" |

**ADDRESS abs_num**: 1=$A$1, 2=A$1, 3=$A1, 4=A1
