# Supported Formula Functions

The `eval` command supports 30 Excel functions.

## Math Functions

| Function | Syntax | Example |
|----------|--------|---------|
| SUM | `=SUM(range)` | `=SUM(A1:A10)` |
| AVERAGE | `=AVERAGE(range)` | `=AVERAGE(B1:B5)` |
| MIN | `=MIN(range)` | `=MIN(C1:C100)` |
| MAX | `=MAX(range)` | `=MAX(D1:D50)` |
| COUNT | `=COUNT(range)` | `=COUNT(A:A)` |

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

## Financial Functions

| Function | Syntax | Example |
|----------|--------|---------|
| NPV | `=NPV(rate, value1, ...)` | `=NPV(0.1, A1:A5)` |
| IRR | `=IRR(values, [guess])` | `=IRR(A1:A10)` |

## Lookup Functions

| Function | Syntax | Example |
|----------|--------|---------|
| VLOOKUP | `=VLOOKUP(value, range, col, [match])` | `=VLOOKUP(A1, B:D, 2, FALSE)` |
| XLOOKUP | `=XLOOKUP(lookup, lookup_arr, return_arr)` | `=XLOOKUP(A1, B:B, C:C)` |

## Conditional Functions

| Function | Syntax | Example |
|----------|--------|---------|
| SUMIF | `=SUMIF(range, criteria, [sum_range])` | `=SUMIF(A:A, ">100", B:B)` |
| COUNTIF | `=COUNTIF(range, criteria)` | `=COUNTIF(A:A, "Yes")` |
| SUMIFS | `=SUMIFS(sum_range, crit_range1, crit1, ...)` | `=SUMIFS(C:C, A:A, "Q1", B:B, ">0")` |
| COUNTIFS | `=COUNTIFS(range1, crit1, range2, crit2, ...)` | `=COUNTIFS(A:A, "Active", B:B, ">100")` |
| SUMPRODUCT | `=SUMPRODUCT(array1, [array2], ...)` | `=SUMPRODUCT(A1:A10, B1:B10)` |
