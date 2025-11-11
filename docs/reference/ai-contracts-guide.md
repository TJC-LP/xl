# AI Contracts Guide

## Purpose

AI contracts are structured Scaladoc comments that make code behavior explicit and machine-readable. They help:
- **LLMs** reason precisely about function behavior
- **Developers** understand preconditions, postconditions, and error cases at a glance
- **Reviewers** verify correctness against stated guarantees
- **Tools** generate tests and validate implementations

## Format

AI contracts use four standard sections in Scaladoc:

```scala
/** Brief description of what the function does
  *
  * REQUIRES: Preconditions (inputs, state, invariants)
  * ENSURES: Postconditions (guarantees, invariants maintained)
  * DETERMINISTIC: Yes/No (with conditions if applicable)
  * ERROR CASES: Possible errors and return types
  *
  * @param paramName Description
  * @return Description
  */
def functionName(...): ReturnType = ...
```

---

## Section Details

### REQUIRES (Preconditions)

**What**: Conditions that must be true when the function is called

**Format**: Bullet list of requirements
- State assumptions about inputs
- Reference valid ranges, non-null constraints, type requirements
- Mention dependencies on external state (if any)

**Examples**:
```scala
REQUIRES: s is non-empty string containing only A-Z and 0-9
REQUIRES: col is in range 0..16383 (Excel column limit)
REQUIRES: wb contains at least one sheet
REQUIRES: None (total function, accepts all inputs)
```

### ENSURES (Postconditions)

**What**: Guarantees about the return value and side effects

**Format**: Bullet list of guarantees
- What the function returns
- Invariants preserved
- Relationships between inputs and outputs
- Properties of the result

**Examples**:
```scala
ENSURES: Returns Right(ref) if s is valid A1 notation, Left(error) otherwise
ENSURES: Result is normalized (start <= end)
ENSURES: strings contains unique values only (deduplicated)
ENSURES: totalCount >= strings.size (equality when no duplicates)
```

### DETERMINISTIC

**What**: Whether the function produces the same output for the same input

**Format**: Yes/No with optional conditions

**Examples**:
```scala
DETERMINISTIC: Yes (pure function, no randomness or I/O)
DETERMINISTIC: Yes (iteration order is stable via sorted)
DETERMINISTIC: No (depends on current system time)
DETERMINISTIC: Yes if input stream is deterministic
```

### ERROR CASES

**What**: Possible error conditions and how they're represented

**Format**: Bullet list of error scenarios
- Invalid inputs
- Boundary conditions
- Resource failures (if applicable)
- Return type in each case

**Examples**:
```scala
ERROR CASES: None (total function)
ERROR CASES: Returns Left("Invalid column letter") if s contains non-letters
ERROR CASES:
  - Empty input → Left("No column letters in: <input>")
  - Invalid characters → Left("Invalid column letter: <char>")
  - Out of range → Left("Column index out of range: <index>")
```

---

## Complete Examples

### Example 1: Total Function (No Errors)

```scala
/** Convert column index (0-based) to Excel letter (A, B, AA, etc.)
  *
  * REQUIRES: col is in range 0..16383 (Excel's maximum column index)
  * ENSURES:
  *   - Returns uppercase letter sequence (A, B, ..., Z, AA, AB, ..., XFD)
  *   - Result is always non-empty
  *   - Result can be parsed back to col via fromLetter
  * DETERMINISTIC: Yes (pure function)
  * ERROR CASES: None (total function over valid column indices)
  *
  * @param col Column index (0-based)
  * @return Excel column letter (e.g., 0 → "A", 27 → "AB")
  */
def toColumnLetter(col: Int): String = ...
```

### Example 2: Partial Function (With Errors)

```scala
/** Parse A1 notation to cell reference
  *
  * REQUIRES: s is non-empty string
  * ENSURES:
  *   - Returns Right(ARef) if s is valid A1 notation (e.g., "A1", "XFD1048576")
  *   - Returns Left(error) if s is invalid
  *   - Input is normalized to uppercase before parsing (case-insensitive)
  *   - Result.toA1 round-trips to normalized input
  * DETERMINISTIC: Yes (pure parsing with stable normalization)
  * ERROR CASES:
  *   - Empty input → Left("No column letters in: ")
  *   - Invalid column → Left("Invalid column letter: <char>")
  *   - Invalid row → Left("Invalid row number: <num>")
  *   - Out of range → Left("Column/row index out of range")
  *
  * @param s Cell reference in A1 notation (case-insensitive)
  * @return Either error message or ARef
  */
def parse(s: String): Either[String, ARef] = ...
```

### Example 3: Function With State Dependencies

```scala
/** Apply patch to sheet and return updated sheet
  *
  * REQUIRES:
  *   - patch is well-formed (all refs are valid)
  *   - sheet.styleRegistry contains all styleIds referenced in patch
  * ENSURES:
  *   - Returns Right(sheet') with patch applied
  *   - sheet'.cells reflects all Put/Remove operations
  *   - Merged regions updated according to Merge operations
  *   - Original sheet is unchanged (immutable)
  *   - Returns Left(XLError) if patch references invalid styleId
  * DETERMINISTIC: Yes (pure transformation, no side effects)
  * ERROR CASES:
  *   - Invalid styleId → Left(XLError.InvalidStyleId)
  *   - Conflicting merge ranges → Left(XLError.InvalidMerge)
  *
  * @param patch Patch to apply
  * @return Either error or updated sheet
  */
def applyPatch(patch: Patch): Either[XLError, Sheet] = ...
```

### Example 4: Complex Transformation

```scala
/** Create SharedStrings table from workbook text cells
  *
  * REQUIRES: wb contains only valid CellValue.Text cells
  * ENSURES:
  *   - strings contains unique text values (deduplicated)
  *   - indexMap maps each string to its SST index (0-based)
  *   - totalCount = number of CellValue.Text instances in workbook
  *   - totalCount >= strings.size (equality only when no duplicates)
  *   - Iteration order is stable (sheets processed in Vector order)
  * DETERMINISTIC: Yes (iteration order is stable via Vector traversal)
  * ERROR CASES: None (total function, handles empty workbook → empty SST)
  *
  * @param wb Workbook to extract strings from
  * @return SharedStrings with deduplicated strings and total count
  */
def fromWorkbook(wb: Workbook): SharedStrings = ...
```

---

## Guidelines

### DO

✅ **Be specific about preconditions**
```scala
REQUIRES: col is in range 0..16383 (Excel column limit)
```

✅ **State guarantees clearly**
```scala
ENSURES: Result round-trips through parse . print == id
```

✅ **Document all error cases**
```scala
ERROR CASES:
  - Empty input → Left("No column letters")
  - Invalid char → Left("Invalid column letter: X")
```

✅ **Reference specs when relevant**
```scala
ENSURES: Follows OOXML spec (ECMA-376 Part 1, §18.8.21)
```

✅ **Use bullet points for readability**
```scala
ENSURES:
  - Returns None if align == Align.default
  - Returns Some(<alignment/>) otherwise
  - Emits only non-default attributes
```

### DON'T

❌ **Don't be vague**
```scala
REQUIRES: Input must be valid  // What does "valid" mean?
```

❌ **Don't skip error cases**
```scala
ERROR CASES: May fail  // What causes failure? What's returned?
```

❌ **Don't omit determinism info**
```scala
// Missing DETERMINISTIC section entirely
```

❌ **Don't use prose paragraphs**
```scala
ENSURES: This function returns either the parsed value or an error...
// Use bullet points instead!
```

❌ **Don't duplicate standard Scaladoc**
```scala
@param s The string parameter  // Already clear from signature
REQUIRES: s is a string  // Redundant with type system
```

---

## When to Use AI Contracts

### Always Use For:

- **Public API functions** - All exported functions should have contracts
- **Complex logic** - Functions with multiple branches or error cases
- **Spec-defined behavior** - OOXML parsing, Excel compatibility, etc.
- **Pure functions with laws** - Monoid operations, round-trip guarantees
- **Functions with preconditions** - Require valid ranges, non-null, etc.

### Optional For:

- **Simple getters/setters** - If behavior is trivial and obvious
- **Private helpers** - If only called from one place with clear context
- **Obvious type conversions** - `toInt`, `toString`, etc. (unless edge cases exist)

### When in Doubt:

Add the contract! It's better to over-document than under-document for AI readability.

---

## Integration with Testing

AI contracts should align with test coverage:

```scala
/** Parse A1 notation to cell reference
  *
  * REQUIRES: s is non-empty string
  * ENSURES: Returns Right(ARef) if valid, Left(error) if invalid
  * DETERMINISTIC: Yes
  * ERROR CASES:
  *   - Empty input → Left("No column letters")
  *   - Invalid column → Left("Invalid column letter")
  *   - Invalid row → Left("Invalid row number")
  */
def parse(s: String): Either[String, ARef] = ...

// Tests should cover each ERROR CASE:
test("parse rejects empty input") {
  assertEquals(ARef.parse(""), Left("No column letters in: "))
}

test("parse rejects invalid column") {
  assertEquals(ARef.parse("$1"), Left("Invalid column letter: $"))
}

test("parse rejects invalid row") {
  assertEquals(ARef.parse("A0"), Left("Invalid row number: 0"))
}
```

---

## AI Contract Checklist

When adding or reviewing AI contracts, verify:

- [ ] REQUIRES section lists all preconditions
- [ ] ENSURES section covers all return value guarantees
- [ ] DETERMINISTIC clearly states Yes/No (with conditions)
- [ ] ERROR CASES lists all failure modes with return types
- [ ] Contracts are consistent with implementation
- [ ] Tests cover all documented ERROR CASES
- [ ] Bullet points used for readability
- [ ] Spec references included where applicable
- [ ] Round-trip properties documented (if applicable)

---

## References

- **CLAUDE.md** - Project instructions for AI assistant
- **ooxml-quality.md** - Example of comprehensive AI contracts in practice
- **purity-charter.md** - Design principles (totality, determinism, purity)
- **testing-guide.md** - Aligning contracts with test coverage

---

## Example Pull Request Template

When adding AI contracts, include in PR description:

```markdown
## AI Contracts Added

- [x] `Styles.defaultFills` - Spec compliance guarantee
- [x] `SharedStrings.fromWorkbook` - Deduplication and counting logic
- [x] `ARef.parse` - All error cases documented
- [x] `Lens.modify` - Lawful lens behavior guarantee

## Contract Completeness

- [x] All public functions have REQUIRES/ENSURES
- [x] All ERROR CASES have corresponding tests
- [x] DETERMINISTIC stated for all functions
- [x] Spec references added where applicable
```

This ensures reviewers can verify contract completeness and correctness.
