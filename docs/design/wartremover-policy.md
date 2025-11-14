# WartRemover Policy

XL uses [WartRemover](https://www.wartremover.org/) to enforce functional programming best practices and prevent common Scala pitfalls at compile time.

## Overview

**Version**: WartRemover 3.4.1
**Integration**: Compiler plugin via Mill build system
**Scope**: All modules (xl-core, xl-ooxml, xl-cats-effect, xl-evaluator, xl-testkit)

## Philosophy Alignment

WartRemover enforces XL's core principles:
- **Purity**: No null, no side effects, no hidden mutations
- **Totality**: No partial functions (.head, .get, etc.)
- **Type Safety**: No unsafe casts or operations
- **Predictability**: No surprising behavior or implicit conversions (except explicit given)

## Wart Tiers

### Tier 1: Core Purity Warts (Errors - Fail Builds)

These warts violate XL's fundamental principles and will **fail compilation**:

| Wart | Description | Why Error? |
|------|-------------|------------|
| `Null` | Prevents use of `null` | Null violates totality; use Option instead |
| `TryPartial` | Prevents `.get` on Try | Partial function; use `.fold` or `.getOrElse` |
| `EitherProjectionPartial` | Prevents `.get` on Either projections | Deprecated API; use `.fold` or pattern matching |
| `IterableOps` | Prevents `.head`/`.tail`/`.last` on collections | Partial functions; use `.headOption`/`.lastOption` |
| `TripleQuestionMark` | Prevents `???` placeholders | Unimplemented code should not compile |
| `ArrayEquals` | Prevents `==` on arrays | Reference equality; use `.sameElements` |
| `JavaConversions` | Prevents automatic Java conversions | Implicit; use explicit conversions |
| `Option2Iterable` | Prevents implicit Option to Iterable | Obscures intent; use `.toSeq` explicitly |

#### Examples

```scala
// ❌ Error: Null
val x: String = null  // Fails compilation

// ✅ Correct: Use Option
val x: Option[String] = None

// ❌ Error: IterableOps
val first = list.head  // Fails compilation

// ✅ Correct: Use headOption
val first = list.headOption.getOrElse(default)

// ❌ Error: TryPartial
val result = Try(operation).get  // Fails compilation

// ✅ Correct: Use fold
val result = Try(operation).fold(handleError, identity)
```

### Tier 2: Code Quality Warts (Warnings - Monitor Only)

These warts highlight code quality issues but **only produce warnings**:

| Wart | Description | Why Warning? |
|------|-------------|-------------|
| `OptionPartial` | Warns on `.get` on Option | Acceptable in tests; prefer `.fold` in production |
| `Var` | Warns on mutable variables | Intentional in performance-critical code (macros) |
| `Return` | Warns on early returns | Intentional in parser optimizations |
| `While` | Warns on while loops | Intentional in macro performance optimizations |
| `AsInstanceOf` | Warns on type casts | Necessary for opaque type internals |
| `IsInstanceOf` | Warns on type checks | Sometimes needed for runtime type inspection |

#### Why Warnings?

- **OptionPartial**: Test code commonly uses `.get` for clearer assertions
- **Var/While/Return**: Performance-critical code (macros, parsers) intentionally uses imperative patterns
- **AsInstanceOf**: Opaque types require internal casts for zero-cost abstractions
- **IsInstanceOf**: Runtime type inspection needed in certain scenarios (e.g., codec dispatch)

## Suppression Guidelines

### When to Suppress

Suppressions are acceptable for:
1. **Macros**: Performance-critical compile-time code
2. **Opaque types**: Internal implementation details
3. **Performance optimizations**: Zero-allocation parsers
4. **Test code**: Clearer assertions with `.get`

### How to Suppress

WartRemover violations can be suppressed using Scala 3's `@annotation.nowarn` or inline comments:

```scala
// Option 1: Inline comment (preferred for single-line suppressions)
parts(0)  // Safe: length == 1 verified above

// Option 2: @SuppressWarnings annotation (for entire methods/classes)
@SuppressWarnings(Array("org.wartremover.warts.Var"))
def imperativeParser(s: String): Result = {
  var index = 0
  // ... performance-critical imperative code
}

// Option 3: @annotation.nowarn (Scala 3 style)
@annotation.nowarn("msg=Var")
def optimizedCode(): Unit = {
  var accumulator = 0
  // ...
}
```

### Suppression Examples

#### Macros (RefLiteral.scala)

```scala
// Macro code uses imperative patterns for compile-time performance
private def parseARef(s: String): ARef =
  var i = 0      // Safe: local mutation in macro
  var col = 0
  while i < n do  // Safe: performance-critical parse loop
    // ...
  parts(0)  // Safe: length == 1 verified above
```

**Why**: Macros run at compile time. Imperative code is faster and doesn't affect runtime purity.

#### Opaque Types (ARef)

```scala
opaque type ARef = Long

private def pack(col: Int, row: Int): ARef =
  val packed = (row.toLong << 32) | (col.toLong & 0xFFFFFFFFL)
  packed.asInstanceOf[ARef]  // Safe: opaque type implementation detail
```

**Why**: Opaque types require internal casts. This is a zero-cost abstraction.

#### Test Code

```scala
test("Parse valid reference") {
  val ref = ARef.parse("A1").toOption.get  // Acceptable in tests
  assertEquals(ref.col, Column.from1(1))
}
```

**Why**: Tests intentionally use `.get` for clearer failure messages. OptionPartial is a warning, not an error.

## Skipped Warts

These warts are **not enabled** because they conflict with XL's design:

| Wart | Why Skipped |
|------|-------------|
| `Throw` | XL uses `IllegalArgumentException` for API validation (fail-fast) |
| `ImplicitConversion` | XL uses `given` conversions extensively for DSL ergonomics |
| `PublicInference` | Scala 3 handles type inference better than Scala 2 |
| `DefaultArguments` | Common in builder APIs and DSL methods |

## Adding New Warts

To enable additional warts:

1. **Evaluate**: Test the wart on the codebase: `./mill __.compile`
2. **Classify**: Decide tier (Error or Warning)
3. **Update build.mill**:
   ```scala
   // For errors:
   "-P:wartremover:traverser:org.wartremover.warts.NewWart"

   // For warnings:
   "-P:wartremover:only-warn-traverser:org.wartremover.warts.NewWart"
   ```
4. **Test**: Verify all tests pass: `./mill __.test`
5. **Document**: Update this file with rationale

## Disabling Warts

If a wart causes excessive false positives:

1. **Document reason** in this file
2. **Remove from build.mill**
3. **Consider alternatives** (e.g., Scalafix rules)

## Integration

### Mill Build System

WartRemover is configured in `build.mill`:

```scala
trait XLModule extends ScalaModule {
  def scalacPluginMvnDeps = Seq(
    mvn"org.wartremover:::wartremover:3.4.1"
  )

  override def scalacOptions = Seq(
    // ... existing options ...

    // Tier 1 warts (errors)
    "-P:wartremover:traverser:org.wartremover.warts.Null",
    // ... 8 more ...

    // Tier 2 warts (warnings)
    "-P:wartremover:only-warn-traverser:org.wartremover.warts.OptionPartial",
    // ... 5 more ...
  )
}
```

### Pre-commit Hooks

WartRemover runs automatically via `.pre-commit-config.yaml`:

```yaml
- id: scala-compile
  name: Compile Scala code (with WartRemover)
  entry: ./mill __.compile
```

Run manually: `pre-commit run scala-compile`

### CI/CD

GitHub Actions runs WartRemover via compilation in `.github/workflows/ci.yml`:

```yaml
- name: Compile with WartRemover
  run: ./mill __.compile
```

Violations will fail CI builds for Tier 1 warts.

## FAQ

### Q: Why isn't OptionPartial an error?

**A**: Test code commonly uses `.get` for clearer assertions. Making it an error would require changes to 20+ test files without improving safety. As a warning, it still provides feedback.

### Q: Can I disable WartRemover for a specific file?

**A**: No. WartRemover runs at the module level. Use `@annotation.nowarn` or comments for specific suppressions.

### Q: What if a wart has a false positive?

**A**: Document the false positive in this file and add a clear inline comment explaining why the code is safe.

### Q: How do I run WartRemover without running all tests?

**A**: `./mill xl-core.compile` (or any module)

## References

- [WartRemover Documentation](https://www.wartremover.org/)
- [Mill Build Tool](https://mill-build.org/)
- [Scala 3 Compiler Options](https://docs.scala-lang.org/scala3/guides/migration/options-lookup.html)
- [XL Core Philosophy](./purity-charter.md)

## Version History

- **2025-11-14**: Initial WartRemover integration (v3.4.1)
  - 8 Tier 1 warts (errors)
  - 6 Tier 2 warts (warnings)
  - Mill 1.0.6 + Scala 3.7.3
  - All 467 tests passing
