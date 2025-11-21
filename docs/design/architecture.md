# Architecture Overview

This document describes the high‑level architecture of XL and how the main modules fit together. It is intended as the single place to understand “how the pieces hang together” before diving into detailed design docs.

## Module Layout

```mermaid
graph TD
  subgraph Core
    Core[xl-core<br/>Domain model]
  end

  subgraph OOXML
    Ooxml[xl-ooxml<br/>OOXML mapping]
  end

  subgraph IO
    Ce[xl-cats-effect<br/>Excel / ExcelIO]
  end

  subgraph Evaluator
    Eval[xl-evaluator<br/>Formula Parser]
  end

  subgraph Test
    Tk[xl-testkit<br/>(law & property tests)]
  end

  Core --> Ooxml
  Core --> Ce
  Core --> Eval
  Core --> Tk
  Ooxml --> Ce
  Ooxml --> Tk
```

- `xl-core`: Pure domain model (`Cell`, `Sheet`, `Workbook`, styles, codecs, optics, macros).
- `xl-ooxml`: Pure OOXML mapping layer (`XlsxReader` / `XlsxWriter`, `OoxmlWorkbook`, `OoxmlWorksheet`, `SharedStrings`, `Styles`).
- `xl-cats-effect`: Effectful interpreters (`Excel[F]` / `ExcelIO`) and true streaming I/O built on Cats Effect, fs2, and fs2-data-xml.
- `xl-evaluator`: Formula parser (`TExpr` GADT, `FormulaParser`, `FormulaPrinter`); evaluator planned (WI-08).
- `xl-testkit`: Reusable generators and law test helpers for the other modules.

## I/O Flow

Two families of APIs share the same domain model:

```mermaid
flowchart LR
  subgraph InMemory["In‑Memory I/O (xl-ooxml)"]
    XR[XlsxReader<br/>ZIP → XML → OOXML → Workbook]
    XW[XlsxWriter<br/>Workbook → OOXML → XML → ZIP]
  end

  subgraph Streaming["Streaming I/O (xl-cats-effect)"]
    SR[StreamingXmlReader<br/>ZIP entry → XML events → RowData]
    SW[StreamingXmlWriter<br/>RowData → XML events → ZIP entry]
  end

  U[User code] --> ExcelAPI[Excel / ExcelIO]
  ExcelAPI -->|read(path)| XR
  ExcelAPI -->|write(wb, path)| XW

  ExcelAPI -->|readStream / readSheetStream| SR
  ExcelAPI -->|writeStreamTrue / writeStreamsSeqTrue| SW
```

- **In‑memory path** (default):
  - `XlsxReader.read` parses a `.xlsx` (ZIP) into in‑memory XML, maps into OOXML model types, then into the pure domain `Workbook`.
  - `XlsxWriter.writeWith` takes a `Workbook`, builds a `StyleIndex` + optional `SharedStrings`, turns domain objects into OOXML, and writes canonical XML parts into a new ZIP.
  - When a `Workbook` was created by `XlsxReader`, a `SourceContext` tracks the original file and a `ModificationTracker`, enabling *surgical modification* (copy unchanged parts verbatim, regenerate only what changed).

- **Streaming path**:
  - `ExcelIO.readStream` / `readSheetStream` open the ZIP and stream a worksheet’s XML through fs2‑data‑xml, yielding a `Stream[F, RowData]` with constant memory use (SST is still materialized once if present).
  - `ExcelIO.writeStreamTrue` / `writeStreamsSeqTrue` write static parts once, then stream worksheet XML events directly to a `ZipOutputStream` from a `Stream[F, RowData]` without ever materializing all rows.

See also:
- `docs/design/io-modes.md` – deeper comparison of in-memory vs streaming modes.
- `docs/reference/performance-guide.md` – guidance on choosing a mode for a given workload.
- `docs/STATUS.md` – current capabilities and performance numbers.

## Formula System Architecture

The formula system (xl-evaluator) provides typed parsing and future evaluation capabilities:

```mermaid
flowchart LR
  subgraph Parse["Formula Parsing (WI-07 ✅)"]
    FS[Formula String<br/>"=SUM(A1:B10)"]
    FP[FormulaParser]
    AST[TExpr AST<br/>FoldRange(...)]
    PR[FormulaPrinter]
  end

  subgraph Transform["AST Operations (Future)"]
    OPT[Optimizations<br/>(constant folding)]
    DEPS[Dependency Graph<br/>(cell references)]
  end

  subgraph Eval["Evaluation (WI-08 Planned)"]
    EV[Evaluator]
    RES[Result Value]
  end

  FS -->|parse| FP
  FP -->|Right(TExpr)| AST
  AST -->|print| PR
  PR -->|String| FS

  AST --> OPT
  AST --> DEPS
  AST --> EV
  EV --> RES
```

### TExpr GADT (Typed Expression Tree)

The core of the formula system is the `TExpr[A]` GADT (Generalized Algebraic Data Type):

```scala
enum TExpr[A] derives CanEqual:
  case Lit[A](value: A)                                          // Literals
  case Ref[A](at: ARef, decode: Cell => Either[CodecError, A])  // Cell references
  case If[A](cond: TExpr[Boolean], ifTrue: TExpr[A], ifFalse: TExpr[A])  // Conditionals

  // Arithmetic (TExpr[BigDecimal])
  case Add(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case Sub(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case Mul(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]
  case Div(x: TExpr[BigDecimal], y: TExpr[BigDecimal]) extends TExpr[BigDecimal]

  // Comparison (TExpr[Boolean])
  case Lt, Lte, Gt, Gte, Eq, Neq  // All extend TExpr[Boolean]

  // Logical (TExpr[Boolean])
  case And(x: TExpr[Boolean], y: TExpr[Boolean]) extends TExpr[Boolean]
  case Or(x: TExpr[Boolean], y: TExpr[Boolean]) extends TExpr[Boolean]
  case Not(x: TExpr[Boolean]) extends TExpr[Boolean]

  // Range aggregation (polymorphic in result type B)
  case FoldRange[A, B](range: CellRange, z: B, step: (B, A) => B, ...) extends TExpr[B]
```

**Type Safety Guarantees**:
- `TExpr[BigDecimal]` — Only numeric operations (Add, Mul, etc.)
- `TExpr[Boolean]` — Only logical operations (And, Or, comparisons)
- `TExpr[String]` — Only text operations (Lit, Concat)
- **Compile-time prevention** of type mixing (cannot Add a Boolean and a BigDecimal)

### FormulaParser (Pure Functional Parser)

Implements recursive descent parser with operator precedence:

**Features**:
- Zero-allocation for common cases (manual char iteration)
- No regex (inline parsing for performance)
- Scientific notation support (1.5E10, 3.14E-7)
- Position-aware error messages (shows line/column)
- Levenshtein distance for function suggestions ("SUMM" → "Did you mean: SUM?")

**Supported Syntax**:
- Literals: 42, 3.14, 1.5E-10, TRUE, "text"
- Cell refs: A1, $A$1, Sheet1!A1
- Ranges: A1:B10
- Operators: +, -, *, /, =, <>, <, <=, >, >=, &
- Functions: SUM, COUNT, AVERAGE, IF, AND, OR, NOT
- Parentheses: for grouping

### Laws Satisfied

1. **Round-trip**: `parse(print(expr)) == Right(expr)` (verified by property tests)
2. **Ring laws**: Add/Mul form commutative semiring over `BigDecimal` nodes
3. **Short-circuit**: And/Or respect left-to-right evaluation semantics
4. **Totality**: All operations return `Either[ParseError, A]` (no exceptions)

### Future: Formula Evaluator (WI-08)

The evaluator will implement: `eval: TExpr[A] => Sheet => Either[EvalError, A]`

**Planned capabilities**:
- Recursive evaluation with cell reference resolution
- Dependency tracking (detect circular references)
- Caching for performance (memoization)
- Short-circuit evaluation for And/Or
- Division by zero handling

See `docs/plan/formula-system.md` for detailed design.

