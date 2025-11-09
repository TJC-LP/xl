
# Codecs & Named Tuples — Derivation, Header Binding, Identity Laws

## `CellCodec[A]`
- Instances for: `String`, `BigDecimal`, `Boolean`, `Int`, `Long`, `Double`, `LocalDate`, `LocalDateTime`.
- Writers honor default number formats selected by **match types**.

## `RowCodec[A]` (case class derivation)
- Scala 3 `Mirror.ProductOf[A]` used to derive `read/write` row functions.
- Column mapping default: A,B,C,…; override via annotations later.

## Named tuples (Scala 3.7)
- `(Name: String, Age: Int, Email: Option[String])` with **field‑name selection** at compile time.

### Header binder macro
```
val binder = schema[(Name: String, Age: Int, Email: Option[String])].bindHeaders(sheet, headerRow = 1)
```
- At runtime: validates header cells once; caches indices into an immutable binder.
- At compile time: extracts label tuple; produces precise error if a header is missing or duplicated.

### Laws
- **Identity:** `read(row, hdr, write(row, hdr, nt, s)) == Right(nt)` up to presentation formatting.
- **Stability:** integer column indices remain stable as long as header row is unchanged.
