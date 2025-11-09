
# Error Model & Safety — Closed ADTs and Guardrails

```scala
enum XLError derives CanEqual:
  case Io(msg: String)                           // interpreters only
  case Parse(path: String, reason: String)
  case Semantic(reason: String)
  case Validation(reason: String)
  case ZipBomb(reason: String, entry: Option[String])
  case Unsupported(feature: String)
```
- **No throws** in public APIs; convert platform exceptions to `XLError` with cause preserved.
- **Zip safety:** entry count/size/ratio limits; deny traversal (`..` or absolute paths).
- **XLSM** preserved as opaque parts; never executed.
- **Formula injection:** optionally escape untrusted text prefixing with `'` when it starts with `= + - @`.
