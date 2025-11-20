
# Style Guide

- Prefer **opaque types** for domain quantities.
- Use **enums** for closed sets; `derives CanEqual` everywhere.
- Use **final case class** for all data model types (prevents subclassing, enables JVM optimizations).
- Keep public functions **total**; return ADTs for errors.
- Prefer **extension methods** over implicit classes.
- Macros must emit **clear diagnostics**; avoid surprises in desugaring.
