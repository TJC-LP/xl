
# Style Guide

- Prefer **opaque types** for domain quantities.
- Use **enums** for closed sets; `derives CanEqual` everywhere.
- Keep public functions **total**; return ADTs for errors.
- Prefer **extension methods** over implicit classes.
- Macros must emit **clear diagnostics**; avoid surprises in desugaring.
