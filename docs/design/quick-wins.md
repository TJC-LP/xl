# Quick Wins and Low-Effort Improvements

This document captures “small but high‑leverage” improvements that are good starter tasks and keep the codebase and documentation sharp.

The items here are intentionally scoped to be safe and incremental; most should not affect core semantics.

## Documentation Wins

- **Keep top‑level status/docs authoritative**
  - Treat `docs/STATUS.md`, `docs/LIMITATIONS.md`, `docs/design/architecture.md`, and the root `README.md` as the main external story.
  - When behavior changes, update these first; older deep‑dive docs can be left as historical context.

- **Deprecate `cell"..."/putMixed` terminology in docs**
  - Many reference docs still use the old `cell"A1"` / `range"A1:B10"` / `putMixed` names.
  - Bring them in line with the current API:
    - Use `ref"A1"` / `ref"A1:B10"` from `com.tjclp.xl.syntax.*`.
    - Refer to “batch `Sheet.put(ref -> value, ...)`” instead of `putMixed`.

- **Tighten Quick Start and migration guides**
  - `docs/QUICK-START.md` and `docs/reference/migration-from-poi.md` contain pre‑ref API examples and some non‑compiling snippets.
  - Rewrite them around:
    - `import com.tjclp.xl.*` (single import ergonomics).
    - `Sheet("Name")` / `Workbook.empty` + `Workbook.put` / `Workbook.update`.
    - `ref"...", fx"...", money"...", percent"..."` macros.

- **Clarify streaming trade‑offs in one place**
  - Use `docs/design/io-modes.md` + `docs/reference/performance-guide.md` as the canonical description of:
    - In‑memory vs streaming I/O.
    - What is preserved (styles, merges, SST) in each path.
    - When to choose which mode.

## Small Code-Facing Wins (Behavior Already Implemented)

These are **documentation alignment** tasks first; implementation is already in place:

- **Merged cells**
  - In‑memory path (`XlsxReader`/`XlsxWriter` + `OoxmlWorksheet`) already round‑trips `Sheet.mergedRanges` via `<mergeCells>`.
  - Make sure all docs say:
    - “Merged cells are supported in in‑memory I/O.”
    - “Streaming writers do not currently emit merges.”

- **Streaming read behavior**
  - `ExcelIO.readStream` / `readSheetStream` / `readStreamByIndex` now use `fs2.io.readInputStream` and fs2‑data‑xml.
  - Clean up any lingering references to the old “broken streaming read” implementation; treat P6.6 as done and focus on feature coverage trade‑offs instead.

## Small Code-Facing Wins (Potential Future Changes)

These are implementation candidates that are safe and localized, but are *not* implemented yet.

- **`Formatted.putFormatted` should reuse the style‑aware `Sheet.put`**
  - Today `Formatted.putFormatted` writes only the value; `Sheet.put(ref -> formatted)` already creates and registers a `CellStyle` with the right `NumFmt`.
  - Implementation win: delegate to the existing `Sheet.put` path so `putFormatted` preserves number formats consistently.

- **Public stream‑based writer helper**
  - `XlsxWriter` already supports writing to an `OutputStreamTarget`; `writeToBytes` currently uses a temp file.
  - Implementation win: add a `writeToStream` helper and re‑implement `writeToBytes` without touching disk (use `ByteArrayOutputStream` instead).

- **Workbook helpers for “update with Either”**
  - Common pattern: `XlsxReader.read(path)` → `Workbook` → sheet‑level operations that return `XLResult[Sheet]` → `Workbook.put`.
  - Implementation win: add a helper like `Workbook.updateEither(name)(Sheet => XLResult[Sheet])` to reduce boilerplate in callers.

## How to Use This List

- Treat this file as a backlog of “good first issues” and cleanups.
- When you land one of these improvements:
  - Update this document and `docs/STATUS.md` if the item materially changes behavior.
  - Prefer small, isolated PRs that tackle one bullet at a time.

