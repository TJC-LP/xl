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
    Eval[xl-evaluator<br/>(formulas – future)]
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
- `xl-cats-effect`: Effectful interpreters (`Excel[F]` / `ExcelIO`) and true streaming I/O built on Cats Effect, fs2, and fs2‑data‑xml.
- `xl-evaluator`: Reserved for the formula engine (not yet implemented).
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
- `docs/design/io-modes.md` – deeper comparison of in‑memory vs streaming modes.
- `docs/reference/performance-guide.md` – guidance on choosing a mode for a given workload.
- `docs/STATUS.md` – current capabilities and performance numbers.

