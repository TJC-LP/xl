
# API Surface — Algebra & Interpreters

## Algebra
```scala
package com.tjclp.xl.api

import com.tjclp.xl.core.*, com.tjclp.xl.core.addr.*
import fs2.Stream

trait Excel[F[_]]:
  def read(path: java.nio.file.Path): F[Workbook]
  def write(wb: Workbook, path: java.nio.file.Path): F[Unit]
  def readStream(path: java.nio.file.Path): Stream[F, WorkbookView]
  def writeStream(path: java.nio.file.Path): fs2.Pipe[F, RowWrite, Unit]

object Excel:
  def apply[F[_]](using F: Excel[F]) = F
```
## Interpreters
- `cats-effect` interpreter wires ZIP FS (read/write), XML streaming, and part assembly.
- All mapping from bytes ↔ ADTs is pure and lives in `xl-ooxml`.

## Typeclasses
- `CellCodec`, `RowCodec`, `NamedRowCodec`; `TableCodec` helper from row codec.
