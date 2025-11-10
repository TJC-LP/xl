<chatName="Elegance and Purity Enhancements Plan for XL"/>

High-level goals
- Push effect purity into public APIs (no hidden exceptions; explicit error channels)
- Increase type safety via new opaque types and refined ADTs
- Introduce an elegant, law-governed DSL for building updates without boilerplate
- Provide lightweight optics to make nested updates and queries composable
- Expand compile-time DSLs (macros) for formulas and ergonomics
- Add expressive combinators for filling ranges and building sheets declaratively
- Parameterize OOXML writer behavior (SST policy, formatting) without breaking existing API

Summary of changes by area
1) Purify Excel IO API with explicit error channels alongside existing API (non-breaking)
2) Introduce StyleId as an opaque type replacing raw Int for styles
3) Patch DSL to eliminate Monoid type-ascription friction and add expressive combinators
4) Lightweight optics (Lens/Optional) and focus DSL for Workbook/Sheet/Cell
5) Macros: add fx"" formula literal and small improvements
6) Range and sheet combinators: fillBy/tabulate/putRow/putCol
7) OOXML writer configuration: SstPolicy and WriterConfig
8) Deterministic iteration helpers for users (sorted views)

Detailed implementation plan

1) Effect purity for Excel IO (non-breaking additive API)
- Objective: Expose an API surface that never throws within F, carrying XLError explicitly.
- Files to modify:
  - xl-cats-effect/src/com/tjclp/xl/io/Excel.scala
  - xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala
- Changes:
  - Add a parallel algebra that returns F[XLResult[A]] instead of raising exceptions:
    - New trait ExcelR[F[_]]:
      - def readR(path: Path): F[XLResult[Workbook]]
      - def writeR(wb: Workbook, path: Path): F[XLResult[Unit]]
      - def readStreamR(path: Path): Stream[F, Either[XLError, RowData]]
      - def readSheetStreamR(path: Path, sheetName: String): Stream[F, Either[XLError, RowData]]
      - def readStreamByIndexR(path: Path, sheetIndex: Int): Stream[F, Either[XLError, RowData]]
      - def writeStreamR(path: Path, sheetName: String): Pipe[F, RowData, Either[XLError, Unit]]
      - def writeStreamTrueR(path: Path, sheetName: String, sheetIndex: Int = 1): Pipe[F, RowData, Either[XLError, Unit]]
      - def writeStreamsSeqTrueR(path: Path, sheets: Seq[(String, Stream[F, RowData])]): F[XLResult[Unit]]
  - Keep existing Excel[F] methods to preserve current tests and ergonomics; implement ExcelR[F] in ExcelIO, forwarding to underlying pure readers/writers and mapping exceptions to XLError where needed.
- Logic notes:
  - In ExcelIO.read/write today, errors are mapped to raiseError(new Exception(...)). New R methods return the XLError in Right/Left rather than raising.
  - For stream variants, wrap results as Either[XLError, RowData]; on structural parse failures, emit a single Left followed by stream termination.
- Example signatures:
  - class ExcelIO[F[_]: Async] extends Excel[F] with ExcelR[F]
  - def readR(path: Path): F[XLResult[Workbook]] =
      Sync[F].delay(XlsxReader.read(path))
  - def readStreamR(path: Path): Stream[F, Either[XLError, RowData]] =
      readStream(path).map(Right(_)).handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))
- Side effects/impact:
  - No breaking changes; adds new interface. Callers can migrate to R-methods for explicit errors.
  - Future: deprecate throwing versions after migration period.

2) Replace raw Int style indices with opaque StyleId
- Objective: Prevent accidental mixing of indices and improve type clarity.
- Files to modify:
  - xl-core/src/com/tjclp/xl/style.scala (new opaque type)
  - xl-core/src/com/tjclp/xl/sheet.scala (styleId fields)
  - xl-core/src/com/tjclp/xl/patch.scala (SetStyle case)
  - xl-core/src/com/tjclp/xl/style.scala (StyleRegistry)
  - xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala (parse/write s= attribute)
  - xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala (StyleIndex, OoxmlStyles)
  - All tests referencing styleId Ints
- Data structure/interface changes:
  - New opaque type:
    - opaque type StyleId = Int
    - object StyleId { def apply(i: Int): StyleId = i; extension (s: StyleId) inline def value: Int = s }
  - Update Cell styleId: Option[StyleId]
    - File: xl-core/src/com/tjclp/xl/cell.scala
    - case class Cell(..., styleId: Option[StyleId] = None, ...)
    - def withStyle(styleId: StyleId): Cell
  - Update Patch.SetStyle to take StyleId:
    - case SetStyle(ref: ARef, styleId: StyleId)
  - Update StyleRegistry index mapping:
    - index: Map[String, StyleId]
    - def register(style: CellStyle): (StyleRegistry, StyleId)
    - def get(idx: StyleId): Option[CellStyle]
  - Update StyleIndex and remappings:
    - styleToIndex: Map[String, StyleId]
    - fromWorkbook returns Map[Int, Map[StyleId, StyleId]] or maintain local Ints but convert at boundary; recommended: unify to StyleId across.
  - OOXML adapters:
    - When emitting/reading s= attribute, use styleId.value and StyleId.apply(...)
- Example signatures:
  - extension (cell: Cell) def withStyle(styleId: StyleId): Cell
  - def indexOf(style: CellStyle): StyleId
- Side effects/impact:
  - Internal refactoring across modules, but with opaque types no runtime cost.
  - Test updates: construct StyleId(1) instead of raw 1, or provide given Conversion[Int, StyleId] for ergonomic migration.
- Imports/dependencies:
  - Add import com.tjclp.xl.StyleId where needed.

3) Patch DSL: eliminate Monoid type-ascription friction and add operators
- Objective: Provide first-class DSL ops without requiring cats type ascription on enum cases.
- Files to add:
  - xl-core/src/com/tjclp/xl/dsl.scala
- Files to modify:
  - xl-core/src/com/tjclp/xl/patch.scala (minor: expose Batch.apply helpers)
- New extensions and builders:
  - extension (ref: ARef)
    - def :=(value: CellValue): Patch = Patch.Put(ref, value)
    - def :=(value: String|Int|Long|Double|BigDecimal|Boolean|LocalDateTime)(using conversions.given): Patch
    - def styled(style: CellStyle): Patch = Patch.SetCellStyle(ref, style)
    - def styleId(id: StyleId): Patch = Patch.SetStyle(ref, id)
    - def clearStyle: Patch = Patch.ClearStyle(ref)
  - extension (range: CellRange)
    - def merge: Patch = Patch.Merge(range)
    - def unmerge: Patch = Patch.Unmerge(range)
    - def remove: Patch = Patch.RemoveRange(range)
  - Batch combinators not requiring cats:
    - object PatchBatch { def apply(ps: Patch*): Patch = Patch.Batch(ps.toVector) }
    - extension (p1: Patch) infix def ++(p2: Patch): Patch = Patch.combine(p1, p2)
    - extension (p: Patch) def ++(ps: Seq[Patch]): Patch = Patch.Batch((p +: ps).toVector)
  - Optional sugar:
    - object sheet { def apply(name: String): XLResult[Sheet] = Sheet(name) } for DSL fluency
- Example usage:
  - val updates = (cell"A1" := "Hello") ++ (cell"A1".styled(headerStyle)) ++ range"A1:B1".merge
  - sheet.applyPatch(updates)
- Side effects/impact:
  - No breaking changes; purely additive.
  - Encourages DSL use without Cats syntax imports.

4) Lightweight optics and focus DSL
- Objective: Provide composable, total, law-governed updates without manual Map juggling.
- Files to add:
  - xl-core/src/com/tjclp/xl/optics.scala
- Types and functions:
  - final case class Lens[S, A](get: S => A, set: (A, S) => S)
    - def modify(f: A => A)(s: S): S
  - final case class Optional[S, A](getOption: S => Option[A], set: (A, S) => S)
    - def modify(f: A => A)(s: S): S
  - Predefined optics:
    - object Optics:
      - val sheetCells: Lens[Sheet, Map[ARef, Cell]]
      - def cellAt(ref: ARef): Optional[Sheet, Cell]
      - val cellValue: Lens[Cell, CellValue]
      - val cellStyleId: Lens[Cell, Option[StyleId]]
  - Focus DSL:
    - extension (s: Sheet)
      - def focus(ref: ARef): Optional[Sheet, Cell] = Optics.cellAt(ref)
      - def modifyCell(ref: ARef)(f: Cell => Cell): Sheet = focus(ref).modify(f)(s)
      - def modifyValue(ref: ARef)(f: CellValue => CellValue): Sheet =
          focus(ref).modify(c => Optics.cellValue.modify(f)(c))(s)
- Example signatures:
  - def cellAt(ref: ARef): Optional[Sheet, Cell] =
      Optional(
        getOption = s => s.cells.get(ref).orElse(Some(Cell.empty(ref))),
        set = (c, s) => s.put(c)
      )
- Side effects/impact:
  - No external dependency (no Monocle). Zero allocation in hot paths if used carefully.
  - Enables very concise transformations and composition.

5) Macros: formula literal fx"..."
- Objective: Provide compile-time validated formula literals with zero-cost emission into CellValue.Formula.
- File to modify:
  - xl-macros/src/com/tjclp/xl/macros.scala
- Add to macros object:
  - extension (inline sc: StringContext)
    - transparent inline def fx(): com.tjclp.xl.CellValue =
        ${ fxImpl('sc) }
- Implementation logic:
  - Parse literal content, perform minimal lexical validation (characters, balanced parentheses, no interpolation).
  - Emit '{ com.tjclp.xl.CellValue.Formula(${Expr(formulaString)}) }
- Example signature:
  - private def fxImpl(sc: Expr[StringContext])(using Quotes): Expr[CellValue]
- Side effects:
  - Aligns with README snippets using fx"...".
  - Add helpful compile error messages via report.errorAndAbort on invalid tokens.

6) Range and sheet combinators
- Objective: Ergonomic, deterministic and total combinators for building sheets.
- File to modify:
  - xl-core/src/com/tjclp/xl/sheet.scala (extensions)
- Add new extension methods:
  - extension (sheet: Sheet)
    - def fillBy(range: CellRange)(f: (Column, Row) => CellValue): Sheet
      - Builds cells row-major, reusing putAll; deterministic order.
    - def tabulate(range: CellRange)(f: (Int, Int) => CellValue): Sheet
      - Same as fillBy, but passes 0-based indices.
    - def putRow(row: Row, startCol: Column, values: Iterable[CellValue]): Sheet
      - Put contiguous cells in a row from startCol
    - def putCol(col: Column, startRow: Row, values: Iterable[CellValue]): Sheet
  - Example signatures:
    - def fillBy(range: CellRange)(f: (Column, Row) => CellValue): Sheet =
        val newCells = range.cells.map { ref => Cell(ref, f(ref.col, ref.row)) }.toVector
        sheet.putAll(newCells)
- Side effects:
  - Purely additive, captures a frequent pattern elegantly.

7) OOXML writer configuration: SstPolicy and WriterConfig
- Objective: Allow users to select SST strategy and pretty printing without affecting default behavior.
- File to modify:
  - xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala
  - xl-ooxml/src/com/tjclp/xl/ooxml/SharedStrings.scala (expose fromWorkbook already present)
- New types:
  - enum SstPolicy derives CanEqual:
    - case Auto
    - case Always
    - case Never
  - case class WriterConfig(
      sstPolicy: SstPolicy = SstPolicy.Auto,
      prettyPrint: Boolean = true
    )
- New methods (additive):
  - def writeWith(workbook: Workbook, outputPath: Path, config: WriterConfig = WriterConfig()): XLResult[Unit]
  - wire sst selection:
    - val sst = config.sstPolicy match
        case Always => Some(SharedStrings.fromWorkbook(workbook))
        case Never  => None
        case Auto   => if SharedStrings.shouldUseSST(workbook) then Some(...) else None
  - pretty printing flag:
    - replace XmlUtil.prettyPrint(xml) with conditional compact rendering when prettyPrint=false
      - add XmlUtil.compact(node: Node): String helper
- Maintain existing def write(...) delegating to writeWith with defaults.
- Side effects:
  - None for existing users; new configuration for power users.

8) Deterministic iteration helpers for users
- Objective: Provide sorted views when users iterate, matching writer canonicalization.
- File to modify:
  - xl-core/src/com/tjclp/xl/sheet.scala
- Add extensions:
  - extension (sheet: Sheet)
    - def rowsSorted: Vector[(Row, Vector[Cell])] = ...
    - def cellsSorted: Vector[Cell] = sheet.cells.values.toVector.sortBy(c => (c.row.index0, c.col.index0))
    - def columnsSorted: Vector[(Column, Vector[Cell])] = ...
- Reasoning:
  - The core maps are immutable and iteration order is unspecified; helpers ensure deterministic observation.

Critical architectural decisions
- Error-channel purity: Introduce ExcelR alongside existing Excel to avoid breaking API and tests. This balances purity and ergonomics. Long-term, consider deprecating throwing methods.
- StyleId opaque type: Improves type safety with zero runtime overhead. It spans modules; schedule as a focused refactor with search-and-replace assisted updates.
- No external optics dependency: A minimal optics layer avoids adding heavy libraries and keeps compile-time footprint low.
- DSL additions are purely additive and avoid implicit conflicts.

Exact locations and code sketches

A) ExcelR additive API
- File: xl-cats-effect/src/com/tjclp/xl/io/Excel.scala
  - Add:
    trait ExcelR[F[_]]:
      def readR(path: Path): F[XLResult[Workbook]]
      def writeR(wb: Workbook, path: Path): F[XLResult[Unit]]
      def readStreamR(path: Path): Stream[F, Either[XLError, RowData]]
      def readSheetStreamR(path: Path, sheetName: String): Stream[F, Either[XLError, RowData]]
      def readStreamByIndexR(path: Path, sheetIndex: Int): Stream[F, Either[XLError, RowData]]
      def writeStreamR(path: Path, sheetName: String): Pipe[F, RowData, Either[XLError, Unit]]
      def writeStreamTrueR(path: Path, sheetName: String, sheetIndex: Int = 1): Pipe[F, RowData, Either[XLError, Unit]]
      def writeStreamsSeqTrueR(path: Path, sheets: Seq[(String, Stream[F, RowData])]): F[XLResult[Unit]]
- File: xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala
  - class ExcelIO[F[_]: Async] extends Excel[F] with ExcelR[F]
  - Implement R methods by calling pure XlsxReader/XlsxWriter and mapping exceptions appropriately.
  - Example:
    def readR(path: Path): F[XLResult[Workbook]] =
      Sync[F].delay(XlsxReader.read(path))
    def readStreamR(path: Path): Stream[F, Either[XLError, RowData]] =
      readStream(path).map(Right(_)).handleErrorWith(e => Stream.emit(Left(XLError.IOError(e.getMessage))))

B) StyleId opaque type and propagation
- File: xl-core/src/com/tjclp/xl/style.scala
  - Add:
    opaque type StyleId = Int
    object StyleId:
      def apply(i: Int): StyleId = i
      extension (s: StyleId) inline def value: Int = s
- File: xl-core/src/com/tjclp/xl/cell.scala
  - Change:
    case class Cell(..., styleId: Option[StyleId] = None, ...)
    def withStyle(styleId: StyleId): Cell
- File: xl-core/src/com/tjclp/xl/patch.scala
  - Change enum case:
    case SetStyle(ref: ARef, styleId: StyleId)
- File: xl-core/src/com/tjclp/xl/style.scala (StyleRegistry)
  - Change fields:
    case class StyleRegistry(styles: Vector[CellStyle] = Vector(CellStyle.default),
                             index: Map[String, StyleId] = Map(CellStyle.canonicalKey(CellStyle.default) -> StyleId(0)))
  - Methods:
    def register(style: CellStyle): (StyleRegistry, StyleId) = ...
    def get(idx: StyleId): Option[CellStyle] = styles.lift(idx.value)
    def indexOf(style: CellStyle): Option[StyleId] = index.get(key)
- File: xl-ooxml/src/com/tjclp/xl/ooxml/Worksheet.scala
  - OoxmlCell.styleIndex: Option[Int] stays OOXML-raw; conversion at boundary:
    // when building OoxmlCell from domain:
    val globalStyleIdx: Option[Int] = cell.styleId.map(_.value).orElse(Some(0)).filter(_ != 0)
    // when converting back to domain:
    Cell(ooxmlCell.ref, value, ooxmlCell.styleIndex.map(StyleId.apply))
- File: xl-ooxml/src/com/tjclp/xl/ooxml/Styles.scala
  - StyleIndex.styleToIndex: Map[String, StyleId]
  - indexOf(style: CellStyle): StyleId
  - When writing XML, use styleId.value
- Side note:
  - For a smoother migration, add given Conversion[Int, StyleId] and back in style.scala companion (scoped, not implicit globally) only for ooxml layer if necessary.

C) Patch DSL
- File: xl-core/src/com/tjclp/xl/dsl.scala (new)
  - Provide operators and combinators as outlined.
  - Example:
    package com.tjclp.xl
    object dsl:
      extension (ref: ARef)
        inline def :=(cv: CellValue): Patch = Patch.Put(ref, cv)
        def styled(style: CellStyle): Patch = Patch.SetCellStyle(ref, style)
        def styleId(id: StyleId): Patch = Patch.SetStyle(ref, id)
        def clearStyle: Patch = Patch.ClearStyle(ref)
      extension (range: CellRange)
        def merge: Patch = Patch.Merge(range)
        def unmerge: Patch = Patch.Unmerge(range)
        def remove: Patch = Patch.RemoveRange(range)
      extension (p1: Patch)
        infix def ++(p2: Patch): Patch = Patch.combine(p1, p2)
- No dependency changes; purely additive.

D) Optics and focus DSL
- File: xl-core/src/com/tjclp/xl/optics.scala (new)
  - Define Lens and Optional with modify helpers.
  - Define Optics.{sheetCells, cellAt, cellValue, cellStyleId}
  - Define Sheet-focused extensions:
    extension (s: Sheet)
      def focus(ref: ARef): Optional[Sheet, Cell] = Optics.cellAt(ref)
      def modifyCell(ref: ARef)(f: Cell => Cell): Sheet = focus(ref).modify(f)(s)
      def modifyValue(ref: ARef)(f: CellValue => CellValue): Sheet = ...
- Reasoning: Embrace compositional updates; makes it easy to pipe transformations.

E) Macro: fx"" literal
- File: xl-macros/src/com/tjclp/xl/macros.scala
  - Add in macros extension:
    transparent inline def fx(): com.tjclp.xl.CellValue = ${ fxImpl('sc) }
  - Add implementation:
    private def fxImpl(sc: Expr[StringContext])(using Quotes): Expr[CellValue] =
      import quotes.reflect.report
      val s = literal(sc)
      // Minimal validation (no interpolation, acceptable token set, balanced parentheses if simple)
      '{ com.tjclp.xl.CellValue.Formula(${Expr(s)}) }
- Impact: Aligns with README examples while keeping validation light and total.

F) Range and sheet combinators
- File: xl-core/src/com/tjclp/xl/sheet.scala
  - Add at end extensions section:
    extension (sheet: Sheet)
      def fillBy(range: CellRange)(f: (Column, Row) => CellValue): Sheet = { ... }
      def tabulate(range: CellRange)(f: (Int, Int) => CellValue): Sheet = { ... }
      def putRow(row: Row, startCol: Column, values: Iterable[CellValue]): Sheet = { ... }
      def putCol(col: Column, startRow: Row, values: Iterable[CellValue]): Sheet = { ... }
- Notes:
  - Use range.cells for deterministic row-major iteration.
  - Implement via putAll to avoid multiple map reallocations.

G) XlsxWriter configuration (SST policy, prettyPrint)
- File: xl-ooxml/src/com/tjclp/xl/ooxml/XlsxWriter.scala
  - New types:
    enum SstPolicy derives CanEqual { case Auto, Always, Never }
    case class WriterConfig(sstPolicy: SstPolicy = SstPolicy.Auto, prettyPrint: Boolean = true)
  - Add:
    def writeWith(workbook: Workbook, outputPath: Path, config: WriterConfig = WriterConfig()): XLResult[Unit] = { ... }
  - Modify sst determination:
    val sst = config.sstPolicy match { ... }
  - Modify writePart to switch pretty vs compact:
    private def xmlToBytes(xml: Elem, pretty: Boolean): Array[Byte] =
      val s = if pretty then XmlUtil.prettyPrint(xml) else XmlUtil.compact(xml)
      s.getBytes(StandardCharsets.UTF_8)
    private def writePart(zip, entryName, xml, pretty: Boolean): Unit = { ... }
  - Add compact in XmlUtil:
    def compact(node: Node): String =
      val printer = new PrettyPrinter(0, 0)
      s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${printer.format(node)}"""
  - Keep existing write(...) delegating to writeWith(...).
- Side effects:
  - None to current behavior; default remains Auto + pretty.

H) Deterministic iteration helpers
- File: xl-core/src/com/tjclp/xl/sheet.scala
  - Add:
    extension (sheet: Sheet)
      def cellsSorted: Vector[Cell] =
        sheet.cells.values.toVector.sortBy(c => (c.row.index0, c.col.index0))
      def rowsSorted: Vector[(Row, Vector[Cell])] = { ... }
      def columnsSorted: Vector[(Column, Vector[Cell])] = { ... }
- Reasoning:
  - Users get deterministic views for inspection/transforms consistent with write order.

Potential side effects and mitigation
- StyleId refactor touches many files; keep changes mechanical and add temporary given Conversion[Int, StyleId] and back within ooxml package scope if needed to ease integration.
- Adding ExcelR is additive; consider later deprecations for throwing methods.
- DSL operators must be placed under a dedicated import (e.g., import com.tjclp.xl.dsl.*) to avoid polluting global namespace and prevent operator conflicts.
- Optics add no runtime overhead but keep them simple to avoid ambiguity with future external optics libs.

Critical decisions
- Keep Cats as an optional dependency for Monoid but do not force users to pull syntax to build patches.
- Provide a migration route toward fully explicit error channels in xl-cats-effect without breaking existing code.
- Maintain deterministic output guarantees by ensuring all new combinators and views adhere to sorted iteration.

Dependencies/imports to update
- None new; no external optics added. Update imports where StyleId is used.

Configuration updates
- None required for build. xml compact helper is internal only.

Example usage after changes
- Elegant patch building without type ascription:
  import com.tjclp.xl.dsl.*
  import com.tjclp.xl.macros.{cell, range}
  val p =
    (cell"A1" := "Title") ++
    (cell"A1".styled(CellStyle.default.withFont(Font("Arial", 14.0, bold = true)))) ++
    range"A1:B1".merge

- Focused updates with optics:
  sheet.modifyValue(cell"A1") {
    case CellValue.Text(s) => CellValue.Text(s.toUpperCase)
    case other => other
  }

- Pure effect IO:
  val excel = ExcelIO.instance[cats.effect.IO]
  val resultIO: cats.effect.IO[XLResult[Workbook]] = excel.readR(path)

- Configurable writer:
  XlsxWriter.writeWith(wb, path, WriterConfig(sstPolicy = SstPolicy.Always, prettyPrint = false))

- Range tabulation:
  sheet.tabulate(range"A1:C3") { (c, r) => CellValue.Number(BigDecimal(c + r)) }

This plan is strictly focused on the technical implementation details required to elevate elegance and purity in the codebase, while minimizing breakage and preserving determinism and performance.
