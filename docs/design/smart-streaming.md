# Smart Streaming Architecture

## Vision

Make the XL CLI handle files of any size with optimal performance by automatically choosing between streaming and in-memory modes based on operation requirements and file characteristics.

**Goal**: Every CLI operation should be as fast as possible, with O(1) memory where feasible.

---

## Shipped (as of 0.11.0)

### Manual Mode Selection
- `--stream` flag opts into O(1) streaming for compatible commands
- `--max-size` flag overrides security limits for large in-memory loads
- Default: in-memory load (safe for files <100MB uncompressed)

### Streaming Support Matrix

| Command | Streaming | In-Memory | Notes |
|---------|-----------|-----------|-------|
| `search` | ✅ O(1) | ✅ | Streaming stops early on limit |
| `stats` | ✅ O(1) | ✅ | Streaming aggregation |
| `bounds` | ✅ O(1) | ✅ | `<dimension>` fast path; `--scan` forces full scan |
| `view` (md/csv/json) | ✅ O(1) | ✅ | No styles needed; `--eval` not supported when streaming |
| `view` (html/svg/pdf) | ❌ | ✅ | Requires style registry |
| `cell` | ✅ O(1) | ✅ | SAX single-cell reader (`SaxSingleCellReader`); dependents (reverse deps) need in-memory |
| `eval` | ❌ | ✅ | Needs formula analysis |
| `sheets` | ❌ | ✅ | Needs workbook metadata |
| `names` | ❌ | ✅ | Needs workbook metadata |
| `put` / `putf` / `style` / `batch` | ✅ O(1) | ✅ | SAX→StAX transform (`ZipTransformer`); copies untouched ZIP entries verbatim |
| Other writes | ❌¹ | ✅ | Read-modify-write cycle |

¹ For all other modifying commands, `--stream` still helps: the full workbook is loaded in
memory, but the write goes through the lower-allocation SAX/StAX backend
(`ExcelIO.writeWorkbookStream`).

---

## Architecture Phases: Status

Each phase below was designed when streaming was read-only (v0.8.0 era). Status as of 0.11.0:

| Phase | Status |
|-------|--------|
| 1. Lazy metadata loading | ⬜ Future — `sheets`/`names` still require full load; no `LazyWorkbook` |
| 2. Style-aware streaming | ⚠️ Partial — `RowData`/`StyledRowData` carry style indices; styled HTML/SVG streaming not shipped |
| 3a. Append-only writes | ⚠️ Partial — `import --stream` streams CSV→new sheet (O(1)); no append-to-existing-sheet API |
| 3b. Range-based modifications | ✅ Shipped in 0.8.0 (#183) — `StreamingTransform` + `ZipTransformer`; CLI `--stream` `put`/`putf`/`style`/`batch` |
| 4. Intelligent mode selection | ⬜ Future — mode selection is still manual via `--stream` |
| 5. Parallel streaming | ⬜ Future |

The original phase designs are kept below as rationale and as the target shape for the
unshipped parts.

### Phase 1: Lazy Metadata Loading

**Status**: ⬜ Future. `LazyWorkbook` does not exist; workbook-level commands still load fully.

**Goal**: Load workbook structure without loading cell data.

```scala
trait LazyWorkbook:
  def sheetNames: Seq[SheetName]           // From workbook.xml
  def definedNames: Map[String, CellRange] // From workbook.xml
  def sheetMetadata(name: SheetName): SheetMetadata // dimensions, col widths
  def styleRegistry: StyleRegistry         // From styles.xml
  def streamSheet(name: SheetName): Stream[IO, RowData]
  def loadSheet(name: SheetName): IO[Sheet] // Full materialization
```

**Implementation**:
1. Parse `xl/workbook.xml` for sheet names and defined names
2. Parse `xl/styles.xml` for style registry (typically <1MB)
3. Parse sheet dimensions from `<dimension>` element without reading cells
4. Defer cell loading until explicitly requested

**Benefits**:
- `sheets` command becomes O(1)
- `names` command becomes O(1)
- Style-aware streaming possible for HTML/SVG

### Phase 2: Style-Aware Streaming

**Status**: ⚠️ Partial. Since #183, `RowData` carries `cellStyles: Map[Int, Int]` (raw style
indices) and `StyledRowData` feeds style-preserving streaming writes (`StylePatcher` patches
`styles.xml` during transforms). The rendering half — streaming HTML/SVG with resolved
`CellStyle`s — has not shipped; `view --stream` remains md/csv/json only.

**Goal**: Stream cells with style information for rich output formats.

```scala
case class StyledRowData(
  rowIndex: Int,
  cells: Map[Int, (CellValue, Option[CellStyle])]
)

def streamSheetStyled(name: SheetName): Stream[IO, StyledRowData]
```

**Implementation**:
1. Load style registry upfront (small, ~1MB for complex files)
2. Stream cells with style ID lookups
3. Emit styled rows for HTML/SVG rendering

**Benefits**:
- `view --format html` becomes streamable
- `view --format svg` becomes streamable
- PDF generation can stream pages

### Phase 3: Streaming Write Operations

**Status**: ✅ Mostly shipped (0.8.0, #183; worksheet-metadata patches extended in 0.9.6, #219).
The implementation matches the 3b design: `ZipTransformer` copies unchanged ZIP entries
verbatim, `StreamingTransform` SAX→StAX-rewrites only the target worksheet (cell values,
styles, plus `<cols>`, merges, and row attributes), and `StylePatcher` appends new styles to
`styles.xml`. Exposed as CLI `--stream` on `put`/`putf`/`style`/`batch`. From 3a, only
streaming CSV import shipped (`import --stream` to a new sheet via `writeStreamWithAutoDetect`);
generic append-to-existing-sheet APIs remain future work.

**Goal**: Enable streaming modifications without full workbook load.

#### 3a. Append-Only Operations
```scala
// Append rows to existing sheet (no read required)
def appendRows(path: Path, sheetName: String, rows: Stream[IO, RowData]): IO[Unit]

// Add new sheet with streaming data
def addStreamingSheet(path: Path, sheetName: String, rows: Stream[IO, RowData]): IO[Unit]
```

#### 3b. Range-Based Modifications
```scala
// Modify specific range without loading entire sheet
def modifyRange(
  path: Path,
  sheetName: String,
  range: CellRange,
  transform: Stream[IO, RowData] => Stream[IO, RowData]
): IO[Unit]
```

**Implementation**:
1. Use ZIP streaming to copy unmodified entries
2. Stream-transform target sheet XML
3. Write modified entries inline

**Benefits**:
- `put` on small range in huge file becomes O(range size)
- `import` can stream CSV directly to new sheet

### Phase 4: Intelligent Mode Selection

**Status**: ⬜ Future. No `FileProfile`/`SmartExecutor`; users choose with `--stream`.

**Goal**: Automatically choose optimal mode based on operation and file.

```scala
case class FileProfile(
  compressedSize: Long,
  estimatedRows: Long,      // From <dimension> element
  estimatedCells: Long,
  hasFormulas: Boolean,     // Detected from sheet XML scan
  hasStyles: Boolean,
  sheetCount: Int
)

def chooseMode(profile: FileProfile, command: CliCommand): ExecutionMode =
  command match
    case Search(_, limit, _) if limit < 1000 =>
      ExecutionMode.Streaming // Always stream search with limits

    case Stats(_) =>
      ExecutionMode.Streaming // Always stream aggregations

    case View(_, _, _, _, format, _) if format.isRasterizable =>
      if profile.estimatedCells > 100_000 then
        ExecutionMode.LazyStyled // Large file, stream with styles
      else
        ExecutionMode.InMemory // Small file, load everything

    case Put(ref, _) if profile.estimatedCells > 1_000_000 =>
      ExecutionMode.StreamingModify // Huge file, stream-modify

    case _ =>
      ExecutionMode.InMemory // Default safe mode
```

**Benefits**:
- No flags needed for most operations
- Optimal performance automatically
- Graceful degradation for complex cases

### Phase 5: Parallel Streaming

**Status**: ⬜ Future. No parallel multi-sheet streaming in xl-cats-effect or the CLI.

**Goal**: Utilize multiple cores for large file operations.

```scala
// Parallel search across sheets
def searchParallel(
  path: Path,
  pattern: Regex,
  sheets: Seq[SheetName]
): Stream[IO, SearchResult] =
  Stream.emits(sheets)
    .parEvalMap(maxConcurrency = 4)(sheet =>
      streamSheet(path, sheet).filter(matchesPattern)
    )
    .flatten

// Parallel aggregation with merge
def statsParallel(
  path: Path,
  range: CellRange,
  sheets: Seq[SheetName]
): IO[Stats] =
  sheets.parTraverse(sheet =>
    streamStats(path, sheet, range)
  ).map(_.combineAll)
```

**Benefits**:
- Multi-sheet search 4x faster
- Cross-sheet aggregations parallelized
- Better CPU utilization on large files

---

## Data Flow Diagrams

### Current: In-Memory Path
```
ZIP File → Parse All XML → Build Workbook → Execute Command → Output
           (O(n) memory)   (O(n) memory)
```

### Future: Smart Streaming Path
```
ZIP File → Probe Metadata → Choose Mode → Execute Optimal Path → Output
              (O(1))           │
                               ├─ Streaming: Parse XML Events → Process → Output
                               │                (O(1) memory)
                               │
                               ├─ Lazy Styled: Load Styles → Stream Cells → Output
                               │               (O(styles))   (O(1) per row)
                               │
                               └─ In-Memory: Full Load → Execute → Output
                                            (O(n) - only when necessary)
```

---

## Migration Path

(The original plan pinned phases to v0.9.0/v1.0.0/v1.1.0; delivery diverged — range
modifications shipped first, lazy metadata has not shipped. Restated against reality:)

### Shipped (0.8.0 – 0.11.0)
- [x] `--stream` flag for opt-in streaming
- [x] `--max-size` flag for large file override
- [x] Streaming reads: search, stats, bounds, view (md/csv/json), cell
- [x] Streaming range modifications: `put`/`putf`/`style`/`batch` via SAX→StAX transform (0.8.0, #183)
- [x] Streaming worksheet-metadata patches: col/row props, merges, visibility (0.9.6, #219)
- [x] Streaming CSV import to new sheet (`import --stream`)
- [x] SAX/StAX workbook writer for all other `--stream` writes (`writeWorkbookStream`)

### Future (unscheduled)
- [ ] Lazy workbook metadata loading (`sheets` and `names` without full load)
- [ ] Auto-streaming for search/stats (no flag needed)
- [ ] Style-aware streaming HTML/SVG for large ranges
- [ ] Intelligent mode selection
- [ ] Append-to-existing-sheet write operations
- [ ] Parallel multi-sheet operations

---

## API Changes Required

### xl-ooxml Module
```scala
// New: Metadata-only parsing
object WorkbookMetadataReader:
  def readMetadata(path: Path): IO[WorkbookMetadata]
  def readSheetDimensions(path: Path, sheet: SheetName): IO[CellRange]

// New: Streaming with styles
object StyledStreamingReader:
  def streamStyled(path: Path, sheet: SheetName, styles: StyleRegistry): Stream[IO, StyledRowData]
```

### xl-cats-effect Module
```scala
// New: Lazy workbook
trait LazyExcel[F[_]]:
  def probe(path: Path): F[FileProfile]
  def lazyRead(path: Path): F[LazyWorkbook]
  def streamStyled(path: Path, sheet: SheetName): Stream[F, StyledRowData]

// New: Streaming modifications
trait StreamingWriter[F[_]]:
  def appendSheet(path: Path, sheet: SheetName, rows: Stream[F, RowData]): F[Unit]
  def modifyRange(path: Path, sheet: SheetName, range: CellRange, f: RowData => RowData): F[Unit]
```

### xl-cli Module
```scala
// New: Intelligent execution
object SmartExecutor:
  def execute(profile: FileProfile, cmd: CliCommand): IO[String]
  def chooseMode(profile: FileProfile, cmd: CliCommand): ExecutionMode

enum ExecutionMode:
  case Streaming
  case LazyStyled
  case InMemory
  case StreamingModify
  case Parallel(concurrency: Int)
```

---

## Performance Targets

("Baseline" is the v0.8.0-era in-memory load time the targets were set against. For current
measured numbers on the shipped paths, see `docs/reference/performance-guide.md`.)

| Operation | File Size | Baseline | Target | Mode | Status |
|-----------|-----------|----------|--------|------|--------|
| `sheets` | 1M rows | 80s | <1s | Lazy metadata | ⬜ Future |
| `search --limit 10` | 1M rows | 80s | <1s | Streaming + early stop | ✅ Shipped |
| `view A1:D100 --format html` | 1M rows | 80s | <2s | Lazy styled | ⬜ Future |
| `put A1 "test"` | 1M rows | 80s | <5s | Streaming modify | ✅ Shipped |
| `search` (4 sheets) | 4x250k rows | 80s | 20s | Parallel streaming | ⬜ Future |

---

## Open Questions

1. **Style Registry Size**: For files with thousands of unique styles, should we stream styles too?

2. **Formula Dependencies**: Can we detect formula-free ranges and stream those while loading formula cells?

3. **Incremental Updates**: Should we support "delta" modifications that append to existing ZIP without rewriting?

4. **Memory Limits**: Should smart mode selection consider available JVM heap?

5. **Progress Reporting**: How to show progress for long-running streaming operations?

---

## References

- [Current Streaming Implementation](../../xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala)
- [SAX Parser](../../xl-cats-effect/src/com/tjclp/xl/io/SaxStreamingReader.scala)
- [CLI Commands](../../xl-cli/src/com/tjclp/xl/cli/commands/)
- [OOXML Spec: Worksheet](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/)
