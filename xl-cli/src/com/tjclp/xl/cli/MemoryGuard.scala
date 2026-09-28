package com.tjclp.xl.cli

import java.nio.file.Path
import java.util.concurrent.atomic.AtomicReference
import java.util.zip.ZipFile

import scala.jdk.CollectionConverters.*
import scala.util.Try
import scala.util.control.NoStackTrace

import cats.effect.IO
import cats.syntax.all.*
import fs2.{Pipe, Stream}

import com.tjclp.xl.cli.contract.{CliError, CliException, ErrorCode, Location, Warning, WarningCode}
import com.tjclp.xl.error.{XLError, XLException, XLResult}
import com.tjclp.xl.io.{ExcelIO, RowData, StyledRowData}
import com.tjclp.xl.ooxml.{WriterConfig, XlsxReader, XlsxWriter}
import com.tjclp.xl.ooxml.XlsxReader.{ReadResult, ReaderConfig}
import com.tjclp.xl.styles.CellStyle
import com.tjclp.xl.workbooks.Workbook

/**
 * The heap as part of the contract (GH-636).
 *
 * `--max-size` lifts the reader's security limit on the uncompressed size; it does not make the
 * workbook fit. The native image bakes an 8 GB heap ceiling (`-R:MaxHeapSize=8g` in
 * `xl-cli/package.mill`, raised by passing `-Xmx<size>` on the command line) and a JVM has whatever
 * `-Xmx` it was given, while an in-memory load of a million-row book of wide text rows needs tens
 * of GB. Left alone, the load ends in a raw `java.lang.OutOfMemoryError` on stderr, exit 1 and —
 * under `--json` — no envelope: two contract violations (ADR-017 §2.3).
 *
 * Three answers, one code:
 *   - [[blocking]] runs a thunk and turns an `OutOfMemoryError` raised inside it into the typed
 *     `RESOURCE_LIMIT` failure. This is the ONLY place the conversion can happen: cats-effect
 *     treats every `VirtualMachineError` as fatal — `IOFiber` never delivers one to `attempt` or
 *     `handleErrorWith`, it shuts the runtime down and `IOApp` halts with a stack trace and exit 1
 *     — so the last-resort handler in `Cli.run` cannot see it; the catch has to sit inside the
 *     thunk, before the runtime does. The failure is pre-built ([[exhausted]]) so the catch
 *     allocates nothing on a heap that has just run out.
 *   - [[admit]] sizes an in-memory load whose `--max-size` lifted the default BEFORE any byte is
 *     parsed, from the uncompressed worksheet and shared-string XML (central directory, no
 *     inflation), in two bands calibrated on measured loads (the table at [[sheetLower]]). A load
 *     whose LOWER estimate already exceeds the heap is hopeless and is REFUSED — `RESOURCE_LIMIT`,
 *     exit 3, the file in `error.location`. A load whose UPPER estimate exceeds the heap while the
 *     lower one does not MAY fit: it proceeds under a `MEMORY_PRESSURE` warning carrying both
 *     estimates and the hint, and if it does exhaust the heap [[blocking]] reports it with the same
 *     code. Below both the load is silent. A refusal is a failure an agent can only reason past
 *     with `-Xmx`, so the lower band is the floor of every measured shape; `-Xmx` is the override
 *     either way: the estimates are compared with the heap the process actually has.
 *
 * The catch covers every stage that builds a large structure, each inside its own thunk: the load
 * ([[excel]]'s `readWith`), the recalculation after an edit and `recalc` itself (WriteCommands),
 * the evaluation behind `view --eval` and the html/svg/raster renders (InMemorySource), and the
 * serialisation ([[excel]]'s `writeWith`, which `writeWorkbookStream` and `write` go through; every
 * command writes through the one [[writer]]). Residuals, stated so the promise is exact: an
 * `OutOfMemoryError` raised first on ANOTHER cats-effect thread — a `--parallel` recalculation
 * fiber, a timer — is still fatal to the runtime; a non-heap `OutOfMemoryError` (Metaspace, "unable
 * to create native thread") is reported with the heap wording; and a central directory that lies
 * about its sizes only bypasses the estimate — the load then lands in the typed catch. Only memory
 * exhaustion is classified this way; any other `Error` keeps its `INTERNAL` classification (and,
 * being fatal to the runtime, its halt). Streaming reads never load a workbook and are untouched.
 * [[excel]] keeps the library's own warning routing and error texts (pinned equal by
 * MemoryGuardSpec).
 */
object MemoryGuard:

  /**
   * The heap ceiling: `-Xmx` (or the native image's baked `MaxHeapSize`); `Long.MaxValue` = none.
   */
  val maxHeapBytes: Long = Runtime.getRuntime.maxMemory

  /**
   * The estimate is `sheet × S + sst × T`: heap bytes per byte of worksheet XML (`S`, dense cells —
   * every cell is an object graph many times the size of its `<c>` markup) and per byte of
   * `sharedStrings.xml` (`T`, text bodies — a string in memory is about its UTF-8 length). Each
   * coefficient has a LOWER value below which no measured load completed (the refusal band) and an
   * UPPER value above which every measured load completed (the warning band). Measured with the
   * reader itself (`XlsxReader.readWithWarnings`, permissive config, HotSpot 25, the smallest
   * `-Xmx` at which the load completes, G1 and ParallelGC), plus the 0.21.0 dogfood book:
   * {{{
   * book (1,000,000 rows)                sheet XML  SST XML  smallest heap that loads (largest that fails)
   *                                                          G1               ParallelGC
   * (a) 8 decimals, e.g. 80000.07        346 MiB    —        6144m (5120m)    8192m (7168m)
   * (c) 8 short integers                 309 MiB    —        6144m (5120m)    8192m (7168m)
   * (b) 8 columns, 4 of distinct text    342 MiB    106 MiB  7168m (6144m)    8192m (7168m)
   * (d) 1 distinct 300-character text     60 MiB    303 MiB  1792m (1536m)    1792m (1536m)
   * nyc1m (0.21.0 dogfood: wide text
   *   rows, ~1M shared strings)          1.09 GB    —        36–45 GB; died on the native image's 8g
   * }}}
   * Heap per byte of XML at the smallest loading heap: 17.8× (a), 19.9× (c), 16.0× (b), 4.9× (d)
   * and 33–41× for nyc1m; the largest failing heaps sit at 14.8×, 16.6×, 13.7× and 4.2×. The
   * sheet-only books pin `S`: nothing loaded below 14.8× its sheet XML, so [[sheetLower]] is 14,
   * and the dogfood's 33× rounds down to [[sheetUpper]] 30. With its 60 MiB of sheet XML charged at
   * 14–30×, (d)'s 303 MiB of strings needed 2.3× (the 1536m failure) to 3× of their own bytes:
   * [[sstLower]] 2, [[sstUpper]] 3. Every measured book's lower estimate is below the heap it
   * loaded in — (a) 4.7 GiB ≤ 6, (c) 4.2 ≤ 6, (b) 4.9 ≤ 7, (d) 1.4 ≤ 1.75 — so none can be refused
   * where it loads (MemoryGuardSpec holds each shape to its band), while nyc1m's 14.2 GiB lower
   * estimate refuses it on the native image's 8 GB. Under the upper band every measured load but
   * nyc1m's completes (it needed up to 41×, the residual the typed `RESOURCE_LIMIT` covers).
   */
  val sheetLower: Int = 14
  val sheetUpper: Int = 30
  val sstLower: Int = 2
  val sstUpper: Int = 3

  /** The next step, on every answer but silence. */
  val hint: String =
    "use --stream for constant-memory reads, or raise the heap with -Xmx<size> " +
      "(native image: xl -Xmx64g …)"

  /** The failure a caught `OutOfMemoryError` reports. Built once, at class initialisation. */
  val exhausted: CliError =
    CliError(
      ErrorCode.RESOURCE_LIMIT,
      s"out of memory: the ${human(maxHeapBytes)} heap is exhausted",
      hint = Some(hint)
    )

  /** Pre-built so the catch in [[blocking]] allocates nothing. */
  private val refusal: Either[Throwable, Nothing] = Left(CliException(exhausted))

  /**
   * Run `thunk` on the blocking pool; an `OutOfMemoryError` it raises becomes the `RESOURCE_LIMIT`
   * [[contract.CliException]] (through the normal error channel, so `attempt`, the run's
   * classification and the envelope all see it). Anything else — a value, a non-fatal exception,
   * another `Error` — passes through exactly as `IO.blocking` would pass it.
   */
  def blocking[A](thunk: => A): IO[A] =
    // Evaluated here, not in the catch: touching the object initialises `exhausted` and `refusal`
    // while there is still memory to do so.
    val refused: Either[Throwable, A] = refusal
    IO.blocking {
      try Right(thunk)
      catch case _: OutOfMemoryError => refused
    }.rethrow

  /**
   * Test seams (`private[cli]`): the heap an input is sized against, keyed by file so a suite's
   * fixture never touches another run; and a fault raised before a guarded stage runs, keyed by
   * `(Some(file), stage)` — or `(None, stage)` for a stage whose file the test cannot know, such as
   * a write's recalculation (the commands see the staging path, not `-o`). Production finds them
   * empty.
   */
  private[cli] object Seams:
    val heap: AtomicReference[Map[Path, Long]] = new AtomicReference(Map.empty)
    val faults: AtomicReference[Map[(Option[Path], String), Throwable]] =
      new AtomicReference(Map.empty)

    /** The fault pinned for this file and stage, or for the stage alone. */
    def faultFor(path: Path, stage: String): Option[Throwable] =
      val pinned = faults.get
      pinned.get((Some(path), stage)).orElse(pinned.get((None, stage)))

  /** The heap `path` is sized against: the process's, unless a test pinned one for that file. */
  private[cli] def heapFor(path: Path): Long = Seams.heap.get.getOrElse(path, maxHeapBytes)

  /**
   * [[blocking]] for a named stage of one file's processing — `load`, `recalc`, `write` — where a
   * test can inject a fault ([[Seams]]).
   */
  def blocking[A](path: Path, stage: String)(thunk: => A): IO[A] =
    val refused: Either[Throwable, A] = refusal
    IO.blocking {
      try
        Seams.faultFor(path, stage).foreach(fault => throw fault)
        Right(thunk)
      catch case _: OutOfMemoryError => refused
    }.rethrow

  /**
   * `--max-size` lifted the default: 0 = unlimited, above the default 100 MB, or a negative value
   * (the parser refuses one, but the reader treats `<= 0` as no limit, so the guard does too).
   */
  def raised(config: ReaderConfig): Boolean =
    config.maxUncompressedSize <= 0L ||
      config.maxUncompressedSize > ReaderConfig.default.maxUncompressedSize

  /** The uncompressed bytes that drive the footprint: worksheet XML and the shared-string table. */
  final case class Footprint(sheetBytes: Long, sstBytes: Long) derives CanEqual:
    def total: Long = sheetBytes + sstBytes
    def atLeast: Long = sheetBytes * sheetLower + sstBytes * sstLower
    def upTo: Long = sheetBytes * sheetUpper + sstBytes * sstUpper

  object Footprint:
    val empty: Footprint = Footprint(0L, 0L)

  /** A worksheet part (not its rels). */
  def isWorksheet(entryName: String): Boolean =
    entryName.startsWith("xl/worksheets/") && entryName.endsWith(".xml") &&
      !entryName.contains("/_rels/")

  /** The shared-string table. */
  def isSharedStrings(entryName: String): Boolean = entryName == "xl/sharedStrings.xml"

  /** The parts whose bytes drive the in-memory footprint. */
  def counted(entryName: String): Boolean = isWorksheet(entryName) || isSharedStrings(entryName)

  /**
   * The uncompressed bytes of the [[counted]] parts, read from the zip's central directory —
   * O(entries), nothing inflated. Total: an unreadable input, a non-zip or an entry without a
   * recorded size counts as 0, so the guard never masks the read's own `IO_READ`.
   */
  def footprint(path: Path): IO[Footprint] =
    IO.blocking {
      val zip = new ZipFile(path.toFile)
      try
        zip
          .entries()
          .asScala
          .filter(e => !e.isDirectory && counted(e.getName))
          .foldLeft(Footprint.empty) { (acc, e) =>
            val size = math.max(0L, e.getSize)
            if isSharedStrings(e.getName) then acc.copy(sstBytes = acc.sstBytes + size)
            else acc.copy(sheetBytes = acc.sheetBytes + size)
          }
      finally zip.close()
    }.handleError(_ => Footprint.empty)

  /** [[footprint]]'s total: worksheet plus shared-string bytes. */
  def footprintBytes(path: Path): IO[Long] = footprint(path).map(_.total)

  /**
   * What the pre-load estimate says about a load: silence, a warning it rides with, or a refusal.
   */
  enum Verdict derives CanEqual:
    case Admit
    case Warn(warning: Warning)
    case Refuse(error: CliError)

  /** The verdict for a [[Footprint]] against a heap of `heap` bytes. */
  def decide(path: Path, fp: Footprint, heap: Long): Verdict =
    decide(path.toString, Location.file(path.toString), fp, heap)

  /** [[decide]] for a load named by `label` (one file, or `diff`'s two) at `location`. */
  def decide(label: String, location: Location, fp: Footprint, heap: Long): Verdict =
    val atLeast = fp.atLeast
    val upTo = fp.upTo
    // Name the shared-string share only when it is material (an empty table is ~150 bytes)
    def xml =
      if fp.sstBytes >= (1L << 20) then
        s"${human(fp.total)} of worksheet and shared-string XML (${human(fp.sstBytes)} of it " +
          "shared strings)"
      else s"${human(fp.total)} of worksheet XML"
    if heap == Long.MaxValue || fp.total <= 0L then Verdict.Admit
    else if atLeast > heap then
      Verdict.Refuse(
        CliError(
          ErrorCode.RESOURCE_LIMIT,
          s"$label does not fit in memory: $xml needs at least ${human(atLeast)} of heap " +
            s"(${sheetLower}× the sheet XML, ${sstLower}× the strings; up to ${human(upTo)} for " +
            s"wide text rows), more than the ${human(heap)} heap",
          hint = Some(hint),
          location = Some(location)
        )
      )
    else if upTo > heap then
      Verdict.Warn(
        Warning(
          WarningCode.MEMORY_PRESSURE,
          s"$label may not fit in memory: $xml needs an estimated ${human(atLeast)}–${human(upTo)} " +
            s"of heap against the ${human(heap)} heap; if the load fails with RESOURCE_LIMIT, $hint",
          Some(location)
        )
      )
    else Verdict.Admit

  /** [[decide]] for `xmlBytes` of worksheet XML and no shared strings. */
  def decide(path: Path, xmlBytes: Long, heap: Long): Verdict =
    decide(path, Footprint(xmlBytes, 0L), heap)

  /**
   * Apply the verdict before a load: when `config` lifted the default, size the file and refuse the
   * hopeless case, announce the doubtful one through `warn`, pass the rest silently. At the default
   * limit the guard is inactive: 100 MB of XML fits any heap the reader is otherwise allowed on.
   * `heap` is the process's unless a test injects one.
   */
  def admit(
    path: Path,
    config: ReaderConfig,
    heap: Long = maxHeapBytes,
    warn: Warning => IO[Unit] = _ => IO.unit
  ): IO[Unit] =
    if !raised(config) then IO.unit
    else footprint(path).flatMap(fp => act(decide(path, fp, heap), warn))

  /**
   * [[admit]] for files loaded TOGETHER (`diff`): one verdict on the sum of their footprints, since
   * both books are resident at once; the message names every file, `location.file` the first.
   */
  def admitAll(
    paths: Vector[Path],
    config: ReaderConfig,
    heap: Long,
    warn: Warning => IO[Unit]
  ): IO[Unit] =
    if !raised(config) || paths.isEmpty then IO.unit
    else
      paths.traverse(footprint).flatMap { fps =>
        val sum = fps.foldLeft(Footprint.empty) { (a, b) =>
          Footprint(a.sheetBytes + b.sheetBytes, a.sstBytes + b.sstBytes)
        }
        // The message names every file; `location.file` stays ONE path (consumers treat it as one)
        val label = paths.mkString(" + ")
        val location = Location.file(paths.headOption.fold("")(_.toString))
        act(decide(label, location, sum, heap), warn)
      }

  private def act(verdict: Verdict, warn: Warning => IO[Unit]): IO[Unit] = verdict match
    case Verdict.Admit => IO.unit
    case Verdict.Warn(warning) => warn(warning)
    case Verdict.Refuse(error) => IO.raiseError(CliException(error))

  /**
   * The `ExcelIO` the CLI loads workbooks through: `readWith` runs [[admit]] (its warning to
   * `warn`), then `parse` under [[blocking]], then the library's own routing — every reader warning
   * to `handler`, a reader failure as the `XLException` `ExcelIO.readWith` raises (GH-621: the
   * error's own message, one prefix — `classifyRead` projects it to `IO_READ`) — except the
   * reader's own refusal: a `SecurityError` (the `--max-size` limit, a ZIP bomb) is not an
   * unreadable file but a policy, so it keeps its code, `SECURITY_ERROR`, its message and its
   * `--max-size 0 or --stream` hint (GH-638). `parse` and `heap` are injection points for the
   * tests; production callers pass the two sinks.
   */
  def excel(
    handler: XlsxReader.Warning => IO[Unit],
    warn: Warning => IO[Unit] = _ => IO.unit,
    parse: (Path, ReaderConfig) => XLResult[ReadResult] = XlsxReader.readWithWarnings(_, _),
    heap: Long = maxHeapBytes,
    admitLoads: Boolean = true,
    spill: Either[CliError, Option[Path]] = Right(None)
  ): ExcelIO[IO] =
    new ExcelIO[IO](handler):
      override def readWith(path: Path, config: ReaderConfig): IO[Workbook] =
        (if admitLoads then admit(path, config, heap, warn) else IO.unit) *>
          blocking(path, "load")(parse(path, config)).flatMap {
            case Right(result) => result.warnings.traverse_(handler).as(result.workbook)
            case Left(security: XLError.SecurityError) =>
              IO.raiseError(
                CliException(CliError.fromXLError(security, Some(Location.file(path.toString))))
              )
            case Left(err) => IO.raiseError(XLException(err))
          }

      // Serialisation builds every part's XML; `writeWorkbookStream` and `write` route through
      // here, so one override covers the three. Same failure text as `ExcelIO.writeWith`.
      override def writeWith(wb: Workbook, path: Path, config: WriterConfig): IO[Unit] =
        blocking(path, "write")(XlsxWriter.writeWith(wb, path, config)).flatMap {
          case Right(_) => IO.unit
          case Left(err) => IO.raiseError(new Exception(s"Failed to write XLSX: ${err.message}"))
        }

      // GH-517: the spill directory is an override, never `withSpillDir` — that rebuilds a PLAIN
      // ExcelIO and would silently drop the two guards above. A malformed setting is not a
      // directory; the spilling writes below refuse it before they consult this.
      override def spillDir: Option[Path] = spill.getOrElse(None)

      // The spilling writers, and only they, answer for the setting: a malformed value is the
      // usage error before a row is pulled or a byte written; a failure of the write itself — the
      // scratch file, the archive — is IO_WRITE naming the target (the library's message names the
      // configured directory); a failure of the caller's own rows keeps its classification.
      private val gate: IO[Unit] = IO.fromEither(spill.leftMap(CliException(_))).void

      override def writeStreamWithAutoDetect(
        path: Path,
        sheetName: String,
        sheetIndex: Int,
        config: WriterConfig
      ): Pipe[IO, RowData, Unit] =
        rows =>
          Stream.exec(gate) ++
            super
              .writeStreamWithAutoDetect(path, sheetName, sheetIndex, config)(
                SourceFailure.tag(rows)
              )
              .handleErrorWith(e => Stream.raiseError[IO](writeFailure(path, spillDir, e)))

      override def writeStreamStyledWithAutoDetect(
        path: Path,
        sheetName: String,
        styles: Vector[CellStyle],
        sheetIndex: Int,
        config: WriterConfig
      ): Pipe[IO, StyledRowData, Unit] =
        rows =>
          Stream.exec(gate) ++
            super
              .writeStreamStyledWithAutoDetect(path, sheetName, styles, sheetIndex, config)(
                SourceFailure.tag(rows)
              )
              .handleErrorWith(e => Stream.raiseError[IO](writeFailure(path, spillDir, e)))

      override def writeStreamsSeqWithAutoDetect(
        path: Path,
        sheets: Seq[(String, Stream[IO, RowData])],
        config: WriterConfig
      ): IO[Unit] =
        gate *> super
          .writeStreamsSeqWithAutoDetect(
            path,
            sheets.map { case (name, rows) => name -> SourceFailure.tag(rows) },
            config
          )
          .adaptError { case e => writeFailure(path, spillDir, e) }

  /** The environment variable that redirects the streaming scratch file (GH-517). */
  val SpillDirVar: String = "XL_SPILL_DIR"

  /**
   * A failure of the rows a spilling write consumes, tagged on the way in so that it is not
   * classified as the write's on the way out ([[writeFailure]] unwraps it).
   */
  private final class SourceFailure(val cause: Throwable) extends Exception(cause) with NoStackTrace

  private object SourceFailure:
    def tag[A](rows: Stream[IO, A]): Stream[IO, A] =
      rows.handleErrorWith(e => Stream.raiseError[IO](new SourceFailure(e)))

  /**
   * The classification of what a spilling write raised (GH-517): the caller's own row failure,
   * unwrapped; a typed failure, as is; anything else — the scratch file that could not be created
   * or filled, the archive that could not be assembled — `IO_WRITE` naming `target`, with the
   * cause's message (the library's names the configured spill directory) and a hint naming the
   * lever, in the shape of the CLI's other write failures.
   */
  private def writeFailure(target: Path, spill: Option[Path], failure: Throwable): Throwable =
    failure match
      case source: SourceFailure => source.cause
      case cli: CliException => cli
      case other =>
        CliException(
          CliError(
            ErrorCode.IO_WRITE,
            s"cannot write $target: ${CliError.messageOf(other)}",
            hint = Some(
              s"check that the output directory and ${spillWhere(spill)} exist, are writable and have room"
            ),
            location = Some(Location.file(target.toString))
          )
        )

  /**
   * Where a scratch file lands, for a diagnostic: the configured `XL_SPILL_DIR`, or the default
   * `java.io.tmpdir` with the lever that moves it.
   */
  def spillWhere(spill: Option[Path]): String =
    spill.fold(
      s"the default temp directory java.io.tmpdir ($SpillDirVar=<dir> redirects the scratch file)"
    )(dir => s"$SpillDirVar ($dir)")

  /**
   * The spill directory from `XL_SPILL_DIR` (GH-517): trimmed; unset or blank is `None` — the JVM's
   * `java.io.tmpdir`, the library's default; a value that is not a path is a usage error (exit 2)
   * naming the variable. `env` is injectable so the parse is unit-testable without the process
   * environment (as `NativeImage.detect` reads its properties).
   */
  private[cli] def spillDirFrom(env: String => Option[String]): Either[CliError, Option[Path]] =
    env(SpillDirVar).map(_.trim).filter(_.nonEmpty) match
      case None => Right(None)
      case Some(raw) =>
        Try(Path.of(raw)).toEither match
          case Right(dir) => Right(Some(dir))
          case Left(e) =>
            Left(
              CliError.usage(
                s"$SpillDirVar is not a path: '$raw' (${CliError.messageOf(e)})",
                Some(s"unset $SpillDirVar or point it at an existing, writable directory")
              )
            )

  /**
   * `XL_SPILL_DIR`, read once per process — the application layer's one ambient read, so the
   * library stays deterministic (`ExcelIO.instance` never consults the environment). Consulted by
   * the spilling writes (the two-pass CSV import into a new workbook,
   * `ImportCommands.importToNewSheetStreaming`) and by the one render backend that needs a scratch
   * file (`Resvg.scratchSvg`: resvg takes file paths only); every other write never spills.
   */
  val spill: Either[CliError, Option[Path]] = spillDirFrom(sys.env.get)

  /**
   * The one `ExcelIO` every command writes through (no reader handlers: the commands receive their
   * workbook already loaded), so serialisation is under the guard wherever it happens — and the
   * spill goes where `XL_SPILL_DIR` points.
   */
  val writer: ExcelIO[IO] = excel(_ => IO.unit, spill = spill)

  /** Binary units with one decimal (`8.0 GB`, `512.0 MB`), the way `-Xmx` counts. */
  def human(bytes: Long): String =
    val units = Vector("KB", "MB", "GB", "TB", "PB")
    if bytes < 1024L then s"$bytes B"
    else
      val exponent = math.min(units.size, (63 - java.lang.Long.numberOfLeadingZeros(bytes)) / 10)
      val value = bytes.toDouble / math.pow(1024d, exponent.toDouble)
      f"$value%.1f ${units(exponent - 1)}"
