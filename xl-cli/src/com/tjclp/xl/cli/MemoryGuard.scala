package com.tjclp.xl.cli

import java.nio.file.Path
import java.util.zip.ZipFile

import scala.jdk.CollectionConverters.*

import cats.effect.IO
import cats.syntax.all.*

import com.tjclp.xl.cli.contract.{CliError, CliException, ErrorCode, Location}
import com.tjclp.xl.error.XLResult
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.XlsxReader
import com.tjclp.xl.ooxml.XlsxReader.{ReadResult, ReaderConfig}
import com.tjclp.xl.workbooks.Workbook

/**
 * The heap as part of the contract (GH-636).
 *
 * `--max-size` lifts the reader's security limit on the uncompressed size; it does not make the
 * workbook fit. The native image bakes an 8 GB heap ceiling (`-R:MaxHeapSize=8g` in
 * `xl-cli/package.mill`, raised only by `-Xmx` as the first argument) and a JVM has whatever `-Xmx`
 * it was given, while an in-memory load of a million-row book needs tens of GB. Left alone, the
 * load ends in a raw `java.lang.OutOfMemoryError` on stderr, exit 1 and — under `--json` — no
 * envelope: two contract violations (ADR-017 §2.3).
 *
 * Two defences, one code:
 *   - [[blocking]] runs a thunk and turns an `OutOfMemoryError` raised inside it into the typed
 *     `RESOURCE_LIMIT` failure. This is the ONLY place the conversion can happen: cats-effect
 *     treats every `VirtualMachineError` as fatal — `IOFiber` never delivers one to `attempt` or
 *     `handleErrorWith`, it shuts the runtime down and `IOApp` halts with a stack trace and exit 1
 *     — so the last-resort handler in `Cli.run` cannot see it; the catch has to sit inside the
 *     thunk, before the runtime does. The failure is pre-built ([[exhausted]]) so the catch
 *     allocates nothing on a heap that has just run out.
 *   - [[admit]] refuses, before any byte is parsed, an in-memory load whose `--max-size` lifted the
 *     default and whose estimated footprint exceeds [[budget]] of the heap. The estimate is
 *     [[multiplier]] × the uncompressed worksheet and shared-string XML (sizes from the zip's
 *     central directory, no inflation): the 0.21.0 dogfood book carried 1.09 GB of sheet XML and
 *     needed 36–45 GB of heap (33–41×), so 30× against 70% of the heap refuses whenever the heap is
 *     below ~43× the XML — the observed need — and passes anything with real headroom. `-Xmx` is
 *     the override: the estimate is compared with the heap the process actually has, so a bigger
 *     heap admits a bigger book, and a load that then still exhausts memory is reported by
 *     [[blocking]] with the same code. Only memory exhaustion is classified this way; any other
 *     `Error` keeps its `INTERNAL` classification (and, being fatal to the runtime, its halt).
 *
 * [[excel]] is the `ExcelIO` the CLI loads workbooks through: `readWith` (and so `read`) runs
 * [[admit]] then the reader under [[blocking]], with the library's own warning routing and error
 * text (pinned equal by MemoryGuardSpec). Streaming reads never load a workbook and are untouched.
 */
object MemoryGuard:

  /**
   * The heap ceiling: `-Xmx` (or the native image's baked `MaxHeapSize`); `Long.MaxValue` = none.
   */
  val maxHeapBytes: Long = Runtime.getRuntime.maxMemory

  /** Heap bytes per byte of worksheet XML the in-memory model is estimated to need (see above). */
  val multiplier: Int = 30

  /** The share of the heap an estimated load may claim before it is refused. */
  val budget: Double = 0.7

  /** The next step, on both failures. */
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

  /** `--max-size` lifted the default (0 = unlimited, or above the default 100 MB). */
  def raised(config: ReaderConfig): Boolean =
    config.maxUncompressedSize == 0L ||
      config.maxUncompressedSize > ReaderConfig.default.maxUncompressedSize

  /**
   * The parts whose bytes drive the in-memory footprint: worksheets and the shared-string table.
   */
  def counted(entryName: String): Boolean =
    (entryName.startsWith("xl/worksheets/") && entryName.endsWith(".xml") &&
      !entryName.contains("/_rels/")) || entryName == "xl/sharedStrings.xml"

  /**
   * The uncompressed bytes of the [[counted]] parts, read from the zip's central directory —
   * O(entries), nothing inflated. Total: an unreadable input, a non-zip or an entry without a
   * recorded size counts as 0, so the guard never masks the read's own `IO_READ`.
   */
  def footprintBytes(path: Path): IO[Long] =
    IO.blocking {
      val zip = new ZipFile(path.toFile)
      try
        zip
          .entries()
          .asScala
          .filter(e => !e.isDirectory && counted(e.getName))
          .map(e => math.max(0L, e.getSize))
          .sum
      finally zip.close()
    }.handleError(_ => 0L)

  /** The refusal, or not, for `xmlBytes` of worksheet XML against a heap of `heap` bytes. */
  def decide(path: Path, xmlBytes: Long, heap: Long): Either[CliError, Unit] =
    val estimate = xmlBytes * multiplier
    val allowance = (heap * budget).toLong
    if heap == Long.MaxValue || xmlBytes <= 0L || estimate <= allowance then Right(())
    else
      Left(
        CliError(
          ErrorCode.RESOURCE_LIMIT,
          s"$path does not fit in memory: ${human(xmlBytes)} of worksheet XML needs an estimated " +
            s"${human(estimate)} of heap (${multiplier}× the XML), more than " +
            s"${(budget * 100).toInt}% of the ${human(heap)} heap",
          hint = Some(hint),
          location = Some(Location.file(path.toString))
        )
      )

  /**
   * Refuse the load up front when `config` lifted the default and the estimate cannot fit `heap`
   * (the process's, unless a test injects one). At the default limit the guard is inactive: 100 MB
   * of XML fits any heap the reader is otherwise allowed on.
   */
  def admit(path: Path, config: ReaderConfig, heap: Long = maxHeapBytes): IO[Unit] =
    if !raised(config) then IO.unit
    else
      footprintBytes(path).flatMap { bytes =>
        IO.fromEither(decide(path, bytes, heap).left.map(CliException(_)))
      }

  /**
   * The `ExcelIO` the CLI loads workbooks through: `readWith` runs [[admit]], then `parse` under
   * [[blocking]], then the library's own routing — every reader warning to `handler`, a reader
   * failure as `Failed to read XLSX: <message>` (the text `ExcelIO.readWith` produces, so
   * `classifyRead`'s `IO_READ` stays byte-identical). `parse` and `heap` are injection points for
   * the tests; production callers pass only the handler.
   */
  def excel(
    handler: XlsxReader.Warning => IO[Unit],
    parse: (Path, ReaderConfig) => XLResult[ReadResult] = XlsxReader.readWithWarnings(_, _),
    heap: Long = maxHeapBytes
  ): ExcelIO[IO] =
    new ExcelIO[IO](handler):
      override def readWith(path: Path, config: ReaderConfig): IO[Workbook] =
        admit(path, config, heap) *> blocking(parse(path, config)).flatMap {
          case Right(result) => result.warnings.traverse_(handler).as(result.workbook)
          case Left(err) => IO.raiseError(new Exception(s"Failed to read XLSX: ${err.message}"))
        }

  /** Binary units with one decimal (`8.0 GB`, `512.0 MB`), the way `-Xmx` counts. */
  def human(bytes: Long): String =
    val units = Vector("KB", "MB", "GB", "TB", "PB")
    if bytes < 1024L then s"$bytes B"
    else
      val exponent = math.min(units.size, (63 - java.lang.Long.numberOfLeadingZeros(bytes)) / 10)
      val value = bytes.toDouble / math.pow(1024d, exponent.toDouble)
      f"$value%.1f ${units(exponent - 1)}"
