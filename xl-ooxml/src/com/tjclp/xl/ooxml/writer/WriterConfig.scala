package com.tjclp.xl.ooxml.writer

import java.io.OutputStream
import java.nio.file.Path
import java.util.zip.ZipEntry

/** Shared Strings Table usage policy */
enum SstPolicy derives CanEqual:
  /** Auto-detect based on heuristics (default) */
  case Auto

  /** Always use SST regardless of content */
  case Always

  /** Never use SST (inline strings only) */
  case Never

/**
 * Compression method for ZIP entries.
 *
 * DEFLATED (default) produces 5-10x smaller files with minimal CPU overhead. STORED is useful for
 * debugging (human-readable ZIP contents).
 */
enum Compression derives CanEqual:
  /** No compression (STORED) - faster writes, larger files, requires CRC32 precomputation */
  case Stored

  /** DEFLATE compression (DEFLATED) - smaller files, standard production use */
  case Deflated

  /** ZIP constant for this compression method */
  def zipMethod: Int = this match
    case Stored => ZipEntry.STORED
    case Deflated => ZipEntry.DEFLATED

/**
 * Formula injection escaping policy for text cell values.
 *
 * When writing untrusted data to Excel, text starting with `=`, `+`, `-`, or `@` could be
 * interpreted as formulas. This policy controls whether such values are automatically escaped.
 */
enum FormulaInjectionPolicy derives CanEqual:
  /**
   * Automatically escape potentially dangerous text by prefixing with single quote.
   *
   * This is the safest default for untrusted data.
   */
  case Escape

  /**
   * Do not escape any values (trust all input).
   *
   * Only use this when writing data from trusted sources where formula characters are intentional.
   */
  case None

/**
 * XML serialization backend for XLSX output.
 *
 * SaxStax is 33% faster than ScalaXml but is newer and less battle-tested. ScalaXml remains the
 * default for stability until SaxStax has been proven stable in production use.
 *
 * Users can opt into SaxStax via CLI (`--backend saxstax`) or programmatically
 * (`WriterConfig.saxStax`).
 */
enum XmlBackend derives CanEqual:
  /** Stable backend using scala-xml. Default for production use. */
  case ScalaXml

  /** High-performance backend using StAX. 33% faster writes, ready for beta testing. */
  case SaxStax

/**
 * Writer configuration options.
 *
 * The default backend is ScalaXml for stability. SaxStax (33% faster) is available for beta testing
 * via `--backend saxstax` CLI flag or `WriterConfig.saxStax`. Once SaxStax has been proven stable
 * in production, it will become the default.
 */
case class WriterConfig(
  sstPolicy: SstPolicy = SstPolicy.Auto,
  compression: Compression = Compression.Deflated,
  prettyPrint: Boolean = false, // Compact XML for production (only applies to ScalaXml backend)
  backend: XmlBackend = XmlBackend.ScalaXml, // ScalaXml default for stability; SaxStax opt-in
  formulaInjectionPolicy: FormulaInjectionPolicy =
    FormulaInjectionPolicy.None // Default: trust input
)

object WriterConfig:
  /** Default production configuration: DEFLATED compression + ScalaXml backend (stable) */
  val default: WriterConfig = WriterConfig()

  /**
   * Secure configuration for untrusted data.
   *
   * Escapes formula injection characters (`=`, `+`, `-`, `@`) at the start of text values. Use this
   * when writing user-provided or external data that could contain malicious formulas.
   */
  val secure: WriterConfig = WriterConfig(
    formulaInjectionPolicy = FormulaInjectionPolicy.Escape
  )

  /**
   * Debug configuration: STORED compression + pretty XML for manual inspection.
   *
   * Uses ScalaXml backend since prettyPrint only works with ScalaXml (SaxStax always outputs
   * compact XML).
   */
  val debug: WriterConfig = WriterConfig(
    compression = Compression.Stored,
    prettyPrint = true,
    backend = XmlBackend.ScalaXml
  )

  /** ScalaXml backend configuration (same as default). */
  val scalaXml: WriterConfig = WriterConfig(backend = XmlBackend.ScalaXml)

  /**
   * SaxStax backend configuration for faster writes (beta).
   *
   * Use this for 33% faster writes when processing large files. ScalaXml is the default for
   * stability; opt into SaxStax explicitly when performance is critical. SaxStax is ready for beta
   * testing and will become the default once proven stable in production.
   */
  val saxStax: WriterConfig = WriterConfig(backend = XmlBackend.SaxStax)

/**
 * Target for XLSX output (file path or output stream).
 *
 * Enables unified handling of different output destinations in surgical modification. File-based
 * targets support verbatim copy optimization for clean workbooks.
 */
sealed trait OutputTarget:
  /** Get path if this is a file target (None for streams) */
  def asPathOption: Option[Path] = None

/** Output target that writes to a file path */
case class OutputPath(path: Path) extends OutputTarget:
  override def asPathOption: Option[Path] = Some(path)

/** Output target that writes to an output stream */
case class OutputStreamTarget(stream: OutputStream) extends OutputTarget
