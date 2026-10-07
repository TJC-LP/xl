package com.tjclp.xl.cli.raster

import java.io.{IOException, OutputStream}
import java.nio.file.{
  AtomicMoveNotSupportedException,
  Files,
  LinkOption,
  Path,
  StandardCopyOption,
  StandardOpenOption
}
import java.nio.file.attribute.{
  AclEntry,
  AclEntryPermission,
  AclEntryType,
  AclFileAttributeView,
  PosixFilePermissions
}
import java.util.UUID

import scala.concurrent.duration.FiniteDuration
import scala.util.Using

import cats.effect.{IO, Resource}
import cats.syntax.all.*

/**
 * Common interface for SVG-to-raster converters.
 *
 * Batik (pure JVM, bundled) is the default backend. The subprocess backends remain automatic
 * fallbacks for environments where AWT is unavailable (native image):
 *   1. Batik (pure JVM, requires AWT - the default)
 *   2. cairosvg (Python, very portable - pip install cairosvg)
 *   3. rsvg-convert (librsvg, fast - apt install librsvg2-bin)
 *   4. resvg (Rust, best quality - cargo install resvg)
 *
 * ImageMagick is not part of the automatic chain (its SVG delegate configuration is fragile,
 * GH-83/GH-86) but stays reachable explicitly via `--rasterizer imagemagick`.
 */
trait Rasterizer:

  /** Human-readable name for this rasterizer (e.g., "Batik", "cairosvg") */
  def name: String

  /** Check if this rasterizer is available on the current system */
  def isAvailable: IO[Boolean]

  /**
   * Convert SVG content to a raster image.
   *
   * @param svg
   *   SVG content as a string
   * @param outputPath
   *   Destination file path
   * @param format
   *   Target format (PNG, JPEG, WebP, PDF)
   * @param dpi
   *   Resolution in DPI (default 144 for retina displays)
   * @return
   *   IO that completes when conversion is done, or fails with RasterError
   */
  def convertSvgToRaster(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int = 144
  ): IO[Unit]

/**
 * Output format for raster images.
 */
enum RasterFormat:
  case Png
  case Jpeg(quality: Int)
  case WebP
  case Pdf

  def extension: String = this match
    case Png => "png"
    case Jpeg(_) => "jpeg"
    case WebP => "webp"
    case Pdf => "pdf"

  def formatArg: String = this match
    case Png => "png"
    case Jpeg(_) => "jpeg"
    case WebP => "webp"
    case Pdf => "pdf"

/**
 * GraalVM native-image runtime detection, used to explain why the bundled Batik backend (which
 * needs AWT) can never work in the shipped native binary.
 *
 * Mirrors `org.graalvm.nativeimage.ImageInfo.inImageRuntimeCode()` without a compile-time GraalVM
 * dependency: ImageInfo itself answers by reading the `org.graalvm.nativeimage.imagecode` system
 * property, which Substrate VM sets to "runtime" inside every native binary. A reflective
 * `Class.forName("org.graalvm.nativeimage.ImageInfo")` would be strictly less reliable here - the
 * class is absent from JVM classpaths and may be unregistered for reflection inside the image - so
 * the property is read directly. `java.vm.name` containing "Substrate VM" is kept as a fallback for
 * older GraalVM releases.
 */
object NativeImage:

  /** True when running inside a GraalVM native binary (never true on a regular JVM). */
  lazy val inNativeImage: Boolean = detect(sys.props.get)

  /** Detection core with injectable properties so both environments are unit-testable. */
  private[cli] def detect(prop: String => Option[String]): Boolean =
    prop("org.graalvm.nativeimage.imagecode").exists(_.equalsIgnoreCase("runtime")) ||
      prop("java.vm.name").exists(_.contains("Substrate"))

/**
 * Errors that can occur during rasterization.
 */
sealed trait RasterError extends Exception:
  def message: String
  override def getMessage: String = message

object RasterError:
  /** No rasterizer available on the system */
  case class NoRasterizerAvailable(
    triedRasterizers: List[String],
    runningNativeImage: Boolean = NativeImage.inNativeImage
  ) extends RasterError:
    def message: String =
      val platformNote =
        if runningNativeImage then
          "Running as a native binary: the bundled Batik backend needs AWT and is\n" +
            "unavailable by design - install one of the external tools below."
        else
          "Note: the JAR distribution rasterizes out-of-the-box (bundled Batik); the\n" +
            "native binary always needs one of the external tools below."
      s"""No SVG rasterizer available. Tried: ${triedRasterizers.mkString(", ")}
         |
         |$platformNote
         |
         |Run `xl rasterizers` to see backend status on this machine.
         |
         |Install one of:
         |  - cairosvg: pip install cairosvg
         |  - rsvg-convert: apt install librsvg2-bin (Debian/Ubuntu)
         |  - resvg: cargo install resvg, or a prebuilt binary from
         |    https://github.com/linebender/resvg/releases
         |  - ImageMagick (not tried automatically): apt install imagemagick,
         |    then pass --rasterizer imagemagick
         |
         |Or use --format svg for vector output.""".stripMargin

  /** Specific rasterizer requested but not available */
  case class RasterizerNotFound(name: String, installHint: String) extends RasterError:
    def message: String =
      s"Rasterizer '$name' not found. $installHint (run `xl rasterizers` for backend status)"

  /** Format not supported by this rasterizer */
  case class FormatNotSupported(rasterizer: String, format: RasterFormat) extends RasterError:
    def message: String = s"$rasterizer does not support ${format.extension} format"

  /** Conversion failed; an empty stderr ends the message at the exit code (GH-690). */
  case class ConversionFailed(rasterizer: String, stderr: String, exitCode: Int)
      extends RasterError:
    def message: String =
      val detail = stderr.strip
      val base = s"$rasterizer conversion failed (exit $exitCode)"
      if detail.isEmpty then base else s"$base: $detail"

  /** GH-690: the backend ran past its deadline and was stopped. */
  case class TimedOut(rasterizer: String, after: FiniteDuration) extends RasterError:
    def message: String =
      s"$rasterizer did not finish within ${after.toSeconds} s and was stopped"

  /**
   * GH-690: the converted image could not be moved onto the output path (a read-only directory, a
   * path that is a directory). The CLI classifies it as `IO_WRITE`.
   */
  case class OutputFailed(path: Path, cause: Throwable) extends RasterError:
    def message: String =
      s"cannot write $path: ${Option(cause.getMessage).getOrElse(cause.toString)}"

  /**
   * GH-690: the output path names no regular file to write an image to: a root (`/`), a directory,
   * a FIFO or a device (`/dev/stdout`). Refused before any backend runs; the CLI classifies it as
   * `IO_WRITE`.
   */
  case class UnsupportedOutput(path: Path, reason: String) extends RasterError:
    def message: String = s"cannot write $path: $reason"

  /**
   * A backend that takes file paths only (resvg) could not create its scratch SVG: in `spillDir`
   * (the CLI's `XL_SPILL_DIR`) when set, else the JVM's `java.io.tmpdir`. The CLI classifies it as
   * `IO_WRITE` naming the directory and the lever (`CliError.fromThrowable`), not `INTERNAL`.
   */
  case class ScratchFileFailed(rasterizer: String, spillDir: Option[Path], cause: Throwable)
      extends RasterError:
    def message: String =
      s"$rasterizer could not write its scratch SVG: ${Option(cause.getMessage).getOrElse(cause.toString)}"

/**
 * Manages the fallback chain of rasterizers.
 */
object RasterizerChain:

  /** Standard screen DPI (CSS reference pixel) */
  private val BaseDpi = 96

  /**
   * Default fallback chain in priority order.
   *
   * Order rationale:
   *   - Batik: the default backend - pure JVM, bundled, no subprocess (covers PNG/JPEG); only
   *     unavailable where AWT is missing (native image)
   *   - cairosvg: Very portable (Python), pre-installed in Claude.ai
   *   - rsvg-convert: Fast (C/Rust), common on Linux
   *   - resvg: Best SVG 2.0 support, but requires Rust/Cargo
   *
   * ImageMagick is deliberately excluded: its SVG support depends on delegate configuration (rsvg)
   * that breaks in hard-to-diagnose ways (GH-83/GH-86). It remains available as an explicit opt-in
   * via `--rasterizer imagemagick`.
   */
  def defaultChain: List[Rasterizer] = PlatformRasterizers.bundled ::: List(
    CairoSvg,
    RsvgConvert,
    Resvg
  )

  /** Every known backend, including explicit-opt-in-only ones (ImageMagick). */
  def allRasterizers: List[Rasterizer] = defaultChain :+ ImageMagick

  /** Map of rasterizer names for --rasterizer flag */
  def byName: Map[String, Rasterizer] = allRasterizers.map(r => r.name.toLowerCase -> r).toMap

  /** Valid rasterizer names for CLI help */
  def validNames: List[String] = allRasterizers.map(_.name.toLowerCase)

  /**
   * Convert SVG to raster using the fallback chain.
   *
   * The SVG is pre-scaled based on DPI to ensure consistent output dimensions across all
   * rasterizers. Without this, Batik and ImageMagick treat DPI as metadata-only while rsvg-convert,
   * CairoSvg, and Resvg scale output dimensions. By scaling the SVG's width/height attributes
   * upfront, all rasterizers produce the expected larger pixel output at higher DPI.
   *
   * @param svg
   *   SVG content
   * @param outputPath
   *   Destination file
   * @param format
   *   Output format
   * @param dpi
   *   Resolution (higher DPI = larger pixel dimensions)
   * @param preferredRasterizer
   *   Optional name to force a specific rasterizer
   * @return
   *   IO containing the name of the rasterizer that succeeded
   */
  def convert(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int = 144,
    preferredRasterizer: Option[String] = None
  ): IO[String] =
    // Scale SVG dimensions based on DPI for consistent output across all rasterizers
    val scaledSvg = scaleSvgForDpi(svg, dpi)

    preferredRasterizer match
      case Some(name) =>
        // User requested specific rasterizer
        byName.get(name.toLowerCase) match
          case Some(rasterizer) =>
            rasterizer.isAvailable.flatMap {
              case true =>
                attempt(rasterizer, scaledSvg, outputPath, format, dpi).as(rasterizer.name)
              case false =>
                IO.raiseError(
                  RasterError.RasterizerNotFound(
                    name,
                    installHintFor(name.toLowerCase)
                  )
                )
            }
          case None =>
            IO.raiseError(
              RasterError.RasterizerNotFound(
                name,
                s"Valid options: ${validNames.mkString(", ")}"
              )
            )

      case None =>
        // Try fallback chain
        tryChain(scaledSvg, outputPath, format, dpi, defaultChain, Nil)

  /**
   * Try each rasterizer in the chain until one succeeds.
   */
  private def tryChain(
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int,
    remaining: List[Rasterizer],
    tried: List[String]
  ): IO[String] =
    remaining match
      case Nil =>
        IO.raiseError(RasterError.NoRasterizerAvailable(tried))

      case rasterizer :: rest =>
        rasterizer.isAvailable.flatMap {
          case false =>
            // Not available, try next
            tryChain(svg, outputPath, format, dpi, rest, tried :+ rasterizer.name)

          case true =>
            // Available, try to convert
            attempt(rasterizer, svg, outputPath, format, dpi)
              .map { _ =>
                // Log if we fell back from the default
                if tried.nonEmpty then
                  System.err.println(
                    s"Warning: ${tried.mkString(", ")} unavailable, using ${rasterizer.name}"
                  )
                rasterizer.name
              }
              .handleErrorWith { error =>
                // Conversion failed, try next (unless it's a format error)
                error match
                  // GH-690: a backend rendered but the image could not replace the output path;
                  // another backend cannot repair the destination
                  case output: RasterError.OutputFailed => IO.raiseError(output)
                  case output: RasterError.UnsupportedOutput => IO.raiseError(output)
                  case _: RasterError.FormatNotSupported =>
                    // Format not supported by this rasterizer, try next
                    tryChain(
                      svg,
                      outputPath,
                      format,
                      dpi,
                      rest,
                      tried :+ s"${rasterizer.name} (format not supported)"
                    )
                  case _ =>
                    // Other error (likely AWT/native image issue), try next
                    tryChain(
                      svg,
                      outputPath,
                      format,
                      dpi,
                      rest,
                      // the whole message: ConversionFailed already keeps only stderr's tail
                      tried :+ s"${rasterizer.name} (${error.getMessage})"
                    )
              }
        }

  /**
   * One backend's conversion, trusted only by its result (GH-690): the backend writes a hidden file
   * beside `outputPath`, which must exist and be non-empty before it replaces `outputPath`. A
   * backend that exits 0 without writing is [[RasterError.ConversionFailed]] (the chain then tries
   * the next one); a file left at `outputPath` by an earlier run cannot pass for this run's output,
   * and a failed run leaves it untouched. The hidden file is removed on every path out.
   *
   * The staging name is short and fixed-length (`.xl-raster-<uuid>.<ext>`), so any output name the
   * file system accepts still fits. It is created before the backend runs, so an output directory
   * that is missing or read-only is [[RasterError.OutputFailed]] up front rather than every
   * backend's failure; it is owner-only when an output already exists, since it will hold that
   * output's next contents. An output path that names no regular file is
   * [[RasterError.UnsupportedOutput]], also before any backend runs. [[publish]] says how the image
   * then replaces the output.
   */
  private def attempt(
    rasterizer: Rasterizer,
    svg: String,
    outputPath: Path,
    format: RasterFormat,
    dpi: Int
  ): IO[Unit] =
    val target = outputPath.toAbsolutePath
    val fileName: IO[String] = Option(target.getFileName) match
      case None => IO.raiseError(RasterError.UnsupportedOutput(target, "not a file path"))
      case Some(name) =>
        IO.blocking(Files.exists(target) && !Files.isRegularFile(target)).flatMap {
          case true => IO.raiseError(RasterError.UnsupportedOutput(target, "not a regular file"))
          case false => IO.pure(name.toString)
        }
    val staged = fileName.flatMap { name =>
      IO.blocking {
        val ext = name.lastIndexOf('.') match
          case i if i >= 0 && (1 to 16).contains(name.length - i - 1) => name.substring(i + 1)
          case _ => format.extension
        val staging = sibling(target, ext)
        if Files.exists(target, LinkOption.NOFOLLOW_LINKS) then createPrivate(staging)
        else Files.createFile(staging)
        // a JVM stopped mid-conversion (SIGTERM) skips the release below; this still runs
        deleteOnExit(staging)
        staging
      }.adaptError { case e: IOException => RasterError.OutputFailed(target, e) }
    }
    // removal is best-effort: a failed clean-up must not replace the conversion's outcome
    def remove(p: Path): IO[Unit] = IO.blocking(Files.deleteIfExists(p)).void.handleError(_ => ())
    Resource.make(staged)(remove).use { staging =>
      rasterizer.convertSvgToRaster(svg, staging, format, dpi) >>
        IO.blocking(Files.isRegularFile(staging) && Files.size(staging) > 0).flatMap {
          case true =>
            IO.blocking(publish(staging, target)).adaptError { case e: IOException =>
              RasterError.OutputFailed(target, e)
            }
          case false =>
            IO.raiseError(
              RasterError
                .ConversionFailed(rasterizer.name, "exited 0 without writing any output", 0)
            )
        }
    }

  /**
   * An empty file only its owner can read: mode 0600 where the file system has POSIX permissions,
   * else (NTFS) an ACL that grants its owner alone, set before anything is written to it.
   */
  private def createPrivate(path: Path): Unit =
    try
      Files.createFile(
        path,
        PosixFilePermissions.asFileAttribute(PosixFilePermissions.fromString("rw-------"))
      ): Unit
    catch
      case _: UnsupportedOperationException =>
        Files.createFile(path)
        Option(Files.getFileAttributeView(path, classOf[AclFileAttributeView])).foreach { view =>
          val ownerOnly = AclEntry
            .newBuilder()
            .setType(AclEntryType.ALLOW)
            .setPrincipal(view.getOwner)
            .setPermissions(AclEntryPermission.values*)
            .build()
          view.setAcl(java.util.List.of(ownerOnly))
        }

  /**
   * A hidden sibling of `target` for this run: `.xl-raster-<uuid>.<ext>`, short whatever `target`.
   */
  private def sibling(target: Path, ext: String): Path =
    target.resolveSibling(s".xl-raster-${UUID.randomUUID()}.$ext")

  /** Delete `path` when the JVM exits, where the security policy allows. */
  private def deleteOnExit(path: Path): Unit =
    try path.toFile.deleteOnExit()
    catch case _: SecurityException => ()

  /** Delete `path` if it exists; a failure to is not the caller's outcome. */
  private def discard(path: Path): Unit =
    try Files.deleteIfExists(path): Unit
    catch case _: IOException => ()

  /**
   * Publish the verified `staging` file as `target`. An output that already exists is overwritten
   * in place — same file, so its owner, group, ACLs, hard links and symlink are all kept (GH-690) —
   * and a new one is moved into place, atomically where the file system can. A symlink whose
   * referent does not exist yet is written through, creating the referent.
   *
   * The overwrite truncates first, so the old contents are copied to a private sibling beforehand,
   * and a copy that then fails (a full disk, a quota) writes them back; an output that cannot even
   * be opened (read-only) was never truncated, so its failure is reported with the backup removed.
   * If the write-back fails too, the sibling is the only intact copy: it is kept where it is and
   * the error names it. It is never registered for deletion at exit either, since a JVM stopped
   * mid-overwrite leaves the same situation. An output that cannot be read (write-only) cannot be
   * backed up, so it is overwritten without a write-back.
   *
   * The in-place overwrite is the trade-off for keeping the output's identity: it is not atomic (a
   * reader watching the file can see it truncated mid-copy), and it briefly needs about three times
   * the image's size on disk (staging, backup, output). `copy` is the write itself, a parameter so
   * tests can make it fail.
   */
  private[raster] def publish(
    staging: Path,
    target: Path,
    copy: (Path, OutputStream) => Unit = (from, out) => Files.copy(from, out): Unit
  ): Unit =
    def open(to: Path): OutputStream =
      Files.newOutputStream(
        to,
        StandardOpenOption.WRITE,
        StandardOpenOption.CREATE,
        StandardOpenOption.TRUNCATE_EXISTING
      )
    def overwrite(from: Path, to: Path): Unit = Using.resource(open(to))(out => copy(from, out))
    // attempt refused these up front; checked again, since a FIFO made since would hang the backup
    if Files.exists(target) && !Files.isRegularFile(target) then
      throw new IOException("not a regular file")
    if Files.exists(target) && !Files.isReadable(target) then
      // write-only: there is nothing to back up, so it is overwritten as a backend would
      overwrite(staging, target)
    else if Files.exists(target) then
      val backup = sibling(target, "orig")
      createPrivate(backup)
      try
        Using.resource(Files.newOutputStream(backup, StandardOpenOption.WRITE))(out =>
          Files.copy(target, out): Unit
        )
      catch
        case e: IOException =>
          discard(backup)
          throw e
      // an output that cannot be opened (read-only) was never truncated: nothing to write back
      val out =
        try open(target)
        catch
          case e: IOException =>
            discard(backup)
            throw e
      try Using.resource(out)(copy(staging, _))
      catch
        case failed: IOException =>
          try overwrite(backup, target)
          catch
            case restoreFailed: IOException =>
              val error = new IOException(
                s"${failed.getMessage}; its previous contents are kept in $backup",
                failed
              )
              error.addSuppressed(restoreFailed)
              throw error
          discard(backup)
          throw failed
      discard(backup)
    else if Files.isSymbolicLink(target) then
      // a dangling symlink: a failed write removes the referent it started, never the link
      try overwrite(staging, target)
      catch
        case failed: IOException =>
          try Files.deleteIfExists(target.toRealPath()): Unit
          catch case _: IOException => ()
          throw failed
    else
      try
        Files.move(
          staging,
          target,
          StandardCopyOption.REPLACE_EXISTING,
          StandardCopyOption.ATOMIC_MOVE
        ): Unit
      catch
        case _: AtomicMoveNotSupportedException =>
          Files.move(staging, target, StandardCopyOption.REPLACE_EXISTING): Unit

  private def installHintFor(name: String): String = name match
    case "batik" =>
      if NativeImage.inNativeImage then
        "Batik is bundled but needs AWT - unavailable in the native binary by design"
      else "Batik is bundled but requires AWT, which is unavailable in this environment"
    case "cairosvg" => "Install: pip install cairosvg"
    case "rsvg-convert" => "Install: apt install librsvg2-bin (Debian/Ubuntu)"
    case "resvg" =>
      "Install: cargo install resvg, or a prebuilt binary from https://github.com/linebender/resvg/releases"
    case "imagemagick" => "Install: apt install imagemagick"
    case _ => s"Unknown rasterizer: $name"

  /**
   * Scale SVG width/height attributes for requested DPI.
   *
   * SVG uses 96 DPI as the reference (CSS px). To produce larger pixel output at higher DPI, we
   * scale the width/height attributes while keeping viewBox unchanged. This ensures consistent
   * behavior across all rasterizers (some treat DPI as metadata-only, others scale).
   *
   * Example: 300 DPI → scale = 3.125x → 100x50 SVG becomes 312x156 pixels
   *
   * @param svg
   *   Original SVG content
   * @param dpi
   *   Target DPI
   * @return
   *   Scaled SVG content
   */
  def scaleSvgForDpi(svg: String, dpi: Int): String =
    if dpi == BaseDpi then svg
    else
      val scale = dpi.toDouble / BaseDpi

      // Match width="N" and height="N" attributes (integer values)
      val widthPattern = """width="(\d+)"""".r
      val heightPattern = """height="(\d+)"""".r

      // Extract and scale dimensions, replacing only the first occurrence (root SVG element).
      // Uses substring to avoid replacing matching values in child elements (e.g., rects).
      val withScaledWidth = widthPattern.findFirstMatchIn(svg).fold(svg) { m =>
        val scaled = (m.group(1).toInt * scale).toInt
        svg.substring(0, m.start) + s"""width="$scaled"""" + svg.substring(m.end)
      }

      val withScaledHeight =
        heightPattern.findFirstMatchIn(withScaledWidth).fold(withScaledWidth) { m =>
          val scaled = (m.group(1).toInt * scale).toInt
          withScaledWidth.substring(0, m.start) + s"""height="$scaled"""" + withScaledWidth
            .substring(m.end)
        }

      withScaledHeight
