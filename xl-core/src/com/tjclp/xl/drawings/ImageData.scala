package com.tjclp.xl.drawings

import java.security.MessageDigest
import scala.collection.immutable.ArraySeq

import com.tjclp.xl.error.{XLError, XLResult}
import com.tjclp.xl.styles.units.{Emu, Px}

/**
 * Raster/vector image formats supported by the drawing layer (GH-221).
 *
 * The seven formats cover everything Excel itself embeds via DrawingML `a:blip`. Dimension sniffing
 * ([[ImageData.dimensionsPx]]) is supported for the four raster formats with trivial headers
 * (Png/Gif/Bmp/Jpeg); Tiff/Emf/Wmf round-trip but cannot be auto-sized.
 */
enum ImageFormat derives CanEqual:
  case Png, Jpeg, Gif, Bmp, Tiff, Emf, Wmf

  /** Canonical file extension used for `xl/media/imageN.<ext>` part names. */
  def extension: String = this match
    case Png => "png"
    case Jpeg => "jpeg"
    case Gif => "gif"
    case Bmp => "bmp"
    case Tiff => "tiff"
    case Emf => "emf"
    case Wmf => "wmf"

  /** MIME content type registered as a `Default` in `[Content_Types].xml`. */
  def contentType: String = this match
    case Png => "image/png"
    case Jpeg => "image/jpeg"
    case Gif => "image/gif"
    case Bmp => "image/bmp"
    case Tiff => "image/tiff"
    case Emf => "image/x-emf"
    case Wmf => "image/x-wmf"

object ImageFormat:
  /**
   * Detect the format from magic bytes. Total: truncated or unrecognized input yields None.
   *
   * Magic numbers: PNG `89504E47`, JPEG `FFD8FF`, GIF `GIF87a`/`GIF89a`, BMP `BM`, TIFF
   * `II*\0`/`MM\0*`, EMF record type 1 (`01000000`) with `" EMF"` at offset 40, WMF placeable
   * (`D7CDC69A`) or standard (`0100 0900`).
   */
  def detect(bytes: ArraySeq[Byte]): Option[ImageFormat] =
    def at(i: Int): Int = bytes(i) & 0xff
    def has(n: Int): Boolean = bytes.length >= n
    if has(4) && at(0) == 0x89 && at(1) == 0x50 && at(2) == 0x4e && at(3) == 0x47 then Some(Png)
    else if has(3) && at(0) == 0xff && at(1) == 0xd8 && at(2) == 0xff then Some(Jpeg)
    else if has(6) && at(0) == 'G' && at(1) == 'I' && at(2) == 'F' && at(3) == '8' &&
      (at(4) == '7' || at(4) == '9') && at(5) == 'a'
    then Some(Gif)
    else if has(4) &&
      ((at(0) == 0x49 && at(1) == 0x49 && at(2) == 0x2a && at(3) == 0x00) ||
        (at(0) == 0x4d && at(1) == 0x4d && at(2) == 0x00 && at(3) == 0x2a))
    then Some(Tiff)
    else if has(44) && at(0) == 0x01 && at(1) == 0x00 && at(2) == 0x00 && at(3) == 0x00 &&
      at(40) == 0x20 && at(41) == 0x45 && at(42) == 0x4d && at(43) == 0x46
    then Some(Emf)
    else if has(4) &&
      ((at(0) == 0xd7 && at(1) == 0xcd && at(2) == 0xc6 && at(3) == 0x9a) ||
        (at(0) == 0x01 && at(1) == 0x00 && at(2) == 0x09 && at(3) == 0x00))
    then Some(Wmf)
    else if has(2) && at(0) == 'B' && at(1) == 'M' then Some(Bmp)
    else None

  /** Case-insensitive extension lookup; `jpg`→Jpeg, `tif`→Tiff aliases included. Total. */
  def fromExtension(ext: String): Option[ImageFormat] =
    ext.toLowerCase match
      case "png" => Some(Png)
      case "jpeg" | "jpg" => Some(Jpeg)
      case "gif" => Some(Gif)
      case "bmp" => Some(Bmp)
      case "tiff" | "tif" => Some(Tiff)
      case "emf" => Some(Emf)
      case "wmf" => Some(Wmf)
      case _ => None

/**
 * Immutable image payload: raw bytes plus declared format.
 *
 * `ArraySeq[Byte]` (not `Array`/`IArray`) so structural equality holds — the surgical writer's
 * snapshot-equality dirty test depends on it. Validation lives in [[ImageData.detect]] (the
 * `ARef.parse` house pattern); constructing a deliberately mismatched (bytes, format) pair is the
 * caller's responsibility and produces a file Excel may refuse to render.
 */
final case class ImageData(bytes: ArraySeq[Byte], format: ImageFormat) derives CanEqual:

  /** SHA-256 of the bytes as lowercase hex: the media dedup key and a fast equality aid. */
  lazy val sha256: String =
    val digest = MessageDigest.getInstance("SHA-256").digest(bytes.toArray)
    digest.map("%02x".format(_)).mkString

  /**
   * Pixel dimensions sniffed from the header: PNG IHDR (bytes 16-23, big-endian), GIF logical
   * screen (6-9, little-endian), BMP info header (18/22, little-endian, absolute value), JPEG SOFn
   * segment scan. Tiff/Emf/Wmf (and malformed headers) yield None. Total.
   */
  def dimensionsPx: Option[(Int, Int)] =
    def at(i: Int): Int = bytes(i) & 0xff
    def has(n: Int): Boolean = bytes.length >= n
    def be16(i: Int): Int = (at(i) << 8) | at(i + 1)
    def be32(i: Int): Int = (at(i) << 24) | (at(i + 1) << 16) | (at(i + 2) << 8) | at(i + 3)
    def le16(i: Int): Int = at(i) | (at(i + 1) << 8)
    def le32(i: Int): Int = at(i) | (at(i + 1) << 8) | (at(i + 2) << 16) | (at(i + 3) << 24)
    format match
      case ImageFormat.Png =>
        Option.when(has(24))((be32(16), be32(20))).filter((w, h) => w > 0 && h > 0)
      case ImageFormat.Gif =>
        Option.when(has(10))((le16(6), le16(8))).filter((w, h) => w > 0 && h > 0)
      case ImageFormat.Bmp =>
        // Height may be negative (top-down DIB); dimensions are magnitudes
        Option
          .when(has(26))((math.abs(le32(18)), math.abs(le32(22))))
          .filter((w, h) => w > 0 && h > 0)
      case ImageFormat.Jpeg => jpegDimensions
      case ImageFormat.Tiff | ImageFormat.Emf | ImageFormat.Wmf => None

  /**
   * Walk JPEG marker segments to the first SOFn frame header (C0-CF excluding C4/C8/CC) and read
   * height/width. Bails to None on any structural inconsistency (truncation, bad lengths).
   */
  @SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
  private def jpegDimensions: Option[(Int, Int)] =
    def at(i: Int): Int = bytes(i) & 0xff
    def be16(i: Int): Int = (at(i) << 8) | at(i + 1)
    val len = bytes.length
    var i = 2
    var result: Option[(Int, Int)] = None
    var done = len < 4 || at(0) != 0xff || at(1) != 0xd8
    while !done && i + 3 < len do
      if at(i) != 0xff then done = true
      else
        val marker = at(i + 1)
        if marker == 0xff then i += 1 // fill byte
        else if marker >= 0xc0 && marker <= 0xcf && marker != 0xc4 && marker != 0xc8 &&
          marker != 0xcc
        then
          if i + 8 < len then
            val h = be16(i + 5)
            val w = be16(i + 7)
            if w > 0 && h > 0 then result = Some((w, h))
          done = true
        else if marker == 0xd8 || (marker >= 0xd0 && marker <= 0xd7) then i += 2 // standalone
        else if marker == 0xd9 || marker == 0xda then done = true // EOI / SOS: no SOF found
        else
          val segLen = be16(i + 2)
          if segLen < 2 then done = true else i += 2 + segLen
    result

  /** Natural extent at 96 DPI (9525 EMU per pixel); None when dimensions cannot be sniffed. */
  def naturalExtent: Option[Extent] =
    dimensionsPx.map((w, h) => Extent.fromPx(w, h))

object ImageData:
  /**
   * Classify raw bytes by magic number. Left when the format is unrecognized — the total
   * alternative to constructing `ImageData` directly with a known format.
   */
  def detect(bytes: ArraySeq[Byte]): XLResult[ImageData] =
    ImageFormat.detect(bytes) match
      case Some(format) => Right(ImageData(bytes, format))
      case None =>
        Left(
          XLError.ParseError(
            "image bytes",
            "unrecognized image format (magic bytes match none of png/jpeg/gif/bmp/tiff/emf/wmf)"
          )
        )
