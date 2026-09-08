package com.tjclp.xl.io.streaming

import java.io.{FilterInputStream, IOException, InputStream}
import java.util.zip.{ZipEntry, ZipFile}

import scala.util.control.NoStackTrace

import com.tjclp.xl.error.XLError
import com.tjclp.xl.ooxml.XlsxReader.ReaderConfig

/**
 * The reader's ZIP-bomb limits applied to ONE entry consumed as a stream (GH-640). `XlsxReader` and
 * `WorkbookMetadataReader` materialise an entry and then check its size and compression ratio; a
 * SAX parser never materialises the part, so the bytes are counted as they are inflated and the
 * read that crosses a limit refuses to go on — `maxUncompressedSize` bytes, or
 * `maxCompressionRatio` times the entry's compressed size, with the same wording as the readers'
 * own checks. A limit of 0 disables its check, as in [[ReaderConfig]].
 */
object ZipEntryGuard:

  /**
   * Raised by the stream at the read that crosses a limit: the parser's `IOException` path,
   * carrying the finished `SecurityError` so the caller has nothing to reconstruct.
   */
  final class LimitExceeded(val error: XLError.SecurityError)
      extends IOException(error.message)
      with NoStackTrace

  /**
   * `entry`'s inflated bytes under `config`'s limits; every read past a limit raises
   * [[LimitExceeded]].
   */
  def open(zipFile: ZipFile, entry: ZipEntry, config: ReaderConfig): InputStream =
    new Bounded(zipFile.getInputStream(entry), entry, config)

  private final class Bounded(in: InputStream, entry: ZipEntry, config: ReaderConfig)
      extends FilterInputStream(in):
    private val sizeCap: Long =
      if config.maxUncompressedSize > 0 then config.maxUncompressedSize else Long.MaxValue
    private val compressed: Long = entry.getCompressedSize
    private val ratioCap: Long =
      if config.maxCompressionRatio > 0 && compressed > 0 then
        config.maxCompressionRatio.toLong * compressed
      else Long.MaxValue

    // The byte counter of the hot read path
    @SuppressWarnings(Array("org.wartremover.warts.Var"))
    private var inflated: Long = 0L

    override def read(): Int =
      val b = in.read()
      if b >= 0 then count(1L)
      b

    override def read(buf: Array[Byte], off: Int, len: Int): Int =
      val n = in.read(buf, off, len)
      if n > 0 then count(n.toLong)
      n

    private def count(n: Long): Unit =
      inflated += n
      if inflated > sizeCap then
        throw new LimitExceeded(
          XLError.SecurityError(
            s"Uncompressed size of '${entry.getName}' exceeds limit (${config.maxUncompressedSize} bytes)"
          )
        )
      if inflated > ratioCap then
        val ratio = inflated.toDouble / compressed.toDouble
        throw new LimitExceeded(
          XLError.SecurityError(
            f"Compression ratio ($ratio%.1f:1) for '${entry.getName}' exceeds limit (${config.maxCompressionRatio}:1) - possible ZIP bomb"
          )
        )
