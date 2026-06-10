package com.tjclp.xl.ooxml

import java.nio.file.Files

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.sheets.Sheet
import com.tjclp.xl.workbooks.Workbook
import com.tjclp.xl.addressing.ARef
import munit.FunSuite

/**
 * GH-261: an untouched workbook read from a file and written to a new path must take the
 * verbatim-copy fast path and produce byte-identical output. The read-side digest previously
 * stopped at the ZIP central directory (ZipInputStream never consumes it) while copyVerbatim
 * hashed the whole file, so the fingerprints could never match and every clean copy failed.
 */
class VerbatimCopySpec extends FunSuite:

  private def freshFile(): java.nio.file.Path =
    val dir = Files.createTempDirectory("xl-verbatim")
    val src = dir.resolve("src.xlsx")
    val wb = Workbook(
      Sheet(SheetName.unsafe("Data"))
        .put(ARef.from0(0, 0), CellValue.Text("hello"))
        .put(ARef.from0(1, 0), CellValue.Number(BigDecimal(42)))
    )
    XlsxWriter.write(wb, src).fold(e => fail(s"setup write failed: $e"), identity)
    src

  test("GH-261: clean read -> write to new path succeeds and is byte-identical"):
    val src = freshFile()
    val dest = src.getParent.resolve("copy.xlsx")
    val read = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    assert(read.sourceContext.exists(_.isClean), "freshly read workbook must be clean")
    XlsxWriter.write(read, dest) match
      case Left(err) => fail(s"verbatim copy failed: $err")
      case Right(_) =>
        assert(
          java.util.Arrays.equals(Files.readAllBytes(src), Files.readAllBytes(dest)),
          "verbatim copy must be byte-identical"
        )

  test("GH-261: fingerprint covers the whole file (mutating trailing bytes is detected)"):
    val src = freshFile()
    val dest = src.getParent.resolve("copy2.xlsx")
    val read = XlsxReader.read(src).fold(e => fail(s"read failed: $e"), identity)
    // Append a byte after reading: size check catches it — then mutate the LAST byte in place
    // (same size, inside the central directory) which only a whole-file digest detects.
    val bytes = Files.readAllBytes(src)
    bytes(bytes.length - 1) = (bytes(bytes.length - 1) ^ 0xff).toByte
    Files.write(src, bytes)
    XlsxWriter.write(read, dest) match
      case Left(_) => () // refusing is acceptable
      case Right(_) =>
        fail("write must not verbatim-copy a file whose trailing bytes changed since read")
