package com.tjclp.xl.drawings

import scala.collection.immutable.ArraySeq

import munit.FunSuite

import com.tjclp.xl.error.XLError
import com.tjclp.xl.styles.units.Emu

/** Unit tests for ImageFormat magic detection and ImageData dimension sniffing (GH-221). */
class ImageDataSpec extends FunSuite:

  private def bytes(xs: Int*): ArraySeq[Byte] =
    ArraySeq.unsafeWrapArray(xs.map(_.toByte).toArray)

  test("detect recognizes all four raster templates") {
    assertEquals(ImageFormat.detect(TestImages.png2x3), Some(ImageFormat.Png))
    assertEquals(ImageFormat.detect(TestImages.gif2x3), Some(ImageFormat.Gif))
    assertEquals(ImageFormat.detect(TestImages.jpeg2x3), Some(ImageFormat.Jpeg))
    assertEquals(ImageFormat.detect(TestImages.bmp2x3), Some(ImageFormat.Bmp))
  }

  test("detect recognizes tiff/emf/wmf magic") {
    assertEquals(ImageFormat.detect(bytes(0x49, 0x49, 0x2a, 0x00)), Some(ImageFormat.Tiff))
    assertEquals(ImageFormat.detect(bytes(0x4d, 0x4d, 0x00, 0x2a)), Some(ImageFormat.Tiff))
    assertEquals(ImageFormat.detect(TestImages.wmfHeader), Some(ImageFormat.Wmf))
    assertEquals(ImageFormat.detect(bytes(0x01, 0x00, 0x09, 0x00)), Some(ImageFormat.Wmf))
    val emf = ArraySeq.unsafeWrapArray(
      (Array(0x01, 0x00, 0x00, 0x00) ++ Array.fill(36)(0x00) ++ Array(0x20, 0x45, 0x4d, 0x46))
        .map(_.toByte)
    )
    assertEquals(ImageFormat.detect(emf), Some(ImageFormat.Emf))
  }

  test("detect is total on truncated and garbage input") {
    assertEquals(ImageFormat.detect(ArraySeq.empty), None)
    assertEquals(ImageFormat.detect(bytes(0x89)), None) // truncated PNG magic
    assertEquals(ImageFormat.detect(bytes(0x47, 0x49, 0x46)), None) // truncated GIF magic
    assertEquals(ImageFormat.detect(bytes(0x00, 0x01, 0x02, 0x03)), None)
    // EMF prefix without the " EMF" signature at offset 40 is not EMF (nor WMF)
    assertEquals(ImageFormat.detect(bytes(0x01, 0x00, 0x00, 0x00)), None)
  }

  test("fromExtension is case-insensitive with jpg/tif aliases") {
    assertEquals(ImageFormat.fromExtension("PNG"), Some(ImageFormat.Png))
    assertEquals(ImageFormat.fromExtension("jpg"), Some(ImageFormat.Jpeg))
    assertEquals(ImageFormat.fromExtension("JPEG"), Some(ImageFormat.Jpeg))
    assertEquals(ImageFormat.fromExtension("tif"), Some(ImageFormat.Tiff))
    assertEquals(ImageFormat.fromExtension("svg"), None)
  }

  test("extension and contentType are aligned per format") {
    assertEquals(ImageFormat.Png.extension, "png")
    assertEquals(ImageFormat.Png.contentType, "image/png")
    assertEquals(ImageFormat.Emf.contentType, "image/x-emf")
    assertEquals(ImageFormat.Wmf.contentType, "image/x-wmf")
    ImageFormat.values.foreach { f =>
      assertEquals(ImageFormat.fromExtension(f.extension), Some(f))
    }
  }

  test("dimensionsPx sniffs png/gif/jpeg/bmp headers") {
    assertEquals(ImageData(TestImages.png2x3, ImageFormat.Png).dimensionsPx, Some((2, 3)))
    assertEquals(ImageData(TestImages.gif2x3, ImageFormat.Gif).dimensionsPx, Some((2, 3)))
    assertEquals(ImageData(TestImages.jpeg2x3, ImageFormat.Jpeg).dimensionsPx, Some((2, 3)))
    assertEquals(ImageData(TestImages.bmp2x3, ImageFormat.Bmp).dimensionsPx, Some((2, 3)))
  }

  test("dimensionsPx is None for unsniffable formats and truncated bytes") {
    assertEquals(ImageData(TestImages.wmfHeader, ImageFormat.Wmf).dimensionsPx, None)
    assertEquals(ImageData(bytes(0x89, 0x50, 0x4e, 0x47), ImageFormat.Png).dimensionsPx, None)
    assertEquals(ImageData(bytes(0xff, 0xd8, 0xff), ImageFormat.Jpeg).dimensionsPx, None)
  }

  test("golden: fixture png is 3x3 px and 28575 EMU at 96 DPI") {
    val img = ImageData(TestImages.fixturePng3x3, ImageFormat.Png)
    assertEquals(img.dimensionsPx, Some((3, 3)))
    assertEquals(img.naturalExtent, Some(Extent(Emu(28575L), Emu(28575L))))
  }

  test("ImageData.detect returns Left on garbage and Right with sniffed format on real bytes") {
    assertEquals(
      ImageData.detect(TestImages.png2x3),
      Right(ImageData(TestImages.png2x3, ImageFormat.Png))
    )
    ImageData.detect(bytes(0x00, 0x01, 0x02)) match
      case Left(XLError.ParseError(_, _)) => () // expected
      case other => fail(s"expected ParseError, got $other")
  }

  test("sha256 is stable hex and equal bytes compare equal structurally") {
    val a = ImageData(TestImages.png2x3, ImageFormat.Png)
    val b = ImageData(
      ArraySeq.unsafeWrapArray(TestImages.png2x3.toArray.clone()),
      ImageFormat.Png
    )
    assertEquals(a, b) // ArraySeq structural equality (the dirty-test invariant)
    assertEquals(a.sha256, b.sha256)
    assertEquals(a.sha256.length, 64)
    assert(a.sha256.forall(c => c.isDigit || (c >= 'a' && c <= 'f')))
  }
