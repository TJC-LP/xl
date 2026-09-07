package com.tjclp.xl.io.streaming

import java.io.ByteArrayOutputStream

import munit.FunSuite

import com.tjclp.xl.cells.CellValue
import com.tjclp.xl.ooxml.StaxSaxWriter

/**
 * GH-556: the in-place `--stream` transform (the streamed `putf` path) is an `<f>` emission
 * boundary too — it must apply Excel's `_xlfn.` storage prefix exactly like the in-memory writers,
 * and never double-prefix a pass-through cell whose text already carries it.
 */
class StreamingFutureFunctionPrefixSpec extends FunSuite:

  private def cellXml(value: CellValue): String =
    val out = new ByteArrayOutputStream()
    val writer = StaxSaxWriter.create(out)
    writer.startElement("c")
    StreamingTransform.writeCellContent(writer, value)
    writer.endElement()
    writer.flush()
    out.toString("UTF-8")

  test("GH-556: a streamed putf of a bare post-2007 function gets the prefix") {
    val xml = cellXml(CellValue.Formula("=MAXIFS(B1:B3,A1:A3,C1)", None))
    assert(xml.contains("<f>_xlfn.MAXIFS(B1:B3,A1:A3,C1)</f>"), xml)
  }

  test("GH-556: FILTER takes _xlfn._xlws.; Excel 2007 functions stay bare") {
    val filter = cellXml(CellValue.Formula("FILTER(A1:A3,B1:B3)", None))
    assert(filter.contains("<f>_xlfn._xlws.FILTER(A1:A3,B1:B3)</f>"), filter)
    val sum = cellXml(CellValue.Formula("SUM(A1:A3)", None))
    assert(sum.contains("<f>SUM(A1:A3)</f>"), sum)
  }

  test("GH-556: a pass-through cell already carrying the prefix is not double-prefixed") {
    val xml = cellXml(CellValue.Formula("_xlfn.XLOOKUP(C1,A1:A3,B1:B3)", None))
    assert(xml.contains("<f>_xlfn.XLOOKUP(C1,A1:A3,B1:B3)</f>"), xml)
    assert(!xml.contains("_xlfn._xlfn."), xml)
  }
