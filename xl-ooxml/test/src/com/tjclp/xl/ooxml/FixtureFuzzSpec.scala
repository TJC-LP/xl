package com.tjclp.xl.ooxml

import java.util.zip.ZipFile
import scala.jdk.CollectionConverters.*
import munit.FunSuite
import com.tjclp.xl.xml.{XmlLimits, XmlPullParser, XmlToken}

/**
 * GH-543: byte-mutation fuzz over REAL fixture XML parts under XmlLimits.default — the parser must
 * return Left or reach EndDocument, never throw, and terminate within a linear step budget. The
 * try/catch lives in this harness only.
 */
class FixtureFuzzSpec extends FunSuite:

  private val mutationsPerPart = 150

  private def xmlParts(fixture: String): List[(String, Array[Byte])] =
    val path = TestFixtures.copyToTemp(fixture)
    val zip = new ZipFile(path.toFile)
    try
      zip
        .entries()
        .asScala
        .filter(e => !e.isDirectory && e.getName.endsWith(".xml"))
        .map { e =>
          val in = zip.getInputStream(e)
          try e.getName -> in.readAllBytes()
          finally in.close()
        }
        .toList
    finally zip.close()

  private def assertTotal(bytes: Array[Byte], label: String): Unit =
    val budget = 2 * bytes.length + 64
    val result =
      try
        val p = XmlPullParser.fromArray(bytes, XmlLimits.default)
        var calls = 0
        var r = p.next()
        calls += 1
        while r.isRight && r != Right(XmlToken.EndDocument) && calls <= budget do
          r = p.next()
          calls += 1
        assert(calls <= budget, s"$label: exceeded step budget ($calls > $budget)")
        r
      catch
        case t: Throwable =>
          fail(s"$label: parser threw ${t.getClass.getName}: ${t.getMessage}")
    result match
      case Left(e) => assert(e.line >= 1 && e.col >= 1 && e.message.nonEmpty, s"$label: $e")
      case Right(XmlToken.EndDocument) => ()
      case other => fail(s"$label: non-terminal $other")

  // Deterministic mutations: seeded per part so failures reproduce.
  private def mutations(bytes: Array[Byte], seed: Long): Iterator[(String, Array[Byte])] =
    val rnd = new java.util.Random(seed)
    Iterator.tabulate(mutationsPerPart) { i =>
      rnd.nextInt(4) match
        case 0 => // flip one byte
          val m = bytes.clone()
          if m.nonEmpty then m(rnd.nextInt(m.length)) = rnd.nextInt(256).toByte
          (s"flip#$i", m)
        case 1 => // truncate
          (s"truncate#$i", bytes.take(rnd.nextInt(bytes.length + 1)))
        case 2 => // insert a byte
          val pos = rnd.nextInt(bytes.length + 1)
          val b = rnd.nextInt(256).toByte
          (s"insert#$i", (bytes.take(pos) :+ b) ++ bytes.drop(pos))
        case _ => // delete a span
          if bytes.isEmpty then (s"delete#$i", bytes)
          else
            val from = rnd.nextInt(bytes.length)
            val len = 1 + rnd.nextInt(math.min(16, bytes.length - from))
            (s"delete#$i", bytes.take(from) ++ bytes.drop(from + len))
    }

  // A representative slice of the corpus: openpyxl + LibreOffice + Excel dialects and the
  // doctype-hostile derivative; every XML part of each, 150 seeded mutations per part.
  private val fuzzFixtures =
    List("small-values.xlsx", "styled-lo.xlsx", "datatable-excel.xlsx", "doctype-hostile.xlsx")

  fuzzFixtures.foreach { fixture =>
    test(s"fuzz: mutated parts of $fixture are total (Left or EndDocument, never throw)") {
      xmlParts(fixture).foreach { case (partName, bytes) =>
        assertTotal(bytes, s"$fixture!$partName (unmutated)")
        mutations(bytes, seed = partName.hashCode.toLong).foreach { case (mLabel, mutated) =>
          assertTotal(mutated, s"$fixture!$partName $mLabel")
        }
      }
    }
  }
