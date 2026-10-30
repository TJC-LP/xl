package com.tjclp.xl.cli.helpers

import java.nio.file.Path
import java.util.zip.ZipFile

import cats.effect.IO

import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.formula.eval.DynamicArrayAuthoring.{Spill, SpillBlocked, SpillBlocker}
import com.tjclp.xl.io.ExcelIO
import com.tjclp.xl.ooxml.{XmlSecurity, FormulaStorage}

/** GH-738: retain only dynamic anchors meeting content targets, never unrelated worksheet cells. */
private[cli] object StreamingSpills:
  /** Plan one content write against the current batch state. Style-only writes never call this. */
  def release(
    spills: Vector[Spill],
    targets: Vector[CellRange]
  ): Either[SpillBlocked, Vector[Spill]] =
    val overwritten = spills.filter(s => targets.exists(_.contains(s.anchor)))
    val blocked = spills.iterator
      .filterNot(overwritten.contains)
      .flatMap { spill =>
        val hits = targets.iterator
          .flatMap(_.intersect(spill.extent))
          .flatMap(_.cellsRowMajor)
          .filter(_ != spill.anchor)
          .take(10)
          .toVector
          .distinct
        Option.when(hits.nonEmpty)(
          SpillBlocked(
            s"=${spill.formula}",
            spill.anchor,
            Some(spill.extent),
            SpillBlocker.Overwritten(hits)
          )
        )
      }
      .nextOption()
    blocked.toLeft(overwritten)

  @SuppressWarnings(Array("org.wartremover.warts.Var")) // SAX cursor state
  def read(source: Path, worksheetPath: String, targets: Vector[CellRange]): Vector[Spill] =
    val zip = new ZipFile(source.toFile)
    try
      // Share the streaming reader's relationship resolution, optional-part fallback and ZIP guards.
      val index = ExcelIO.instance[IO].loadCellMetadataIndexSync(zip).fold(throw _, identity)
      if index.byCm.isEmpty then Vector.empty
      else
        val spills = Vector.newBuilder[Spill]
        val handler = new DefaultHandler:
          private var cell: Option[ARef] = None
          private var dynamic = false
          private var extent: Option[CellRange] = None
          private var inFormula = false
          private val formula = new StringBuilder

          override def startElement(
            uri: String,
            local: String,
            qName: String,
            a: Attributes
          ): Unit =
            val name =
              if local.nonEmpty then local else qName.split(':').lastOption.getOrElse(qName)
            if name == "c" then
              cell = Option(a.getValue("r")).flatMap(ARef.parse(_).toOption)
              dynamic =
                Option(a.getValue("cm")).flatMap(_.trim.toIntOption).exists(index.byCm.contains)
            else if name == "f" then
              formula.clear()
              extent = Option
                .when(dynamic && Option(a.getValue("t")).contains("array"))(
                  Option(a.getValue("ref")).flatMap(CellRange.parse(_).toOption)
                )
                .flatten
                .filter(range => targets.exists(_.intersects(range)))
              inFormula = extent.nonEmpty

          override def characters(ch: Array[Char], start: Int, length: Int): Unit =
            if inFormula then formula.appendAll(ch, start, length)

          override def endElement(uri: String, local: String, qName: String): Unit =
            val name =
              if local.nonEmpty then local else qName.split(':').lastOption.getOrElse(qName)
            if name == "f" then
              for
                anchor <- cell
                range <- extent
                if range.start == anchor && range.cellCount > 1
              do spills += Spill(anchor, FormulaStorage.fromStored(formula.result()), range)
              inFormula = false
              extent = None
        val in =
          XmlSecurity.stripLeadingDoctypeStream(zip.getInputStream(zip.getEntry(worksheetPath)))
        try XmlSecurity.secureSaxParserFactory().newSAXParser().parse(in, handler)
        finally in.close()
        spills.result().sortBy(s => (s.anchor.row.index0, s.anchor.col.index0))
    finally zip.close()
