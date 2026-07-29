package com.tjclp.xl.ooxml

import java.io.InputStream
import java.nio.file.{Files, Path, StandardCopyOption}

/**
 * Loader for the committed real-file fixture corpus (GH-240).
 *
 * Fixtures live in xl-ooxml/test/resources/fixtures and are produced by foreign writers (openpyxl,
 * LibreOffice) via scripts/generate-fixtures.py. See fixtures/PROVENANCE.md.
 */
object TestFixtures:

  /** Fixtures produced by openpyxl 3.1.x (inline-string dialect). */
  val openpyxl: List[String] = List(
    "small-values.xlsx",
    "styled.xlsx",
    "formulas.xlsx",
    "autofilter.xlsx",
    "chart-bar.xlsx",
    "chart-stacked.xlsx",
    "chart-scatter.xlsx",
    "image.xlsx",
    "comments-hyperlinks.xlsx",
    "condformat.xlsx"
  )

  /** Fixtures converted through LibreOffice headless (SST dialect, cached formulas). */
  val libreOffice: List[String] = List(
    "small-values-lo.xlsx",
    "styled-lo.xlsx",
    "formulas-lo.xlsx",
    "condformat-lo.xlsx"
  )

  /**
   * Fixtures authored by Microsoft Excel itself (GH-419): the first genuine-Excel files in the
   * corpus. datatable-excel.xlsx carries all three native data-table shapes (2-D, 1-D row, 1-D
   * column) in Excel's own record dialect; datatable-excel-edge.xlsx pins the edge geometry (empty
   * corner, bare single-cell `ref="Q2"`, two-result-column 1-D table). See PROVENANCE.md.
   */
  val excel: List[String] = List(
    "datatable-excel.xlsx",
    "datatable-excel-edge.xlsx"
  )

  /**
   * Fixtures derived from the committed set by deterministic zip surgery (see PROVENANCE.md).
   * image-shape.xlsx = image.xlsx with an `<sp>` shape anchor appended to the same wsDr (GH-221
   * mixed typed-picture + preserved-fragment coverage). doctype-hostile.xlsx = small-values-lo.xlsx
   * made "hostile but Excel-valid" (GH-350): DOCTYPE prologs on workbook/sheet/sst/styles (internal
   * subsets with comment traps), a UTF-8 BOM on styles.xml, 3000 extra cellXfs, and an
   * externalLinks part. formula-records.xlsx = deterministic zip assembly (GH-430): 2-D data table
   * with a cached #NUM! `ca="1"` interior, 1-D row and column data tables, a multi-cell CSE array
   * group, and a single-cell array record.
   */
  val derived: List[String] = List(
    "image-shape.xlsx",
    "doctype-hostile.xlsx",
    "formula-records.xlsx"
  )

  /**
   * Intentionally broken fixtures (GH-349): NOT part of `all` — corpus laws (parse, round-trip)
   * must not run against them. malformed-workbook.xlsx = small-values.xlsx with the closing
   * `</workbook>` tag mangled, so Xerces raises a well-formedness error whose text comes from the
   * `com.sun.org.apache.xerces.internal.impl.msg.XMLMessages` bundle. Guards that parse failures
   * surface the real Xerces diagnostic (also exercised against the native binary in release.yml).
   */
  val malformed: List[String] = List(
    "malformed-workbook.xlsx"
  )

  /** Every committed well-formed fixture (excludes `malformed`). */
  val all: List[String] = openpyxl ++ libreOffice ++ excel ++ derived

  /**
   * Copy a fixture from the test classpath to a temp file (readers need a real Path). The caller
   * owns deletion; tests lean on deleteOnExit for simplicity.
   */
  def copyToTemp(name: String): Path =
    val resource = s"/fixtures/$name"
    val stream: InputStream = Option(getClass.getResourceAsStream(resource)) match
      case Some(s) => s
      case None => throw new IllegalStateException(s"Fixture not on classpath: $resource")
    try
      val target = Files.createTempFile(s"xl-fixture-${name.stripSuffix(".xlsx")}-", ".xlsx")
      target.toFile.deleteOnExit()
      Files.copy(stream, target, StandardCopyOption.REPLACE_EXISTING)
      target
    finally stream.close()
