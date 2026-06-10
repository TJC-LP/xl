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
    "image.xlsx",
    "comments-hyperlinks.xlsx"
  )

  /** Fixtures converted through LibreOffice headless (SST dialect, cached formulas). */
  val libreOffice: List[String] = List(
    "small-values-lo.xlsx",
    "styled-lo.xlsx",
    "formulas-lo.xlsx"
  )

  /** Every committed fixture. */
  val all: List[String] = openpyxl ++ libreOffice

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
