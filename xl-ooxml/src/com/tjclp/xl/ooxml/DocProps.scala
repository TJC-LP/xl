package com.tjclp.xl.ooxml

import java.time.{LocalDate, LocalDateTime, OffsetDateTime, ZoneOffset}
import java.time.format.DateTimeFormatter
import java.time.temporal.ChronoUnit

import scala.util.Try
import scala.xml.*

import com.tjclp.xl.workbooks.WorkbookMetadata

/**
 * Document properties parts (GH-242): docProps/core.xml (OPC core properties) and docProps/app.xml
 * (extended properties).
 *
 * Emission is deterministic and model-driven: ONLY fields present on [[WorkbookMetadata]] are
 * emitted — no GUIDs, no wall-clock timestamps. `created`/`modified` serialize in W3CDTF (UTC,
 * second precision) iff set. A part with no present fields is omitted entirely.
 *
 * Parsing is lenient (foreign files carry many unmodeled fields — title, keywords, revision, ... —
 * which are ignored): element absent → None, element present → Some(text). W3CDTF values with an
 * explicit offset are normalized to UTC; date-only values parse to midnight.
 */
object DocProps:

  /** Zip entry paths for the two modeled docProps parts. */
  val corePath = "docProps/core.xml"
  val appPath = "docProps/app.xml"

  // Namespaces for core.xml (OPC core properties, ECMA-376 Part 2 §11)
  private val nsCoreProps =
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  private val nsDc = "http://purl.org/dc/elements/1.1/"
  private val nsDcTerms = "http://purl.org/dc/terms/"
  private val nsDcmiType = "http://purl.org/dc/dcmitype/"
  private val nsXsi = "http://www.w3.org/2001/XMLSchema-instance"

  // Namespaces for app.xml (extended properties, ECMA-376 Part 1 §22.2)
  private val nsExtendedProps =
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  private val nsVt =
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

  /** Modeled docProps fields parsed from a package (all None when parts are absent). */
  final case class Data(
    creator: Option[String] = None,
    created: Option[LocalDateTime] = None,
    modified: Option[LocalDateTime] = None,
    lastModifiedBy: Option[String] = None,
    application: Option[String] = None,
    appVersion: Option[String] = None
  ) derives CanEqual

  object Data:
    val empty: Data = Data()

  private val w3cDtfFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss'Z'")

  /** W3CDTF rendering of a model datetime: treated as UTC, truncated to seconds. */
  def formatW3cDtf(dt: LocalDateTime): String =
    dt.truncatedTo(ChronoUnit.SECONDS).format(w3cDtfFormatter)

  /**
   * Lenient W3CDTF parse: offset forms are normalized to UTC, zone-less forms are taken as UTC,
   * date-only forms parse to midnight. Unparseable text → None (never an error: docProps in the
   * wild are too messy to fail a read over).
   */
  def parseW3cDtf(s: String): Option[LocalDateTime] =
    val trimmed = s.trim
    if trimmed.isEmpty then None
    else
      Try(
        OffsetDateTime.parse(trimmed).withOffsetSameInstant(ZoneOffset.UTC).toLocalDateTime
      ).toOption
        .orElse(Try(LocalDateTime.parse(trimmed)).toOption)
        .orElse(Try(LocalDate.parse(trimmed).atStartOfDay).toOption)

  /** Scope chain for `<cp:coreProperties>` (constructed once; deterministic rendering). */
  private val coreScope: NamespaceBinding =
    NamespaceBinding(
      "cp",
      nsCoreProps,
      NamespaceBinding(
        "dc",
        nsDc,
        NamespaceBinding(
          "dcterms",
          nsDcTerms,
          NamespaceBinding("dcmitype", nsDcmiType, NamespaceBinding("xsi", nsXsi, TopScope))
        )
      )
    )

  private def prefixed(prefix: String, label: String, text: String): Elem =
    Elem(prefix, label, Null, TopScope, minimizeEmpty = true, Text(text))

  private def w3cDtfElem(label: String, dt: LocalDateTime): Elem =
    Elem(
      "dcterms",
      label,
      new PrefixedAttribute("xsi", "type", "dcterms:W3CDTF", Null),
      TopScope,
      minimizeEmpty = true,
      Text(formatW3cDtf(dt))
    )

  /**
   * Build docProps/core.xml from the model. Returns None when no core field is set (the part is
   * omitted, along with its content-type override and package relationship).
   */
  def buildCoreXml(meta: WorkbookMetadata): Option[Elem] =
    val children = Vector(
      meta.creator.map(prefixed("dc", "creator", _)),
      meta.lastModifiedBy.map(prefixed("cp", "lastModifiedBy", _)),
      meta.created.map(w3cDtfElem("created", _)),
      meta.modified.map(w3cDtfElem("modified", _))
    ).flatten
    Option.when(children.nonEmpty)(
      Elem("cp", "coreProperties", Null, coreScope, minimizeEmpty = true, children*)
    )

  /**
   * Build docProps/app.xml from the model. Returns None when neither application nor appVersion is
   * set.
   */
  def buildAppXml(meta: WorkbookMetadata): Option[Elem] =
    val children = Vector(
      meta.application.map(a => Elem(null, "Application", Null, TopScope, true, Text(a))),
      meta.appVersion.map(v => Elem(null, "AppVersion", Null, TopScope, true, Text(v)))
    ).flatten
    val scope = NamespaceBinding(null, nsExtendedProps, NamespaceBinding("vt", nsVt, TopScope))
    Option.when(children.nonEmpty)(
      Elem(null, "Properties", Null, scope, minimizeEmpty = true, children*)
    )

  /** Element text by local name; present-but-empty elements yield Some(""). */
  private def textOf(elem: Elem, label: String): Option[String] =
    (elem \ label).headOption.map(_.text)

  /** Parse the modeled fields of docProps/core.xml (unmodeled fields are ignored). */
  def parseCoreXml(elem: Elem): Data =
    Data(
      creator = textOf(elem, "creator"),
      created = textOf(elem, "created").flatMap(parseW3cDtf),
      modified = textOf(elem, "modified").flatMap(parseW3cDtf),
      lastModifiedBy = textOf(elem, "lastModifiedBy")
    )

  /** Parse the modeled fields of docProps/app.xml (unmodeled fields are ignored). */
  def parseAppXml(elem: Elem): Data =
    Data(
      application = textOf(elem, "Application"),
      appVersion = textOf(elem, "AppVersion")
    )

  /** Merge core- and app-derived fields (disjoint by construction). */
  def merge(core: Data, app: Data): Data =
    core.copy(application = app.application, appVersion = app.appVersion)
