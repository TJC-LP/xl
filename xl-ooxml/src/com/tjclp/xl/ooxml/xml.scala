package com.tjclp.xl.ooxml

import scala.xml.*

/**
 * XML serialization and deserialization for OOXML parts.
 *
 * All serialization is deterministic with canonical attribute/element ordering for stable diffs and
 * golden tests.
 */

/** Trait for types that can be serialized to XML */
trait XmlWritable:
  /** Convert to XML element with deterministic formatting */
  def toXml: Elem

/** Trait for types that can be deserialized from XML */
trait XmlReadable[A]:
  /** Parse from XML element, returning error message on failure */
  def fromXml(elem: Elem): Either[String, A]

/** XML utilities for OOXML */
object XmlUtil:
  /** OOXML namespaces */
  val nsSpreadsheetML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  val nsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  val nsPackageRels = "http://schemas.openxmlformats.org/package/2006/relationships"
  val nsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types"

  /** Relationship type URIs */
  val relTypeOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
  val relTypeWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  val relTypeStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
  val relTypeSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"

  /** Content type URIs */
  val ctWorkbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
  val ctWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
  val ctStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
  val ctSharedStrings =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
  val ctRelationships = "application/vnd.openxmlformats-package.relationships+xml"

  /** Sort attributes by name for deterministic output */
  def sortAttributes(attrs: MetaData): MetaData =
    val sorted = attrs.asAttrMap.toSeq.sortBy(_._1)
    sorted.foldLeft(Null: MetaData) { case (acc, (key, value)) =>
      new UnprefixedAttribute(key, value, acc)
    }

  /** Create element with sorted attributes */
  def elem(label: String, attrs: (String, String)*)(children: Node*): Elem =
    val sortedAttrs = attrs.sortBy(_._1).foldLeft(Null: MetaData) { case (acc, (key, value)) =>
      new UnprefixedAttribute(key, value, acc)
    }
    Elem(null, label, sortedAttrs, TopScope, minimizeEmpty = true, children*)

  /** Create element with namespace */
  def elemNS(prefix: String, label: String, ns: String, attrs: (String, String)*)(
    children: Node*
  ): Elem =
    val sortedAttrs = attrs.sortBy(_._1).foldLeft(Null: MetaData) { case (acc, (key, value)) =>
      new UnprefixedAttribute(key, value, acc)
    }
    Elem(
      prefix,
      label,
      sortedAttrs,
      NamespaceBinding(prefix, ns, TopScope),
      minimizeEmpty = true,
      children*
    )

  /** Pretty print XML with proper indentation */
  def prettyPrint(node: Node): String =
    val printer = new PrettyPrinter(80, 2)
    s"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${printer.format(node)}"""

  /** Get required attribute value */
  def getAttr(elem: Elem, name: String): Either[String, String] =
    elem.attribute(name).map(_.text).toRight(s"Missing required attribute: $name")

  /** Get optional attribute value */
  def getAttrOpt(elem: Elem, name: String): Option[String] =
    elem.attribute(name).map(_.text)

  /** Get required child element */
  def getChild(elem: Elem, label: String): Either[String, Elem] =
    (elem \ label).headOption match
      case Some(e: Elem) => Right(e)
      case _ => Left(s"Missing required child element: $label")

  /** Get all child elements with given label */
  def getChildren(elem: Elem, label: String): Seq[Elem] =
    (elem \ label).collect { case e: Elem => e }
