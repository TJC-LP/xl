package com.tjclp.xl.ooxml

import scala.xml.{
  Atom,
  Elem,
  MetaData,
  NamespaceBinding,
  Null,
  PrefixedAttribute,
  Text,
  TopScope,
  UnprefixedAttribute,
  PCData
}

/**
 * Shared helpers for emitting scala.xml trees through the SaxWriter algebra.
 *
 * Keeps namespace/attribute handling consistent across interpreters (StAX, Scala XML, fs2, etc.).
 */
trait SaxSupport:

  /** Deterministic attribute extraction from MetaData */
  protected def metaDataAttributes(md: MetaData): Seq[(String, String)] =
    def loop(m: MetaData, acc: List[(String, String)]): List[(String, String)] = m match
      case Null => acc
      case PrefixedAttribute(pre, key, value, next) =>
        loop(next, (s"$pre:$key", value.text) :: acc)
      case UnprefixedAttribute(key, value, next) =>
        loop(next, (key, value.text) :: acc)

    loop(md, Nil).reverse.distinctBy(_._1).sortBy(_._1)

  /** Deterministic namespace declaration extraction from NamespaceBinding chain */
  protected def namespaceAttributes(scope: NamespaceBinding): Seq[(String, String)] =
    def loop(ns: NamespaceBinding, acc: List[(String, String)]): List[(String, String)] =
      if ns == null || ns == TopScope then acc
      else
        val prefix = Option(ns.prefix).getOrElse("")
        val name = if prefix.isEmpty then "xmlns" else s"xmlns:$prefix"
        loop(ns.parent, (name, ns.uri) :: acc)

    loop(scope, Nil).distinctBy(_._1).sortBy(_._1)

  /**
   * Combine namespace and metadata attributes with deduplication.
   *
   * When reading XML files, namespace declarations may appear in both the scope (NamespaceBinding)
   * and root attributes (MetaData). This method merges them and deduplicates by attribute name,
   * keeping the first occurrence (from namespace declarations).
   */
  protected def combinedAttributes(
    scope: NamespaceBinding,
    attrs: MetaData
  ): Seq[(String, String)] =
    (namespaceAttributes(scope) ++ metaDataAttributes(attrs)).distinctBy(_._1)

  /** Stream an existing scala.xml Elem through a SaxWriter */
  protected def emitElem(writer: SaxWriter, elem: Elem): Unit =
    val qName =
      Option(elem.prefix).filter(_.nonEmpty).map(p => s"$p:${elem.label}").getOrElse(elem.label)
    // getURI returns null for undefined prefixes; Option() wrapper handles this safely
    val nsUri = Option(elem.prefix)
      .flatMap(p => Option(elem.scope).flatMap(sc => Option(sc.getURI(p))))
      .filter(_.nonEmpty)

    nsUri match
      case Some(uri) => writer.startElement(qName, uri)
      case None => writer.startElement(qName)

    SaxWriter.withAttributes(writer, metaDataAttributes(elem.attributes)*) {
      elem.child.foreach {
        case e: Elem => emitElem(writer, e)
        case t: Text => writer.writeCharacters(t.data)
        case pc: PCData => writer.writeCharacters(pc.data)
        // Atom[String] for interpolated values in XML literals (e.g., <elem>{stringVal}</elem>)
        // Must come after Text/PCData since those extend Atom[String]
        case a: Atom[?] => writer.writeCharacters(a.data.toString)
        case _ => ()
      }
    }

    writer.endElement()

object SaxSupport extends SaxSupport:
  extension (writer: SaxWriter)
    def writeElem(elem: Elem): Unit = emitElem(writer, elem)
    def namespaceAttributes(scope: NamespaceBinding): Seq[(String, String)] =
      SaxSupport.namespaceAttributes(scope)
    def metaDataAttributes(md: MetaData): Seq[(String, String)] =
      SaxSupport.metaDataAttributes(md)
    def combinedAttributes(scope: NamespaceBinding, attrs: MetaData): Seq[(String, String)] =
      SaxSupport.combinedAttributes(scope, attrs)
