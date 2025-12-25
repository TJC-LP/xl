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

  /**
   * Extract only namespace declarations that are LOCAL to this element (not inherited from parent).
   *
   * In scala.xml, NamespaceBinding is a linked list that includes inherited parent namespaces. When
   * emitting child elements, we should only declare namespaces that are NEW at this level, not
   * re-declare all ancestor namespaces (which breaks XML structure).
   */
  protected def localNamespaceAttributes(
    scope: NamespaceBinding,
    parentScope: NamespaceBinding
  ): Seq[(String, String)] =
    def loop(ns: NamespaceBinding, acc: List[(String, String)]): List[(String, String)] =
      // Stop when we reach the parent scope (or TopScope/null)
      if ns == null || ns == TopScope || ns == parentScope then acc
      else
        val prefix = Option(ns.prefix).getOrElse("")
        val name = if prefix.isEmpty then "xmlns" else s"xmlns:$prefix"
        loop(ns.parent, (name, ns.uri) :: acc)

    loop(scope, Nil).distinctBy(_._1).sortBy(_._1)

  /**
   * Stream an existing scala.xml Elem through a SaxWriter.
   *
   * For prefixed elements (like `x15ac:absPath`), the namespace is declared via
   * `startElement(qName, uri)`. For child elements, we track the parent's scope to avoid
   * re-declaring inherited namespaces.
   *
   * Key insight: scala.xml's scope chain contains ALL inherited namespaces from ancestors, but when
   * re-serializing, we only need to declare namespaces that are NEW at each level.
   */
  protected def emitElem(writer: SaxWriter, elem: Elem): Unit =
    // Use elem's own scope as parent so we don't emit its inherited namespaces
    emitElemWithParentScope(writer, elem, elem.scope)

  /** Internal helper that tracks parent scope to emit only local namespace declarations */
  private def emitElemWithParentScope(
    writer: SaxWriter,
    elem: Elem,
    parentScope: NamespaceBinding
  ): Unit =
    val qName =
      Option(elem.prefix).filter(_.nonEmpty).map(p => s"$p:${elem.label}").getOrElse(elem.label)
    // getURI returns null for undefined prefixes; Option() wrapper handles this safely
    val nsUri = Option(elem.prefix)
      .flatMap(p => Option(elem.scope).flatMap(sc => Option(sc.getURI(p))))
      .filter(_.nonEmpty)

    nsUri match
      case Some(uri) => writer.startElement(qName, uri)
      case None => writer.startElement(qName)

    // Emit only namespace declarations that are NEW at this element level
    // (not inherited from parent), plus regular attributes from MetaData
    val localNs = localNamespaceAttributes(elem.scope, parentScope)
    val allAttrs = localNs ++ metaDataAttributes(elem.attributes)

    SaxWriter.withAttributes(writer, allAttrs*) {
      elem.child.foreach {
        case e: Elem => emitElemWithParentScope(writer, e, elem.scope)
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
