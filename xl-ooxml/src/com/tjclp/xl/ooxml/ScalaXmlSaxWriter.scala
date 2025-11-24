package com.tjclp.xl.ooxml

import scala.xml.{
  Elem,
  MetaData,
  NamespaceBinding,
  Node,
  Null,
  PrefixedAttribute,
  Text,
  TopScope,
  UnprefixedAttribute
}

/**
 * SAX writer interpreter that builds a scala.xml tree.
 *
 * Useful for compatibility and tests where a pure Elem is still needed but we want to drive the
 * same SAX protocol used by streaming writers.
 */
class ScalaXmlSaxWriter extends SaxWriter:
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private var stack: List[ScalaXmlSaxWriter.ElemState] = Nil
  @SuppressWarnings(Array("org.wartremover.warts.Var"))
  private var root: Option[Elem] = None

  def startDocument(): Unit = ()
  def endDocument(): Unit = ()

  def startElement(name: String): Unit =
    stack = ScalaXmlSaxWriter.ElemState(name, None, Nil, Nil) :: stack

  def startElement(name: String, namespace: String): Unit =
    stack = ScalaXmlSaxWriter.ElemState(name, Some(namespace), Nil, Nil) :: stack

  def writeAttribute(name: String, value: String): Unit =
    stack match
      case head :: tail =>
        stack = head.copy(attributes = (name, value) :: head.attributes) :: tail
      case Nil => ()

  def writeCharacters(text: String): Unit =
    stack match
      case head :: tail =>
        stack = head.copy(children = Text(text) :: head.children) :: tail
      case Nil => ()

  def endElement(): Unit =
    stack match
      case head :: tail =>
        val elem = head.toElem
        tail match
          case parent :: rest =>
            stack = parent.copy(children = elem :: parent.children) :: rest
          case Nil =>
            root = Some(elem)
            stack = Nil
      case Nil => ()

  def flush(): Unit = ()

  /** Final scala.xml Elem built from the SAX event stream */
  def result: Option[Elem] = root

object ScalaXmlSaxWriter:
  private[ooxml] case class ElemState(
    name: String,
    namespace: Option[String],
    attributes: List[(String, String)],
    children: List[Node]
  ):
    def toElem: Elem =
      val attrsMeta = attributes.reverse.foldLeft(Null: MetaData) { case (acc, (k, v)) =>
        k.split(":", 2) match
          case Array(prefix, local) => new PrefixedAttribute(prefix, local, v, acc)
          case _ => new UnprefixedAttribute(k, v, acc)
      }
      val scope =
        namespace.fold[NamespaceBinding](TopScope)(ns => NamespaceBinding(null, ns, TopScope))
      Elem(null, name, attrsMeta, scope, minimizeEmpty = true, children.reverse*)
