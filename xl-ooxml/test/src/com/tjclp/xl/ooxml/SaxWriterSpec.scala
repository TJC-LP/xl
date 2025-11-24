package com.tjclp.xl.ooxml

import munit.FunSuite
import scala.collection.mutable

class SaxWriterSpec extends FunSuite:

  test("withAttributes writes attributes in deterministic order before body") {
    val writer = MockSaxWriter()

    SaxWriter.withAttributes(writer, "z" -> "last", "a" -> "first", "m" -> "mid") {
      writer.writeCharacters("body")
    }

    assertEquals(
      writer.events.toList,
      List("@a=first", "@m=mid", "@z=last", "#body")
    )
  }

  test("SaxSerializable implementations can emit SAX events") {
    val writer = MockSaxWriter()
    Greeting("Ada").writeSax(writer)

    assertEquals(
      writer.events.toList,
      List("startDocument", "<greeting>", "@name=Ada", "#hi", "</>", "endDocument")
    )
  }

  test("startElement with namespace captures URI in event stream") {
    val writer = MockSaxWriter()

    writer.startElement("worksheet", "http://schemas.example/worksheet")
    writer.endElement()

    assertEquals(
      writer.events.toList,
      List("<http://schemas.example/worksheet:worksheet>", "</>")
    )
  }

private class MockSaxWriter extends SaxWriter:
  val events: mutable.ArrayBuffer[String] = mutable.ArrayBuffer.empty

  def startDocument(): Unit = events += "startDocument"
  def endDocument(): Unit = events += "endDocument"
  def startElement(name: String): Unit = events += s"<$name>"
  def startElement(name: String, namespace: String): Unit =
    events += s"<$namespace:$name>"
  def writeAttribute(name: String, value: String): Unit = events += s"@$name=$value"
  def writeCharacters(text: String): Unit = events += s"#$text"
  def endElement(): Unit = events += "</>"
  def flush(): Unit = events += "flush"

private case class Greeting(name: String) extends SaxSerializable:
  def writeSax(writer: SaxWriter): Unit =
    writer.startDocument()
    writer.startElement("greeting")
    SaxWriter.withAttributes(writer, "name" -> name) {
      writer.writeCharacters("hi")
    }
    writer.endElement()
    writer.endDocument()
