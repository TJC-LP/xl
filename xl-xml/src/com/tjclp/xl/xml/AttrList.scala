package com.tjclp.xl.xml

/**
 * Flat-array attribute list — ONE instance per parser, reused across elements (no per-attribute
 * allocation on the row/cell hot path beyond the decoded value Strings).
 *
 * Aliasing contract: the instance returned by [[XmlPullParser.attrs]] is valid only until the next
 * `next()` call; take [[toVector]] for an owned copy on cold paths. Namespace declarations
 * (`xmlns`, `xmlns:*`) appear verbatim alongside regular attributes (they also feed the parser's
 * namespace stack).
 *
 * Lookups are linear scans — OOXML elements carry at most a handful of attributes.
 */
@SuppressWarnings(Array("org.wartremover.warts.Var", "org.wartremover.warts.While"))
final class AttrList private[xml] ():
  private var attrNames = new Array[String](8)
  private var attrValues = new Array[String](8)
  private var count = 0

  private[xml] def clear(): Unit = count = 0

  private[xml] def add(qName: String, attrValue: String): Unit =
    if count == attrNames.length then
      val grownNames = new Array[String](attrNames.length * 2)
      val grownValues = new Array[String](attrValues.length * 2)
      System.arraycopy(attrNames, 0, grownNames, 0, count)
      System.arraycopy(attrValues, 0, grownValues, 0, count)
      attrNames = grownNames
      attrValues = grownValues
    attrNames(count) = qName
    attrValues(count) = attrValue
    count += 1

  def size: Int = count

  /** Index of `qName`, or -1 when absent. */
  def indexOf(qName: String): Int =
    var i = 0
    var found = -1
    while found < 0 && i < count do
      if attrNames(i) == qName then found = i
      i += 1
    found

  /** Name at `i`; "" when out of range (total: never throws). */
  def name(i: Int): String = if i >= 0 && i < count then attrNames(i) else ""

  /** Value at `i`; "" when out of range (total: never throws). */
  def value(i: Int): String = if i >= 0 && i < count then attrValues(i) else ""

  /** Value of `qName`, matching the `Option(attributes.getValue(_))` SAX call shape. */
  def get(qName: String): Option[String] =
    val i = indexOf(qName)
    if i < 0 then None else Some(attrValues(i))

  /** Owned copy for cold paths. */
  def toVector: Vector[(String, String)] =
    val b = Vector.newBuilder[(String, String)]
    var i = 0
    while i < count do
      b += ((attrNames(i), attrValues(i)))
      i += 1
    b.result()
