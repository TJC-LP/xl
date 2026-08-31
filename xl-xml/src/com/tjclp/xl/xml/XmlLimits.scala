package com.tjclp.xl.xml

/**
 * Security limits enforced by [[XmlPullParser]] — the portable port of xl-ooxml's XmlSecurity
 * posture. Every limit is finite; there is deliberately no "unlimited" configuration.
 *
 *   - `maxDepth` — maximum element nesting depth (enforced at each start-tag push).
 *   - `maxAttrsPerElement` — maximum attributes on one element (xmlns declarations count).
 *   - `maxNameLength` — maximum element/attribute name length in characters; the default mirrors
 *     JAXP's `jdk.xml.maxXMLNameLimit` default (1000) for differential parity with the JAXP path.
 *   - `maxTotalChars` — ceiling on total decoded characters, mirroring XmlSecurity's GH-457
 *     entity-size ceiling (2e9). DOCTYPE skipping is bounded by this counter, which retires the
 *     16KB prolog-scan truncation weakness of the GH-350 stream stripper.
 *
 * Violations surface as [[XmlParseError]] values naming the tripped limit, with position.
 */
final case class XmlLimits(
  maxDepth: Int = 256,
  maxAttrsPerElement: Int = 512,
  maxNameLength: Int = 1000,
  maxTotalChars: Long = 2_000_000_000L
) derives CanEqual

object XmlLimits:
  val default: XmlLimits = XmlLimits()
