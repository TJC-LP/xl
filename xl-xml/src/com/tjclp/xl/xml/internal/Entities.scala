package com.tjclp.xl.xml.internal

/**
 * The five XML 1.0 predefined entities — the ONLY named entities the portable parser honors.
 * General/parameter entities are structurally impossible (stronger than JAXP's
 * `disallow-doctype-decl`): nothing declared inside a skipped DOCTYPE is ever registered, so any
 * other named reference fails as an undeclared entity.
 */
private[xml] object Entities:
  /** Resolve a predefined entity name to its character, or -1 if not one of the five. */
  def predefined(name: String): Int = name match
    case "amp" => '&'.toInt
    case "lt" => '<'.toInt
    case "gt" => '>'.toInt
    case "apos" => '\''.toInt
    case "quot" => '"'.toInt
    case _ => -1

  /** XML 1.0 Char production for character-reference validation (surrogates are NOT Chars). */
  def isXmlChar(c: Int): Boolean =
    c == 0x9 || c == 0xa || c == 0xd || (c >= 0x20 && c <= 0xd7ff) ||
      (c >= 0xe000 && c <= 0xfffd) || (c >= 0x10000 && c <= 0x10ffff)
