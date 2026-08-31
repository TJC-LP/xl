package com.tjclp.xl.xml

/**
 * A parse failure as a value (purity charter: no thrown exceptions escape the public API).
 *
 * `line` and `col` are 1-based character positions (the SAXParseException convention), pointing at
 * the position where the error was detected.
 */
final case class XmlParseError(line: Int, col: Int, message: String) derives CanEqual:
  /**
   * Matches XmlSecurity.parseSafe's SAXParseException rendering so user-visible error text does not
   * churn when the portable parser replaces JAXP behind `parseSafe` (wave A3).
   */
  def render: String = s"XML parse error at line $line, column $col: $message"
