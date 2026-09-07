package com.tjclp.xl.cli.output

/**
 * The three text escapers of the CLI, each defined exactly once (W2.4): every JSON string the
 * renderers and typed payloads emit goes through [[json]], every CSV field through [[csv]], every
 * markdown table cell through [[markdown]].
 */
object Escape:

  /** A JSON string literal, quotes included, per RFC 8259. */
  def json(s: String): String =
    val sb = new StringBuilder
    sb.append('"')
    s.foreach {
      case '"' => sb.append("\\\"")
      case '\\' => sb.append("\\\\")
      case '\n' => sb.append("\\n")
      case '\r' => sb.append("\\r")
      case '\t' => sb.append("\\t")
      case '\b' => sb.append("\\b")
      case '\f' => sb.append("\\f")
      case c if c < 32 => sb.append(f"\\u${c.toInt}%04x")
      case c => sb.append(c)
    }
    sb.append('"')
    sb.toString

  /**
   * An RFC 4180 CSV field: wrapped in quotes when it holds a comma, a quote or a line break, with
   * inner quotes doubled; otherwise verbatim.
   */
  def csv(s: String): String =
    val needsQuoting = s.contains(',') || s.contains('"') || s.contains('\n') || s.contains('\r')
    if needsQuoting then s"\"${s.replace("\"", "\"\"")}\"" else s

  /** A markdown table cell: pipes escaped, line breaks flattened. */
  def markdown(s: String): String =
    s.replace("|", "\\|").replace("\n", " ").replace("\r", "")
