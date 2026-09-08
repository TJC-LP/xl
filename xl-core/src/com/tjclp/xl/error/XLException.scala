package com.tjclp.xl.error

/**
 * Exception wrapper for [[XLError]] providing structured error data in exception-based code.
 *
 * This class bridges the pure Either-based error model with exception-based APIs (like Easy Mode).
 * The underlying [[XLError]] is preserved for programmatic error recovery and structured logging.
 *
 * '''Note''': Constructor is `private[xl]` to ensure exceptions are only created by the library
 * itself through the `.unsafe` boundary. External code should use `Either[XLError, A]` directly.
 *
 * '''Example: Catching and inspecting errors'''
 * {{{
 * import com.tjclp.xl.*
 *
 * try {
 *   sheet.put("InvalidRef", "Value").unsafe  // .unsafe throws XLException
 * } catch {
 *   case ex: XLException =>
 *     ex.error match {
 *       case XLError.InvalidCellRef(ref, reason) => println(s"Bad ref: \$ref - \$reason")
 *       case other => println(s"Error: \${other.message}")
 *     }
 * }
 * }}}
 *
 * '''Detail messages''' (GH-589): the sync facade's edge methods may say more than the error
 * renders on its own — `Excel.readSheet` names the nearest sheet names and every available sheet —
 * while keeping the structured `error` (and so its `code`/`hint`) exactly what the domain returned.
 * The one-argument form is the error's own message, byte for byte.
 *
 * @param error
 *   The underlying structured error
 * @param detail
 *   The exception message; defaults to `error.message`
 */
final class XLException private[xl] (val error: XLError, detail: String)
    extends RuntimeException(detail):
  private[xl] def this(error: XLError) = this(error, error.message)
