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
 * @param error
 *   The underlying structured error
 */
final class XLException private[xl] (val error: XLError) extends RuntimeException(error.message)
