package com.tjclp.xl.styles.font

/**
 * Underline style (ECMA-376 Part 1, §18.8.40 `<u>` / §18.18.86 ST_UnderlineValues).
 *
 * `SingleAccounting`/`DoubleAccounting` are Excel's accounting variants (drawn across the full cell
 * width, e.g. the "Actual / Est. / Projections" period-role band on financial models);
 * `Single`/`Double` underline the text only. `None` means no underline — no `<u/>` is emitted.
 */
enum Underline derives CanEqual:
  case None, Single, Double, SingleAccounting, DoubleAccounting

object Underline:
  /**
   * Canonical ST_UnderlineValues token (ECMA-376 Part 1, §18.18.86). Explicit total mapping — the
   * schema tokens are camelCase (`singleAccounting`), which no `toString` transformation of the
   * enum case names produces reliably (the GH-287 border-token precedent).
   */
  def token(u: Underline): String = u match
    case Underline.None => "none"
    case Underline.Single => "single"
    case Underline.Double => "double"
    case Underline.SingleAccounting => "singleAccounting"
    case Underline.DoubleAccounting => "doubleAccounting"

  /** Strict inverse of [[token]]: unknown tokens yield `scala.None`. */
  def fromToken(token: String): Option[Underline] = token match
    case "none" => Some(Underline.None)
    case "single" => Some(Underline.Single)
    case "double" => Some(Underline.Double)
    case "singleAccounting" => Some(Underline.SingleAccounting)
    case "doubleAccounting" => Some(Underline.DoubleAccounting)
    case _ => scala.None
