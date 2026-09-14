package com.tjclp.xl.codec

import scala.annotation.StaticAnnotation

/**
 * The header text of a record field (GH-614): what `putRowsWithHeader`/`putTable` write above the
 * column and what `readRowsByHeader` matches, when the sheet's header is not the field's name —
 * `Coupon (%)`, `Portfolio Co.`, anything an identifier cannot spell. The field keeps its name in
 * [[RowCodec.fields]] and in every [[RowCodecError]].
 *
 * {{{
 * final case class Deal(@header("Portfolio Co.") portfolioCo: String, @header("Rev (USD m)") rev: BigDecimal)
 *   derives RowCodec
 * }}}
 *
 * `name` must be a string literal, non-blank, and distinct from every other header of the record
 * (an annotated one or a plain field name) — each is checked when the codec is derived, as a
 * compile error. The runtime twin, for a header known only at runtime, is [[RowCodec.withHeaders]].
 */
final class header(val name: String) extends StaticAnnotation
