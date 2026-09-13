package com.tjclp.xl.styles.fill

import java.util.Locale

/** Fill pattern (ECMA-376 Part 1, §18.18.55 `ST_PatternType`): all 19 schema values. */
enum PatternType derives CanEqual:
  case None, Solid, Gray125, Gray0625
  case DarkGray, MediumGray, LightGray
  case DarkHorizontal, DarkVertical, DarkDown, DarkUp
  case DarkGrid, DarkTrellis
  case LightHorizontal, LightVertical, LightDown, LightUp
  case LightGrid, LightTrellis

object PatternType:
  /**
   * Canonical ST_PatternType token. Explicit total mapping — the schema tokens are camelCase
   * (`mediumGray`, `lightTrellis`), which no `toString` transformation of the case names produces
   * reliably (the GH-287 border-token precedent; the streaming style codec once lowercased them,
   * GH-566). Every writer — DOM, SAX and the streaming patcher — spells the token through here.
   */
  def token(p: PatternType): String = p match
    case PatternType.None => "none"
    case PatternType.Solid => "solid"
    case PatternType.Gray125 => "gray125"
    case PatternType.Gray0625 => "gray0625"
    case PatternType.DarkGray => "darkGray"
    case PatternType.MediumGray => "mediumGray"
    case PatternType.LightGray => "lightGray"
    case PatternType.DarkHorizontal => "darkHorizontal"
    case PatternType.DarkVertical => "darkVertical"
    case PatternType.DarkDown => "darkDown"
    case PatternType.DarkUp => "darkUp"
    case PatternType.DarkGrid => "darkGrid"
    case PatternType.DarkTrellis => "darkTrellis"
    case PatternType.LightHorizontal => "lightHorizontal"
    case PatternType.LightVertical => "lightVertical"
    case PatternType.LightDown => "lightDown"
    case PatternType.LightUp => "lightUp"
    case PatternType.LightGrid => "lightGrid"
    case PatternType.LightTrellis => "lightTrellis"

  private val byLowerToken: Map[String, PatternType] =
    values.iterator.map(p => token(p).toLowerCase(Locale.ROOT) -> p).toMap

  /**
   * Lenient inverse of [[token]]: any casing is accepted (`lightUp`, `lightup`, `LIGHTUP`) since
   * files written by pre-GH-287 xl carry lowercased tokens; an unknown token yields `scala.None`.
   */
  def fromToken(token: String): Option[PatternType] =
    byLowerToken.get(token.toLowerCase(Locale.ROOT))
