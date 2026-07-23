package com.tjclp.xl.ooxml

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.FormulaKind

/**
 * String-level codec between CT_CellFormula record attributes and [[FormulaKind]] (GH-430).
 *
 * One codec serves the DOM reader/writers, the direct SAX emitter, and the streaming
 * readers/writers, so every emitter renders a record identically — the determinism guarantee.
 * Output uses the fixed CT_CellFormula schema attribute order `t, aca, ref, dt2D, dtr, del1, del2,
 * r1, r2, ca` with `1` for set booleans, which reproduces Excel-written records byte-for-byte.
 */
private[xl] object FormulaKindCodec:

  /**
   * OOXML xsd:boolean lexical space as observed in the field: Excel writes `1`, LibreOffice writes
   * `true`/`false`. Junk lexicals deterministically read as false (the schema default).
   */
  def parseBoolean(raw: String): Boolean =
    raw.trim match
      case "1" | "true" => true
      case _ => false

  private def boolAttr(get: String => Option[String], name: String): Boolean =
    get(name).exists(parseBoolean)

  /**
   * Recognize a modeled non-shared record from `<f>` attributes.
   *
   * Returns `Some` only for `t="array"` / `t="dataTable"` records whose load-bearing `ref` parses;
   * anything else (absent/`normal`/`shared`/junk `t`, missing or corrupt `ref`) yields `None` and
   * the caller falls back to plain-formula behavior — lenient-total, never throws, never invents a
   * range. Unparseable `r1`/`r2` degrade to `None` (attr dropped on re-emit; the `ref` is the
   * load-bearing attribute).
   */
  def fromAttrs(t: Option[String], get: String => Option[String]): Option[FormulaKind] =
    t match
      case Some("array") =>
        get("ref").flatMap(CellRange.parse(_).toOption).map { range =>
          FormulaKind.ArrayFormula(range, aca = boolAttr(get, "aca"), ca = boolAttr(get, "ca"))
        }
      case Some("dataTable") =>
        get("ref").flatMap(CellRange.parse(_).toOption).map { range =>
          FormulaKind.DataTable(
            ref = range,
            dt2D = boolAttr(get, "dt2D"),
            dtr = boolAttr(get, "dtr"),
            r1 = get("r1").flatMap(ARef.parse(_).toOption),
            r2 = get("r2").flatMap(ARef.parse(_).toOption),
            del1 = boolAttr(get, "del1"),
            del2 = boolAttr(get, "del2"),
            ca = boolAttr(get, "ca")
          )
        }
      case _ => None

  /**
   * Render a record's `<f>` attributes in fixed CT_CellFormula schema order. `Normal` renders none.
   * False flags are omitted, except `dt2D`/`dtr` which are always explicit (`"1"`/`"0"`) on a data
   * table record, matching Excel's own output.
   */
  def toAttrs(kind: FormulaKind): List[(String, String)] =
    kind match
      case FormulaKind.Normal => Nil
      case FormulaKind.ArrayFormula(ref, aca, ca) =>
        List("t" -> "array")
          ++ (if aca then List("aca" -> "1") else Nil)
          ++ List("ref" -> ref.toA1)
          ++ (if ca then List("ca" -> "1") else Nil)
      case FormulaKind.DataTable(ref, dt2D, dtr, r1, r2, del1, del2, ca) =>
        List(
          "t" -> "dataTable",
          "ref" -> ref.toA1,
          "dt2D" -> (if dt2D then "1" else "0"),
          "dtr" -> (if dtr then "1" else "0")
        )
          ++ (if del1 then List("del1" -> "1") else Nil)
          ++ (if del2 then List("del2" -> "1") else Nil)
          ++ r1.map(r => "r1" -> r.toA1).toList
          ++ r2.map(r => "r2" -> r.toA1).toList
          ++ (if ca then List("ca" -> "1") else Nil)
