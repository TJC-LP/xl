package com.tjclp.xl.ooxml

import com.tjclp.xl.addressing.{ARef, CellRange}
import com.tjclp.xl.cells.{ArrayMode, FormulaKind}

/**
 * The resolved `cm` (cell metadata) indices of a package (GH-714): the 1-based `cellMetadata/bk`
 * index of every block that names XLDAPR dynamic-array properties with `fDynamic` set, mapped to
 * the dynamic mode it describes. Built once per package from `xl/metadata.xml` and shared by the
 * DOM and streaming readers; an index absent from the map is not a dynamic-array marker.
 */
private[xl] final case class CellMetadataIndex(byCm: Map[Int, ArrayMode.Dynamic])

private[xl] object CellMetadataIndex:
  val empty: CellMetadataIndex = CellMetadataIndex(Map.empty)

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
   * Returns `Some` for `t="array"` / `t="dataTable"` records whose load-bearing `ref` parses, and
   * (GH-435) for any other non-shared `<f>` that carries a set `ca`/`aca` — those two flags are
   * legal on every CT_CellFormula, so a plain `<f ca="1">` keeps them. A record whose `ref` is
   * missing or corrupt falls back to that flagged plain formula instead of dropping the flags with
   * the range; `None` means "nothing to model" and the caller uses plain-formula behavior —
   * lenient-total, never throws, never invents a range. Unparseable `r1`/`r2` degrade to `None`
   * (attr dropped on re-emit; the `ref` is the load-bearing attribute). `t="shared"` models
   * nothing: those expand per GH-370, so both the DOM and streaming readers rebuild the text and
   * carry no record.
   */
  def fromAttrs(t: Option[String], get: String => Option[String]): Option[FormulaKind] =
    val aca = boolAttr(get, "aca")
    val ca = boolAttr(get, "ca")
    val plain = Option.when(aca || ca)(FormulaKind.Normal(aca = aca, ca = ca))
    t match
      case Some("array") =>
        get("ref")
          .flatMap(CellRange.parse(_).toOption)
          .map(range => FormulaKind.ArrayFormula(range, aca = aca, ca = ca))
          .orElse(plain)
      case Some("dataTable") =>
        get("ref")
          .flatMap(CellRange.parse(_).toOption)
          .map { range =>
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
          .orElse(plain)
      case Some("shared") => None
      case _ => plain

  /**
   * Render a record's `<f>` attributes in fixed CT_CellFormula schema order. `Normal` renders only
   * the calc flags it carries (none for the common unflagged case). False flags are omitted, except
   * `dt2D`/`dtr` which are always explicit (`"1"`/`"0"`) on a data table record, matching Excel's
   * own output. A 1x1 data-table interior renders a BARE single-cell `ref` (`ref="Q2"`,
   * fixture-verified — Excel never writes `Q2:Q2`); the ArrayFormula arm keeps the range form
   * deliberately (`E10:E10` is pinned by FormulaRecordPreservationSpec; revisit only if Excel's own
   * single-cell CSE output is ever verified to differ).
   */
  def toAttrs(kind: FormulaKind): List[(String, String)] =
    kind match
      case FormulaKind.Normal(aca, ca) =>
        (if aca then List("aca" -> "1") else Nil) ++ (if ca then List("ca" -> "1") else Nil)
      case FormulaKind.ArrayFormula(ref, aca, ca, mode) =>
        // GH-714: Excel writes a 1x1 dynamic array with a bare ref (`ref="C5"`); CSE keeps the pin
        val refText = mode match
          case _: ArrayMode.Dynamic if ref.start == ref.end => ref.start.toA1
          case _ => ref.toA1
        List("t" -> "array")
          ++ (if aca then List("aca" -> "1") else Nil)
          ++ List("ref" -> refText)
          ++ (if ca then List("ca" -> "1") else Nil)
      case FormulaKind.DataTable(ref, dt2D, dtr, r1, r2, del1, del2, ca) =>
        List(
          "t" -> "dataTable",
          "ref" -> (if ref.start == ref.end then ref.start.toA1 else ref.toA1),
          "dt2D" -> (if dt2D then "1" else "0"),
          "dtr" -> (if dtr then "1" else "0")
        )
          ++ (if del1 then List("del1" -> "1") else Nil)
          ++ (if del2 then List("del2" -> "1") else Nil)
          ++ r1.map(r => "r1" -> r.toA1).toList
          ++ r2.map(r => "r2" -> r.toA1).toList
          ++ (if ca then List("ca" -> "1") else Nil)

  /**
   * GH-714: upgrade an array record to a dynamic array when the cell's `cm` attribute resolves,
   * through the package's metadata part, to XLDAPR properties with `fDynamic` set. Anything else —
   * no `cm`, a junk or out-of-range index, a non-XLDAPR block, `fDynamic=0`, a missing part, or a
   * `cm` on a record that is not an array — leaves the kind exactly as read (lenient-total).
   */
  def withCellMetadata(
    kind: Option[FormulaKind],
    cm: Option[String],
    idx: CellMetadataIndex
  ): Option[FormulaKind] =
    kind match
      case Some(arr: FormulaKind.ArrayFormula) =>
        cm.flatMap(_.trim.toIntOption).flatMap(idx.byCm.get) match
          case Some(dynamic) => Some(arr.copy(mode = dynamic))
          case None => kind
      case _ => kind

  /**
   * The `<c>` attributes a record contributes (GH-714): `cm` for a dynamic array, from the index
   * the writer's metadata plan allocated (`None` = no part to point at, so the record is written in
   * its legacy shape rather than with a dangling `cm`); nothing for every other kind.
   */
  def cellAttrs(kind: FormulaKind, cmOf: ArrayMode.Dynamic => Option[Int]): List[(String, String)] =
    cellMetadataOf(kind, cmOf).map(n => "cm" -> n.toString).toList

  /** The `cm` index a record is written with: [[cellAttrs]] as a number. */
  def cellMetadataOf(kind: FormulaKind, cmOf: ArrayMode.Dynamic => Option[Int]): Option[Int] =
    kind match
      case FormulaKind.ArrayFormula(_, _, _, dynamic: ArrayMode.Dynamic) => cmOf(dynamic)
      case _ => None

  /** No metadata part: dynamic records are written without `cm`. */
  val noCellMetadata: ArrayMode.Dynamic => Option[Int] = _ => None
