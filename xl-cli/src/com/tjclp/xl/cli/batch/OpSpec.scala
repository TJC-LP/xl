package com.tjclp.xl.cli.batch

import java.util.Locale

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cli.helpers.BatchParser.BatchOp

/** The JSON shape of one batch-op property (ADR-017 §2.6). */
enum FieldKind derives CanEqual:
  case Str, Num, Int, Bool, Ref, Range, Sheet, Color, NumFmt, Formula, Value, Values, Formulas
  case Enum(values: Vector[String])
  case Arr(of: FieldKind)

/**
 * One property of a batch op. `name` is the canonical camelCase key (today's JSON key); `aliases`
 * are the other accepted names (`numFormat` for `format`, `anchor` for `from`, …). Every name is
 * also accepted in its kebab-case and camelCase spelling, silently and without a warning.
 */
final case class Field(
  name: String,
  kind: FieldKind,
  required: Boolean,
  aliases: Vector[String] = Vector.empty,
  doc: String = ""
) derives CanEqual:

  /** Lookup order: the canonical name, then each alias, each in its given, kebab and camel form. */
  def lookupOrder: Vector[String] = (name +: aliases).flatMap(Spelling.variants).distinct

  /** Every key this field answers to. */
  def spellings: Set[String] = lookupOrder.toSet

/** Key-spelling conversions shared by the registry and the parser. Pure. */
object Spelling:

  /** `fitToWidth` → `fit-to-width`; already-kebab or single-word names are unchanged. */
  def kebab(s: String): String =
    s.replaceAll("([a-z0-9])([A-Z])", "$1-$2").toLowerCase(Locale.ROOT)

  /** `fit-to-width` → `fitToWidth`; names without a hyphen are unchanged. */
  def camel(s: String): String = s.split('-').toList match
    case head :: tail => head + tail.map(_.capitalize).mkString
    case Nil => s

  /** The name itself plus its kebab and camel forms, de-duplicated, given form first. */
  def variants(s: String): Vector[String] = Vector(s, kebab(s), camel(s)).distinct

  /** Op-name normalisation for [[OpRegistry.find]]: case- and separator-insensitive. */
  def normalize(s: String): String =
    s.toLowerCase(Locale.ROOT).replace("-", "").replace("_", "")

/**
 * The specification of one batch op name (ADR-017 §2.6): its fields, which of them are mutually
 * exclusive (`oneOf`), whether it takes the `sheet` key, whether it changes cell content (mirrors
 * `WriteCommands.isCellMutating`), whether the streaming writer can apply it (mirrors the supported
 * set of `StreamingWriteCommands.buildStreamingBatchPatches`), and a valid `example`.
 */
final case class OpSpec(
  name: String,
  aliases: Vector[String],
  fields: Vector[Field],
  oneOf: Vector[Set[String]],
  sheetScoped: Boolean,
  cellMutating: Boolean,
  structural: Boolean,
  needsFormula: Boolean,
  streamable: Boolean,
  cliVerb: Option[String],
  since: String,
  doc: String,
  example: ujson.Obj
) derives CanEqual:

  def field(name: String): Option[Field] = fields.find(_.name == name)

  /**
   * The canonical key a JSON key resolves to: `op`; `sheet` when the op is sheet-scoped; a field's
   * name through any of its spellings; None for a key the op does not know.
   */
  def canonicalName(key: String): Option[String] =
    if key == "op" then Some("op")
    else if key == "sheet" && sheetScoped then Some("sheet")
    else fields.find(_.spellings.contains(key)).map(_.name)

  /** The canonical property names, for the unknown-property warning's `Known:` list. */
  def knownProperties: Vector[String] =
    val scope = if sheetScoped then Vector("sheet") else Vector.empty
    ("op" +: scope) ++ fields.map(_.name)

/**
 * Where a value's number format came from (ADR-017 invariant 4). Only an `Explicit` hint — the op's
 * `format` key — may override a non-General numFmt (GH-560); an `Inferred` one (a detected date or
 * currency string) keeps `Sheet.put`'s General-only merge rule.
 */
enum FormatHint derives CanEqual:
  case Inferred, Explicit

/**
 * One parsed op with the sheet its `sheet` key named (None = the batch default), its 1-based
 * position in the document, and how its value's number format arose. Every op gets a sheet without
 * touching the `BatchOp` cases the specs construct.
 */
final case class ScopedOp(
  op: BatchOp,
  sheet: Option[SheetName],
  index: Int,
  hint: FormatHint = FormatHint.Inferred
) derives CanEqual
