package com.tjclp.xl.sheets

import com.tjclp.xl.addressing.CellRange

/**
 * Data validation (GH-375): one worksheet `<dataValidation>` entry.
 *
 * Document order in `Sheet.dataValidations` is emission order inside the single `<dataValidations>`
 * container. The typed envelope (sqref ranges) shifts under structural edits even while foreign
 * entries ride through as [[DataValidation.Preserved]] — the [[com.tjclp.xl.cf.ConditionalFormat]]
 * pattern.
 */
enum DataValidation derives CanEqual:
  /**
   * One typed validation: `kind` applies to every range in `ranges`.
   *
   * An entry with empty `ranges` is unexpressible in OOXML and is dropped at emission (the
   * authoring API cannot construct one).
   *
   * @param allowBlank
   *   Blank cells always pass validation (Excel's "Ignore blank", `allowBlank="1"`). The OOXML
   *   schema default is false; Excel's UI default is true.
   * @param showDropdown
   *   Show the in-cell dropdown arrow for list validations. NOTE the OOXML INVERSION: the
   *   serialized attribute `showDropDown="1"` SUPPRESSES the dropdown, so this friendly field emits
   *   the attribute only when false.
   */
  case Rules(
    ranges: Vector[CellRange],
    kind: DvKind,
    allowBlank: Boolean = true,
    showDropdown: Boolean = true
  )

  /**
   * Whole-entry fallback for an unmodeled `<dataValidation>` (whole/decimal/date/custom types,
   * operators, prompt/error messages, unknown attrs). CONTRACT =
   * [[com.tjclp.xl.cf.ConditionalFormat.Preserved]]: constructed only by xl-ooxml's reader; the
   * payload is the scope-self-contained canonical XML of one whole element, re-emitted verbatim.
   * Users must not construct this case; a hand-built payload that is not canonical XML is silently
   * dropped at emission.
   */
  case Preserved(xml: String)

object DataValidation:
  /**
   * List-dropdown validation. `formula` is stored VERBATIM as the `<formula1>` text: either an
   * inline list with literal quotes (`"\"Low,Med,High\""`) or a range reference (`"$Z$1:$Z$3"`,
   * `"Lists!$A$1:$A$5"`). Attach ranges via `Sheet.withDataValidation`.
   */
  def list(
    formula: String,
    allowBlank: Boolean = true,
    showDropdown: Boolean = true
  ): DataValidation.Rules =
    Rules(Vector.empty, DvKind.List(formula), allowBlank, showDropdown)

  /**
   * List-dropdown validation from inline values: `listOf("Yes", "No")` stores `"Yes,No"` (the
   * quoted comma-separated form Excel uses). Quotes in values are doubled per formula
   * string-literal rules; values containing commas are inexpressible inline — use a range reference
   * via [[list]] instead.
   */
  def listOf(first: String, rest: String*): DataValidation.Rules =
    val joined = (first +: rest).map(_.replace("\"", "\"\"")).mkString(",")
    list("\"" + joined + "\"")

/**
 * Validation kind for [[DataValidation.Rules]] — minimal but extensible (GH-375): list dropdowns
 * first; whole/decimal/date/textLength bounds are future cases. Formula text is stored verbatim,
 * never parsed (the [[com.tjclp.xl.cf.CfRule]] convention).
 */
enum DvKind derives CanEqual:
  /** `type="list"`: `formula` is the `<formula1>` text (inline quoted list or range reference). */
  case List(formula: String)
