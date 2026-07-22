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
   * @param messages
   *   Input-prompt and error-alert configuration (GH-429). Excel stamps
   *   `showInputMessage="1" showErrorMessage="1"` on virtually every validation it writes, so the
   *   typed model must carry them or every Excel-authored validation degrades to [[Preserved]] and
   *   goes structurally inert.
   */
  case Rules(
    ranges: Vector[CellRange],
    kind: DvKind,
    allowBlank: Boolean = true,
    showDropdown: Boolean = true,
    messages: DvMessages = DvMessages.default
  )

  /**
   * Whole-entry fallback for an unmodeled `<dataValidation>` (unknown type tokens, foreign or
   * prefixed attrs such as `xr:uid`/`imeMode`, unexpected children). CONTRACT =
   * [[com.tjclp.xl.cf.ConditionalFormat.Preserved]]: constructed only by xl-ooxml's reader — or
   * derived from a reader-constructed payload by the structural envelope rewrite
   * ([[SqrefShift.shiftPayload]]), which preserves canonicality; the payload is the
   * scope-self-contained canonical XML of one whole element, re-emitted verbatim. Users must not
   * construct this case; a hand-built payload that is not canonical XML is silently dropped at
   * emission.
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

  /** Custom-formula validation (`type="custom"`): `formula` is the `<formula1>` text, verbatim. */
  def custom(formula: String): DataValidation.Rules =
    Rules(Vector.empty, DvKind.Custom(formula))

  /**
   * "Any value" validation (`type` absent or `"none"`): no constraint, but the entry can still
   * carry an input prompt / error alert via [[withPrompt]] / [[withError]] — Excel's "message-only"
   * validations.
   */
  def anyValue: DataValidation.Rules =
    Rules(Vector.empty, DvKind.AnyValue)

  /** Whole-number bound validation (`type="whole"`). Formula text is stored verbatim. */
  def whole(
    op: DvOperator,
    formula1: String,
    formula2: Option[String] = None
  ): DataValidation.Rules =
    Rules(Vector.empty, DvKind.Bounded(DvBoundedType.Whole, op, formula1, formula2))

  /** Decimal bound validation (`type="decimal"`). */
  def decimal(
    op: DvOperator,
    formula1: String,
    formula2: Option[String] = None
  ): DataValidation.Rules =
    Rules(Vector.empty, DvKind.Bounded(DvBoundedType.Decimal, op, formula1, formula2))

  /** Date bound validation (`type="date"`). */
  def date(
    op: DvOperator,
    formula1: String,
    formula2: Option[String] = None
  ): DataValidation.Rules =
    Rules(Vector.empty, DvKind.Bounded(DvBoundedType.Date, op, formula1, formula2))

  /** Time bound validation (`type="time"`). */
  def time(
    op: DvOperator,
    formula1: String,
    formula2: Option[String] = None
  ): DataValidation.Rules =
    Rules(Vector.empty, DvKind.Bounded(DvBoundedType.Time, op, formula1, formula2))

  /** Text-length bound validation (`type="textLength"`). */
  def textLength(
    op: DvOperator,
    formula1: String,
    formula2: Option[String] = None
  ): DataValidation.Rules =
    Rules(Vector.empty, DvKind.Bounded(DvBoundedType.TextLength, op, formula1, formula2))

  extension (rules: DataValidation.Rules)
    /**
     * Attach an input prompt shown when a validated cell is selected. Attach-implies-show (Excel-UI
     * parity): sets `showInputMessage = true`.
     */
    def withPrompt(title: String, body: String): DataValidation.Rules =
      rules.copy(messages =
        rules.messages
          .copy(showInputMessage = true, promptTitle = Some(title), prompt = Some(body))
      )

    /**
     * Attach an error alert shown when validation fails. Attach-implies-show (Excel-UI parity):
     * sets `showErrorMessage = true`. `style` maps to `errorStyle` (Stop blocks entry, Warning and
     * Information allow overriding).
     */
    def withError(
      title: String,
      body: String,
      style: DvErrorStyle = DvErrorStyle.Stop
    ): DataValidation.Rules =
      rules.copy(messages =
        rules.messages.copy(
          showErrorMessage = true,
          errorTitle = Some(title),
          error = Some(body),
          errorStyle = style
        )
      )

/**
 * Validation kind for [[DataValidation.Rules]] (GH-375, widened by GH-429 to cover what Excel
 * actually writes). Formula text is stored verbatim, never parsed (the [[com.tjclp.xl.cf.CfRule]]
 * convention).
 */
enum DvKind derives CanEqual:
  /** `type="list"`: `formula` is the `<formula1>` text (inline quoted list or range reference). */
  case List(formula: String)

  /** `type="custom"`: `formula` is the boolean `<formula1>` expression, verbatim. */
  case Custom(formula: String)

  /** `type` absent or `"none"`: no constraint (message-only validations). */
  case AnyValue

  /**
   * The bounded families (`whole`/`decimal`/`date`/`time`/`textLength`): `operator` compares the
   * cell against `formula1` (and `formula2` for Between/NotBetween). Absent operator = Between (the
   * OOXML schema default).
   */
  case Bounded(
    dataType: DvBoundedType,
    operator: DvOperator,
    formula1: String,
    formula2: Option[String]
  )

/** Bounded validation data types — the OOXML `type` tokens for [[DvKind.Bounded]]. */
enum DvBoundedType derives CanEqual:
  case Whole, Decimal, Date, Time, TextLength

/** Comparison operator for [[DvKind.Bounded]]; the OOXML schema default is Between. */
enum DvOperator derives CanEqual:
  case Between, NotBetween, Equal, NotEqual, LessThan, LessThanOrEqual, GreaterThan,
    GreaterThanOrEqual

/** Error-alert severity (`errorStyle`); the OOXML schema default is Stop. */
enum DvErrorStyle derives CanEqual:
  case Stop, Warning, Information

/**
 * Input-prompt and error-alert configuration for one [[DataValidation.Rules]] entry (GH-429).
 *
 * Defaults equal the OOXML schema defaults, so absent attributes parse to [[DvMessages.default]]
 * and the default emits no attributes. Text fields are free-form (no length `require`s — totality
 * over cleverness); multiline text round-trips through `_x000A_` attribute escapes.
 */
final case class DvMessages(
  showInputMessage: Boolean = false,
  showErrorMessage: Boolean = false,
  promptTitle: Option[String] = None,
  prompt: Option[String] = None,
  errorTitle: Option[String] = None,
  error: Option[String] = None,
  errorStyle: DvErrorStyle = DvErrorStyle.Stop
) derives CanEqual

object DvMessages:
  val default: DvMessages = DvMessages()
