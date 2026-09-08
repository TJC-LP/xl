package com.tjclp.xl.cli.contract

import com.tjclp.xl.formula.Arity
import com.tjclp.xl.formula.functions.{FunctionRegistry, FunctionSpec}

/**
 * One row of `xl functions --json` (ADR-017 §2.13): what the evaluator knows about a function,
 * without its implementation. `args` names each argument slot as the parser describes it (`optional
 * text`, `number or range...`); `maxArgs` is `None` for a variadic function. `dynamicDeps` and
 * `volatile` are the evaluator's recalculation flags (`FunctionFlags`): the cells a call reads are
 * decided at evaluation time; the value can change with no input changing. `specialForm` marks
 * `LET`, a parser-level construct the registry does not hold.
 */
final case class FunctionDoc(
  name: String,
  minArgs: Int,
  maxArgs: Option[Int],
  args: List[String],
  returnsDate: Boolean,
  returnsTime: Boolean,
  dynamicDeps: Boolean,
  volatile: Boolean,
  specialForm: Boolean
) derives CanEqual:

  /**
   * `{name, minArgs, maxArgs, args, returnsDate, returnsTime, dynamicDeps, volatile, specialForm}`.
   */
  def toJson: ujson.Obj =
    ujson.Obj(
      "name" -> ujson.Str(name),
      "minArgs" -> ujson.Num(minArgs),
      "maxArgs" -> maxArgs.fold[ujson.Value](ujson.Null)(n => ujson.Num(n)),
      "args" -> ujson.Arr.from(args.map(ujson.Str.apply)),
      "returnsDate" -> ujson.Bool(returnsDate),
      "returnsTime" -> ujson.Bool(returnsTime),
      "dynamicDeps" -> ujson.Bool(dynamicDeps),
      "volatile" -> ujson.Bool(volatile),
      "specialForm" -> ujson.Bool(specialForm)
    )

object FunctionDoc:

  /**
   * `LET(name1, value1, [name2, value2, ...], calculation)`: at least one name/value pair and a
   * final calculation, parsed by `FormulaParser.parseLet` rather than through a registry spec.
   */
  val let: FunctionDoc = FunctionDoc(
    name = "LET",
    minArgs = 3,
    maxArgs = None,
    args = List("name", "value", "name, value...", "calculation"),
    returnsDate = false,
    returnsTime = false,
    dynamicDeps = false,
    volatile = false,
    specialForm = true
  )

  /** The row for one registry spec: arity bounds from [[Arity]], slots from the arg spec. */
  def of(spec: FunctionSpec[?]): FunctionDoc =
    val (min, max) = spec.arity match
      case Arity.Exact(n) => (n, Some(n))
      case Arity.Range(lo, hi) => (lo, Some(hi))
      case Arity.AtLeast(n) => (n, None)
    FunctionDoc(
      name = spec.name.toUpperCase,
      minArgs = min,
      maxArgs = max,
      args = spec.argSpec.describeParts,
      returnsDate = spec.flags.returnsDate,
      returnsTime = spec.flags.returnsTime,
      dynamicDeps = spec.flags.dynamicDeps,
      volatile = spec.flags.volatile,
      specialForm = false
    )

  /**
   * Every registry function sorted by name, then `LET`. `FunctionRegistry.all` is a macro that
   * enumerates the spec traits at compile time; it is expanded once, here.
   */
  lazy val all: Vector[FunctionDoc] =
    FunctionRegistry.all.map(of).sortBy(_.name).toVector :+ let

  /** The `functions` payload: one object per row, in [[all]]'s order. */
  def toJson(docs: Vector[FunctionDoc]): ujson.Arr = ujson.Arr.from(docs.map(_.toJson))
