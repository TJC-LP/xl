package com.tjclp.xl.codec

import scala.quoted.*

/** The compile-time half of `RowCodec.derived` (GH-614): a record's `@header` texts. */
object RowCodecMacros:

  /**
   * One entry per field of `A` in declaration order: `Some(text)` for a field annotated
   * `@header(text)`, `None` otherwise — aligned with `MirroredElemLabels`. `A` is dealiased first,
   * as `Mirror.ProductOf` does, so a codec derived through a type alias sees the class's
   * annotations. Compilation aborts for a `@header` whose argument is not a string literal or is
   * blank, and when two fields end up with the same header (a field's default header being its
   * name).
   */
  inline def headerOverrides[A]: List[Option[String]] = ${ headerOverridesImpl[A] }

  private def headerOverridesImpl[A: Type](using Quotes): Expr[List[Option[String]]] =
    import quotes.reflect.*
    // Dealiased like Mirror.ProductOf: `type DealAlias = Deal` must read Deal's constructor
    val record = TypeRepr.of[A].dealias.typeSymbol
    val annotationClass = TypeRepr.of[header].typeSymbol
    // The clause Mirror.ProductOf reads: the first parameter clause holding terms — a generic
    // record's leading type-parameter clause is skipped, later clauses are not fields.
    val params = record.primaryConstructor.paramSymss.find(_.exists(_.isTerm)).getOrElse(Nil)

    def abort(param: Symbol, message: String): Nothing =
      val full = s"RowCodec.derived: $message"
      param.pos.fold(report.errorAndAbort(full))(pos => report.errorAndAbort(full, pos))

    def literal(term: Term): Option[String] = term match
      case Literal(StringConstant(s)) => Some(s)
      case NamedArg(_, inner) => literal(inner)
      case Typed(inner, _) => literal(inner)
      case Inlined(_, Nil, inner) => literal(inner)
      case _ => None

    val overrides: List[(String, Option[String])] = params.map { param =>
      val field = s"@header on field '${param.name}' of ${record.name}"
      val text = param.getAnnotation(annotationClass).map {
        case Apply(_, List(arg)) =>
          literal(arg).getOrElse(abort(param, s"$field needs a string literal"))
        case _ => abort(param, s"$field needs a string literal")
      }
      text.filter(_.trim.isEmpty).foreach(_ => abort(param, s"$field is blank"))
      (param.name, text)
    }

    overrides
      .map((name, text) => (name, text.getOrElse(name)))
      .groupMap(_._2)(_._1)
      .toList
      .sortBy(_._1)
      .collectFirst { case (text, names) if names.sizeIs > 1 => (text, names) }
      .foreach { (text, names) =>
        val quoted = names.map(n => s"'$n'")
        val listed = s"${quoted.dropRight(1).mkString(", ")} and ${quoted.takeRight(1).mkString}"
        report.errorAndAbort(
          s"RowCodec.derived: fields $listed of ${record.name} both have header '$text'"
        )
      }

    Expr(overrides.map(_._2))
