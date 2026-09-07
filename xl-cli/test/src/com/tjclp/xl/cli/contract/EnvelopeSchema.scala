package com.tjclp.xl.cli.contract

import java.nio.charset.StandardCharsets
import java.nio.file.Files

import munit.Assertions.fail

/**
 * `xl-cli/test/resources/schema/envelope.schema.json` — the published shape of the `--json`
 * envelope — plus a hand-written structural checker over `ujson.Value` for the JSON Schema subset
 * the file uses: `type` (a name or a list of names), `enum`, `pattern`, `required`, `properties`,
 * `additionalProperties: false`, `items`, `oneOf` and local `$ref`s (`#/$defs/...`). No schema
 * library: the contract suite depends on nothing xl-cli itself does not.
 */
object EnvelopeSchema:

  lazy val schema: ujson.Value =
    val path = Golden.repoRoot.resolve("xl-cli/test/resources/schema/envelope.schema.json")
    ujson.read(Files.readString(path, StandardCharsets.UTF_8))

  /** Every violation, as `<json pointer>: <problem>`; empty when the value conforms. */
  def validate(value: ujson.Value): Vector[String] = SchemaCheck.check(schema, schema, value, "")

  /** Fail the enclosing test with every violation listed. */
  def assertValid(value: ujson.Value): Unit =
    val problems = validate(value)
    if problems.nonEmpty then
      fail(
        s"envelope does not conform to envelope.schema.json:\n${problems.mkString("\n")}\n" +
          s"value:\n${ujson.write(value, indent = 2)}"
      )

object SchemaCheck:

  def check(
    root: ujson.Value,
    schema: ujson.Value,
    value: ujson.Value,
    at: String
  ): Vector[String] =
    schema match
      case obj: ujson.Obj =>
        obj.value.get("$ref") match
          case Some(ujson.Str(ref)) => check(root, resolve(root, ref), value, at)
          case _ =>
            val problems = Vector.newBuilder[String]
            obj.value.get("type").foreach { t =>
              val allowed = t match
                case ujson.Str(name) => Vector(name)
                case ujson.Arr(names) => names.toVector.collect { case ujson.Str(n) => n }
                case _ => Vector.empty
              if !allowed.exists(hasType(value, _)) then
                problems += s"$at: expected type ${allowed.mkString("|")}, got ${typeName(value)}"
            }
            obj.value.get("enum").foreach {
              case ujson.Arr(options) if !options.contains(value) =>
                problems += s"$at: ${ujson.write(value)} is not one of ${ujson.write(ujson.Arr(options))}"
              case _ => ()
            }
            obj.value.get("pattern").foreach {
              case ujson.Str(regex) =>
                value match
                  case ujson.Str(s) if !s.matches(regex) =>
                    problems += s"$at: '$s' does not match $regex"
                  case _ => ()
              case _ => ()
            }
            obj.value.get("oneOf").foreach {
              case ujson.Arr(alternatives) =>
                val matching =
                  alternatives.count(alt => check(root, alt, value, at).isEmpty)
                if matching != 1 then
                  problems += s"$at: matched $matching of ${alternatives.size} oneOf alternatives"
              case _ => ()
            }
            value match
              case ujson.Obj(fields) =>
                val properties = obj.value.get("properties") match
                  case Some(ujson.Obj(p)) => p
                  case _ => scala.collection.mutable.LinkedHashMap.empty[String, ujson.Value]
                obj.value.get("required").foreach {
                  case ujson.Arr(names) =>
                    names.collect { case ujson.Str(n) if !fields.contains(n) => n }.foreach { n =>
                      problems += s"$at: missing required key '$n'"
                    }
                  case _ => ()
                }
                if obj.value.get("additionalProperties").contains(ujson.False) then
                  fields.keys.filterNot(properties.contains).foreach { extra =>
                    problems += s"$at: unexpected key '$extra'"
                  }
                properties.foreach { (name, sub) =>
                  fields.get(name).foreach(v => problems ++= check(root, sub, v, s"$at/$name"))
                }
              case ujson.Arr(items) =>
                obj.value.get("items").foreach { itemSchema =>
                  items.zipWithIndex.foreach { (item, i) =>
                    problems ++= check(root, itemSchema, item, s"$at/$i")
                  }
                }
              case _ => ()
            problems.result()
      case ujson.True => Vector.empty
      case ujson.False => Vector(s"$at: schema forbids any value")
      case other => Vector(s"$at: unsupported schema node ${ujson.write(other)}")

  /** `#/a/b` against the root document. */
  private def resolve(root: ujson.Value, ref: String): ujson.Value =
    ref.stripPrefix("#").split("/").toVector.filter(_.nonEmpty).foldLeft(root) { (node, key) =>
      node match
        case ujson.Obj(fields) => fields.getOrElse(key, ujson.False)
        case ujson.Arr(items) => key.toIntOption.flatMap(items.lift).getOrElse(ujson.False)
        case _ => ujson.False
    }

  private def hasType(value: ujson.Value, name: String): Boolean = (value, name) match
    case (_: ujson.Obj, "object") => true
    case (_: ujson.Arr, "array") => true
    case (_: ujson.Str, "string") => true
    case (ujson.Num(d), "number") => !d.isNaN
    case (ujson.Num(d), "integer") => d.isWhole
    case (ujson.True | ujson.False, "boolean") => true
    case (ujson.Null, "null") => true
    case _ => false

  private def typeName(value: ujson.Value): String = value match
    case _: ujson.Obj => "object"
    case _: ujson.Arr => "array"
    case _: ujson.Str => "string"
    case _: ujson.Num => "number"
    case ujson.True | ujson.False => "boolean"
    case ujson.Null => "null"
