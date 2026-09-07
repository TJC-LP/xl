package com.tjclp.xl.cli.contract

import java.nio.charset.StandardCharsets
import java.nio.file.{Files, Path}

import munit.FunSuite

import com.tjclp.xl.cli.BuildInfo

/**
 * Renders the generated reference pages from `Schema.json` (ADR-017 §2.13). Pure: the JSON in, the
 * markdown out. Nothing here prints a version — the pages describe the contract, and a release bump
 * must not rewrite them.
 */
object DocsGen:

  /** The five pages under `docs/reference/generated/`, by file name. */
  val files: Vector[String] =
    Vector("cli-verbs.md", "batch-ops.md", "functions.md", "exit-codes.md", "error-codes.md")

  val regenerate: String =
    "XL_UPDATE_DOCS=1 ./mill xl-cli.test.testOnly com.tjclp.xl.cli.contract.DocsGenSpec"

  private def header(source: String): String =
    s"<!-- GENERATED from `$source` by DocsGenSpec; do not edit by hand.\n" +
      s"     A diff here is a contract change. Regenerate: $regenerate -->\n"

  /** Markers around the generated block inside a hand-written page. */
  val regionBegin: String = "<!-- generated:functions:begin -->"
  val regionEnd: String = "<!-- generated:functions:end -->"

  def render(schema: ujson.Value): Map[String, String] = Map(
    "cli-verbs.md" -> cliVerbs(schema),
    "batch-ops.md" -> batchOps(schema),
    "functions.md" -> functions(schema),
    "exit-codes.md" -> exitCodes(schema),
    "error-codes.md" -> errorCodes(schema)
  )

  // --- cells -----------------------------------------------------------------------------------

  /** A markdown table cell: pipes escaped, newlines flattened, empty shown as an em dash. */
  private def cell(s: String): String =
    val flat = s.replace("\r", "").replace("\n", " ").replace("|", "\\|").trim
    if flat.isEmpty then "—" else flat

  private def code(s: String): String = if s.isEmpty then "—" else s"`$s`"

  private def table(headers: Vector[String], rows: Vector[Vector[String]]): String =
    val head = headers.mkString("| ", " | ", " |")
    val rule = headers.map(_ => "---").mkString("| ", " | ", " |")
    (head +: rule +: rows.map(_.mkString("| ", " | ", " |"))).mkString("\n") + "\n"

  private def str(v: ujson.Value): String = v match
    case ujson.Str(s) => s
    case ujson.Null => ""
    case other => ujson.write(other)

  private def flag(v: ujson.Value): Boolean = v == ujson.True

  // --- pages -----------------------------------------------------------------------------------

  def cliVerbs(schema: ujson.Value): String =
    val globals = schema("globals").arr.toVector.map { g =>
      Vector(
        code(str(g("name"))),
        g("short").strOpt.fold("—")(code),
        if flag(g("takesValue")) then "yes" else "no",
        cell(str(g("doc")))
      )
    }
    val verbs = schema("verbs").arr.toVector.map { v =>
      val needs = v("needs")
      val needsCell = Vector(
        Option.when(flag(needs("file")))("`-f`"),
        Option.when(flag(needs("sheet")))("`-s`"),
        Option.when(flag(needs("output")))("`-o`/`-i`"),
        Option.when(flag(needs("streaming")))("`--stream`")
      ).flatten
      Vector(
        code(v("path").arr.map(str).mkString(" ")),
        if needsCell.isEmpty then "—" else needsCell.mkString(" "),
        v("exit").arr.map(e => e.num.toInt.toString).mkString(" "),
        v("batchTwin").strOpt.fold("—")(code),
        str(v("since")),
        cell(str(v("summary")))
      )
    }
    header("xl schema --json") +
      "\n# xl verbs\n\n" +
      "Every verb in `xl --help` order, with what it needs and how it can exit. Global flags go\n" +
      "anywhere on the command line, before or after the verb. `needs` reads: `-f` an input\n" +
      "workbook; `-s` ONE sheet (a sheet-qualified ref names it, else `-s`, else the only sheet\n" +
      "of a single-sheet book, else `SHEET_REQUIRED`); `-o`/`-i` an output (or in-place edit);\n" +
      "`--stream` that the verb accepts O(1)-memory streaming. `exit` lists the exit codes the\n" +
      "verb can end with (see [exit-codes.md](exit-codes.md)); `batch twin` is the batch op that\n" +
      "makes the same edit; `since` is the release the verb is documented from. Run\n" +
      "`xl <verb> --help` for a verb's own options and `xl schema --json` for this table as JSON.\n\n" +
      "## Global flags\n\n" +
      table(Vector("flag", "short", "takes a value", "meaning"), globals) +
      "\n## Verbs\n\n" +
      table(Vector("verb", "needs", "exit", "batch twin", "since", "summary"), verbs)

  def batchOps(schema: ujson.Value): String =
    val ops = schema("batchOps")
    val specs = ops("items")("oneOf").arr.toVector
    val summary = specs.map { op =>
      Vector(
        code(str(op("title"))),
        op("x-aliases").arr.map(a => code(str(a))).mkString(", ") match
          case "" => "—"
          case s => s
        ,
        if flag(op("x-streamable")) then "yes" else "no",
        if flag(op("x-cellMutating")) then "yes" else "no",
        if flag(op("x-sheetScoped")) then "yes" else "no",
        op("x-cliVerb").strOpt.fold("—")(code),
        str(op("x-since")),
        cell(str(op("description")))
      )
    }
    val details = specs.map { op =>
      val name = str(op("title"))
      val required = op("required").arr.map(str).toSet
      val properties = op("properties").obj.toVector
      val canonical = properties.filter { (_, p) => !p.obj.contains("x-alias-of") }
      val aliasesOf = properties.collect {
        case (alias, p) if p.obj.contains("x-alias-of") => str(p("x-alias-of")) -> alias
      }
      def typeName(t: ujson.Value): String = t match
        case ujson.Arr(names) => names.map(str).mkString(" \\| ")
        case other => str(other)
      val rows = canonical.map { (key, p) =>
        val kind = p.obj.get("const").map(c => s"const ${ujson.write(c)}").getOrElse {
          val base = typeName(p("type"))
          val enumeration = p.obj.get("enum").map(e => " one of " + e.arr.map(str).mkString(", "))
          val items = p.obj.get("items").map(i => s" of ${typeName(i("type"))}")
          base + items.getOrElse("") + enumeration.getOrElse("")
        }
        val aliases = aliasesOf.collect { case (canon, alias) if canon == key => alias }
        Vector(
          code(key),
          kind,
          if required.contains(key) then "yes" else "no",
          if aliases.isEmpty then "—" else aliases.map(code).mkString(", "),
          cell(str(p.obj.getOrElse("description", ujson.Str(""))))
        )
      }
      val oneOf = op.obj.get("oneOf").fold("") { groups =>
        val named = groups.arr.map { g =>
          g.obj.get("required").map(r => r.arr.map(str).mkString("{", ", ", "}")).getOrElse {
            g("anyOf").arr.headOption
              .flatMap(_.obj.get("required"))
              .map(r => r.arr.map(str).mkString("{", ", ", "}"))
              .getOrElse("{}")
          }
        }
        s"\nExactly one of: ${named.mkString(" \\| ")}.\n"
      }
      s"### `$name`\n\n${str(op("description"))}\n\n```json\n${ujson.write(op("x-example"))}\n```\n\n" +
        table(Vector("property", "type", "required", "aliases", "description"), rows) + oneOf
    }
    header("xl batch --schema") +
      "\n# xl batch operations\n\n" +
      str(ops("description")) + "\n\n" +
      "`xl batch --schema` prints this table as a JSON Schema (draft 2020-12): an array whose\n" +
      "items are exactly one of the per-op object schemas below, discriminated by `op`. Every\n" +
      "property is accepted in camelCase or kebab-case; the aliases column lists the other names\n" +
      "a property answers to. An op the streaming writer cannot apply is refused by index under\n" +
      "`--stream` before any byte is written (`UNSUPPORTED_IN_STREAM`, exit 2). `put`/`putf`\n" +
      "`format` is an explicit hint and replaces the cell's number format; a detected format\n" +
      "(currency, percent, ISO date) only applies to a General cell.\n\n" +
      "## Operations\n\n" +
      table(
        Vector(
          "op",
          "aliases",
          "stream",
          "mutates cells",
          "sheet key",
          "CLI twin",
          "since",
          "summary"
        ),
        summary
      ) +
      "\n## Properties\n\n" +
      details.mkString("\n")

  /** The function table alone — also the block embedded in the skill's FORMULAS.md. */
  def functionTable(schema: ujson.Value): String =
    val rows = schema("functions").arr.toVector.map { f =>
      val min = f("minArgs").num.toInt
      val arity = f("maxArgs") match
        case ujson.Null => s"$min+"
        case max if max.num.toInt == min => min.toString
        case max => s"$min–${max.num.toInt}"
      val flags = Vector(
        Option.when(flag(f("returnsDate")))("date"),
        Option.when(flag(f("returnsTime")))("time"),
        Option.when(flag(f("dynamicDeps")))("dynamic deps"),
        Option.when(flag(f("specialForm")))("special form")
      ).flatten
      Vector(
        code(str(f("name"))),
        arity,
        f("args").arr.map(str).mkString(", ") match
          case "" => "—"
          case s => cell(s)
        ,
        if flags.isEmpty then "—" else flags.mkString(", ")
      )
    }
    table(Vector("function", "args", "arguments", "flags"), rows)

  def functions(schema: ujson.Value): String =
    val count = schema("functions").arr.count(f => !flag(f("specialForm")))
    header("xl functions --json") +
      "\n# Formula functions\n\n" +
      s"$count functions in the evaluator's registry, plus `LET` — a parser-level special form\n" +
      "(lexical bindings) that the registry does not hold. `args` is the accepted argument count\n" +
      "(`n+` = at least n, no upper bound); `arguments` names each slot as the parser describes it\n" +
      "(`optional …` may be omitted, `…...` repeats). `flags`: `date`/`time` — the result is a date\n" +
      "or time (drives number-format inference); `dynamic deps` — the cells read are decided at\n" +
      "evaluation time (INDIRECT/OFFSET), so such cells are always recalculated. Run\n" +
      "`xl functions --json` for this table as JSON, `xl eval \"=F(...)\"` to try one.\n\n" +
      functionTable(schema)

  def exitCodes(schema: ujson.Value): String =
    val rows = schema("exitCodes").arr.toVector.map { e =>
      Vector(code(e("code").num.toInt.toString), cell(str(e("meaning"))))
    }
    header("xl schema --json") +
      "\n# Exit codes\n\n" +
      "Every `xl` invocation ends with one of these four codes; `xl --help` prints the same table.\n" +
      "Exit 1 is reserved for \"completed with findings or a failed gate\" and is never a failure.\n" +
      "Branch on the exit code and the `code:` line (or `error.code` under `--json`), never on\n" +
      "message text.\n\n" +
      table(Vector("exit", "meaning"), rows)

  def errorCodes(schema: ujson.Value): String =
    val errors = schema("errorCodes").arr.toVector.map { e =>
      Vector(code(str(e("code"))), e("exit").num.toInt.toString)
    }
    val warnings = schema("warningCodes").arr.toVector.map(w => Vector(code(str(w))))
    header("xl schema --json") +
      "\n# Error and warning codes\n\n" +
      "The complete vocabulary a `code:` line (text mode) or `error.code` (`--json`) can carry:\n" +
      "the CLI's own codes first, then every domain (`XLError`) code. The exit code follows from\n" +
      "the error code alone. Warnings are `Warning[<CODE>]: <message>` lines on stderr — or\n" +
      "`warnings[]` in the envelope — and never change the exit code.\n\n" +
      "## Error codes\n\n" +
      table(Vector("code", "exit"), errors) +
      "\n## Warning codes\n\n" +
      table(Vector("code"), warnings)

  /** Replace the block between the markers (kept) with `body`; None when a marker is missing. */
  def replaceRegion(text: String, body: String): Option[String] =
    val begin = text.indexOf(regionBegin)
    val end = text.indexOf(regionEnd)
    Option.when(begin >= 0 && end > begin) {
      text.substring(0, begin + regionBegin.length) + "\n" + body + text.substring(end)
    }

/**
 * The generated reference (the five pages under `docs/reference/generated/`) and the function table
 * embedded in the xl-cli skill's `reference/FORMULAS.md` must equal what [[DocsGen]] renders from
 * `Schema.json`. A drift is a contract change: review it, then re-render deliberately with
 * `XL_UPDATE_DOCS=1` — the same discipline as the goldens.
 */
class DocsGenSpec extends FunSuite:

  private val update = sys.env.get("XL_UPDATE_DOCS").contains("1")
  private val generatedDir: Path = Golden.repoRoot.resolve("docs/reference/generated")
  private val formulasMd: Path =
    Golden.repoRoot.resolve("plugin/skills/xl-cli/reference/FORMULAS.md")

  private lazy val rendered: Map[String, String] = DocsGen.render(Schema.json(BuildInfo.version))

  private def read(path: Path): Option[String] =
    Option.when(Files.isRegularFile(path))(Files.readString(path, StandardCharsets.UTF_8))

  private def check(path: Path, expected: String): Unit =
    read(path) match
      case Some(actual) if actual == expected => ()
      case Some(actual) =>
        fail(
          s"${Golden.repoRoot.relativize(path)} differs from what Schema.json renders " +
            s"(re-render deliberately with ${DocsGen.regenerate}):\n" +
            UnifiedDiff.render(actual, expected)
        )
      case None => fail(s"${Golden.repoRoot.relativize(path)} is missing; ${DocsGen.regenerate}")

  DocsGen.files.foreach { name =>
    test(s"generated: $name") {
      val path = generatedDir.resolve(name)
      val expected = rendered.getOrElse(name, fail(s"no renderer for $name"))
      if update then
        Files.createDirectories(generatedDir)
        Files.writeString(path, expected, StandardCharsets.UTF_8)
      else check(path, expected)
    }
  }

  test("generated pages never carry the build version (a release bump must not rewrite them)") {
    rendered.foreach { (name, text) =>
      assert(!text.contains(BuildInfo.version), s"$name mentions ${BuildInfo.version}")
      assert(text.startsWith("<!-- GENERATED"), s"$name lacks the generated header")
    }
  }

  test("the skill's FORMULAS.md embeds the generated function table between its markers") {
    val body = DocsGen.functionTable(Schema.json(BuildInfo.version))
    val current = read(formulasMd).getOrElse(fail(s"$formulasMd is missing"))
    val expected = DocsGen
      .replaceRegion(current, body)
      .getOrElse(fail(s"$formulasMd lacks ${DocsGen.regionBegin} / ${DocsGen.regionEnd}"))
    if update then Files.writeString(formulasMd, expected, StandardCharsets.UTF_8)
    else if current != expected then
      fail(
        s"plugin/skills/xl-cli/reference/FORMULAS.md's generated block is stale " +
          s"(re-render deliberately with ${DocsGen.regenerate}):\n" +
          UnifiedDiff.render(current, expected)
      )
  }

  test("every registry function is named in FORMULAS.md (prose sections may add, never drop)") {
    val text = read(formulasMd).getOrElse(fail(s"$formulasMd is missing"))
    FunctionDoc.all.foreach { doc =>
      assert(text.contains(s"`${doc.name}`"), s"FORMULAS.md does not name ${doc.name}")
    }
  }
