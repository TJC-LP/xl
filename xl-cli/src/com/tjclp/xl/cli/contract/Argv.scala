package com.tjclp.xl.cli.contract

import scala.annotation.tailrec

/**
 * The command line before decline sees it (ADR-017 §2.2).
 *
 * decline parses the global options — `-f`, `-s`, `-o`, `--json`, … — in front of the verb only;
 * typed after it they were "Unexpected option". [[hoist]] moves every recognised global (with its
 * value) in front of the verb so both spellings are one command line, without reordering anything
 * else and without touching what follows a `--`. Two verbs own a flag that shares a global's name
 * ([[verbOwned]]): `view --strict` is view's own evaluation gate and `new --sheet`/`--backend` are
 * new's own options, so those are left where they are.
 *
 * Pure: `hoist(hoist(a)) == hoist(a)`, the result is a permutation of the input, and a command line
 * without globals comes back unchanged (ArgvSpec pins all three).
 */
object Argv:

  /** Global flag names and whether each consumes the token after it. */
  val globals: Map[String, Boolean] = Map(
    "-f" -> true,
    "--file" -> true,
    "-s" -> true,
    "--sheet" -> true,
    "-o" -> true,
    "--output" -> true,
    "-i" -> false,
    "--in-place" -> false,
    "--backend" -> true,
    "--max-size" -> true,
    "--stream" -> false,
    "--no-recalc" -> false,
    "--preserve-caches" -> false,
    "--json" -> false,
    "--strict" -> false
  )

  /**
   * Flags a verb owns under a global's name: for that verb they are the verb's own and are not
   * hoisted. `--strict` after `view` gates `--eval` (Main's `strictOpt`); `--sheet`/`--backend`
   * after `new` are its repeatable sheet list and writer backend.
   */
  val verbOwned: Map[String, Set[String]] = Map(
    "view" -> Set("--strict"),
    "new" -> Set("--sheet", "--backend")
  )

  /** Flags that never name a verb and take no value: decline's own. */
  private val meta: Set[String] = Set("--help", "--version", "-v")

  /**
   * Every top-level verb, in usage order — the names decline renders in `xl --help`, pinned against
   * that rendering by ArgvSpec so the two cannot drift.
   */
  val verbs: Vector[String] = Vector(
    "rasterizers",
    "functions",
    "new",
    "diff",
    "lint",
    "eval",
    "evala",
    "sheets",
    "names",
    "bounds",
    "view",
    "cell",
    "search",
    "stats",
    "filter",
    "describe",
    "audit",
    "deps",
    "batch",
    "put",
    "putf",
    "style",
    "row",
    "col",
    "group-rows",
    "group-cols",
    "ungroup-rows",
    "ungroup-cols",
    "autofit",
    "recalc",
    "import",
    "import-md",
    "add-sheet",
    "remove-sheet",
    "rename-sheet",
    "move-sheet",
    "copy-sheet",
    "merge",
    "unmerge",
    "comment",
    "remove-comment",
    "clear",
    "fill",
    "sort",
    "freeze",
    "unfreeze",
    "copy",
    "name",
    "insert-rows",
    "delete-rows",
    "insert-cols",
    "delete-cols",
    "chart",
    "add-image",
    "sheet-view",
    "tab-color",
    "autofilter",
    "page-setup",
    "header-footer",
    "cf"
  )

  /**
   * The global a token spells, as `(name, consumes the next token)`: the bare name (`-f`, `--file`)
   * or the `--name=value` form (which carries its value and consumes nothing).
   */
  private def globalOf(token: String): Option[(String, Boolean)] =
    globals.get(token).map(takesValue => (token, takesValue)).orElse {
      val eq = token.indexOf('=')
      Option
        .when(token.startsWith("--") && eq > 0)(token.substring(0, eq))
        .filter(name => globals.get(name).contains(true))
        .map(name => (name, false))
    }

  /**
   * Move every recognised global (and its value) in front of the first non-flag token — the verb —
   * keeping the globals' own order and every other token's relative order. `--` ends the scan: it
   * and everything after it are appended untouched. A global the verb owns ([[verbOwned]]) is not a
   * global on that command line.
   */
  def hoist(args: List[String]): List[String] =
    val (scanned, rest) = args.span(_ != "--")
    val owned = verbOf(scanned).flatMap(verbOwned.get).getOrElse(Set.empty)
    @tailrec
    def walk(
      tokens: List[String],
      hoisted: Vector[String],
      others: Vector[String]
    ): (Vector[String], Vector[String]) =
      tokens match
        case Nil => (hoisted, others)
        case token :: tail =>
          globalOf(token).filterNot((name, _) => owned.contains(name)) match
            case Some((_, true)) =>
              tail match
                case value :: after => walk(after, hoisted :+ token :+ value, others)
                case Nil => (hoisted :+ token, others)
            case Some((_, false)) => walk(tail, hoisted :+ token, others)
            case None => walk(tail, hoisted, others :+ token)
    val (hoisted, others) = walk(scanned, Vector.empty, Vector.empty)
    (hoisted ++ others).toList ++ rest

  /**
   * The verb the command line is heading for: the first token that is neither a global, a global's
   * value, nor `--help`/`--version`, before any `--`. It need not be a known verb — [[verbs]] tells
   * — and is `None` when no such token exists.
   */
  def verbOf(args: List[String]): Option[String] =
    @tailrec
    def find(tokens: List[String]): Option[String] = tokens match
      case Nil => None
      case token :: tail =>
        globalOf(token) match
          case Some((_, true)) => find(tail.drop(1))
          case Some((_, false)) => find(tail)
          case None if meta.contains(token) => find(tail)
          case None => Some(token)
    find(args.takeWhile(_ != "--"))
