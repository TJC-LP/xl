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

  /** The workbook extensions `xl` opens: a positional token ending in one was meant for `-f`. */
  private val workbookExtensions: Vector[String] =
    Vector(".xlsx", ".xlsm", ".xltx", ".xltm", ".xls")

  /** Verbs whose own positional IS a file (`new <output>`, `lint [file]`). */
  private val fileVerbs: Set[String] = Set("new", "lint")

  /** Verb options that take a file value: their token is not a misplaced `-f`. */
  private val fileOptions: Set[String] = Set("-g", "--file2")

  /**
   * GH-619: the workbook path a wrong command line passed positionally — the most frequent way an
   * agent opens (`xl view input.xlsx A1:B4`). Defined only when no `-f`/`--file` was given, the
   * verb does not take a file of its own, and a token that is not a flag, not the value of a global
   * or of a file option, and not behind `--` ends in a workbook extension. The usage error's hint
   * then says "did you mean -f <path>?" instead of pointing at `--help`.
   */
  def positionalFile(args: List[String]): Option[String] =
    val hasFile = args.exists(a => a == "-f" || a == "--file" || a.startsWith("--file="))
    if hasFile || verbOf(args).exists(fileVerbs.contains) then None
    else
      @tailrec
      def scan(rest: List[String]): Option[String] = rest match
        case Nil => None
        case "--" :: _ => None
        case flag :: _ :: tail
            if flag.startsWith("-") && (globals.getOrElse(flag, false) || fileOptions(flag)) =>
          scan(tail)
        case token :: tail =>
          if !token.startsWith("-") && looksLikeWorkbook(token) then Some(token) else scan(tail)
      scan(args)

  /**
   * The [[positionalFile]] the parser's errors corroborate: decline reported the missing `--file`
   * (or the workbook token sits where the verb's first positional goes, so it displaced the real
   * argument) or an unexpected argument — the token itself, or one that came after it. A failure
   * about something else (`Invalid integer: x`) gets the plain `--help` hint.
   */
  def misplacedFile(args: List[String], errors: List[String]): Option[String] =
    positionalFile(args).filter { token =>
      val scanned = args.takeWhile(_ != "--")
      val tokenAt = scanned.indexOf(token)
      val firstPositional = verbOf(scanned).flatMap { verb =>
        scanned.drop(scanned.indexOf(verb) + 1).find(t => !t.startsWith("-"))
      }
      val missingFile = errors.exists(_.startsWith("Missing expected flag --file"))
      val unexpected = errors.filter(_.startsWith("Unexpected argument: "))
      val namesToken = unexpected.contains(s"Unexpected argument: $token")
      val displaced = unexpected.exists { e =>
        scanned.indexOf(e.stripPrefix("Unexpected argument: ")) > tokenAt
      }
      namesToken || (firstPositional.contains(token) && (missingFile || displaced))
    }

  /** The output globals a verb that never writes refuses, each with the pair the error names. */
  private val outputGlobals: Map[String, String] = Map(
    "-o" -> "-o/--output",
    "--output" -> "-o/--output",
    "-i" -> "-i/--in-place",
    "--in-place" -> "-i/--in-place"
  )

  /**
   * GH-667: the known verb that never writes, given an output global — as `(verb, flag pair)`. A
   * verb never writes when every form of it in [[Schema.verbs]] has `needs.output == false` and it
   * is not `new`, whose output file is its positional. decline cannot say so itself: once `-o` is
   * consumed only the parsers that take it survive, none of which knows a read-only verb, so its
   * first error is "Unexpected argument: view" — the verb, not the flag. The flag is read as
   * decline reads it ([[wantsJson]]'s rule): a token that is the VALUE of another global is data,
   * and nothing behind `--` counts. `None` for a write verb, a verb with a writing form (`sheets
   * hide`), an unknown verb (UNKNOWN_VERB's business) or a command line without an output flag.
   */
  def outputOnReadOnlyVerb(args: List[String]): Option[(String, String)] =
    def neverWrites(verb: String): Boolean =
      val forms = Schema.verbs.filter(_.path.headOption.contains(verb))
      verb != "new" && forms.nonEmpty && forms.forall(!_.needs.output)
    for
      verb <- verbOf(args).filter(neverWrites)
      (flag, _) <- globalAmong(args, outputGlobals.keySet)
    yield (verb, outputGlobals.getOrElse(flag, flag))

  /**
   * GH-676: the verbs no form of which takes the `-s/--sheet` global, each with the reason the
   * usage error's hint gives. decline cannot name the flag here either: once `-s` is consumed only
   * the parsers that take it survive, so its first error is "Unexpected argument: names". The one
   * table the runner reads; ArgvSpec pins it against the parser (exactly the verbs to which adding
   * `-s` makes decline refuse the verb).
   */
  val sheetlessVerbs: Map[String, String] = Map(
    "names" -> "names lists every defined name in the workbook, each with its scope",
    "sheets" ->
      ("sheets works on the whole workbook; hide and show take the sheet as their argument " +
        "(`sheets hide <name>`)"),
    "lint" -> "lint validates the whole package",
    "new" ->
      ("new names the sheets it creates with --sheet after the verb " +
        "(`xl new out.xlsx --sheet Data`)"),
    "rasterizers" -> "rasterizers reads no workbook",
    "functions" -> "functions reads no workbook",
    "schema" -> "schema reads no workbook"
  )

  /** The spellings of the sheet global. */
  private val sheetGlobals: Set[String] = Set("-s", "--sheet")

  /**
   * GH-676: the [[sheetlessVerbs]] verb given the `-s/--sheet` global, read as decline reads it (a
   * global's value is data, nothing behind `--` counts, and `--sheet` after `new` is new's own
   * option — before the verb it is the global in either spelling).
   */
  def sheetOnSheetlessVerb(args: List[String]): Option[String] =
    for
      verb <- verbOf(args).filter(sheetlessVerbs.contains)
      _ <- aroundVerb(args, sheetGlobals)(globalAmong(_, _), _.orElse(_))
    yield verb

  /**
   * GH-676: `new` given an output global, as `(flag pair, its value)`: `new` names its output file
   * positionally (`xl new out.xlsx`), so decline's first error was "Unexpected argument: new".
   */
  def outputOnNew(args: List[String]): Option[(String, Option[String])] =
    verbOf(args).filter(_ == "new").flatMap { _ =>
      globalAmong(args, outputGlobals.keySet).map { (flag, value) =>
        (outputGlobals.getOrElse(flag, flag), value)
      }
    }

  /**
   * A global the verb cannot take (GH-667, GH-676): the usage error to report, and the command line
   * with that mistake repaired — the flag dropped. `new -o out.xlsx` has a second repair, its value
   * moved to the positional, for the line that names no output otherwise. The runner re-parses the
   * repairs and reports the misuse when either parse is clean; otherwise the misuse was not the
   * only mistake, and decline's error for `repaired` — the one naming the other mistake — is
   * reported instead (GH-681).
   */
  final case class Misuse(
    message: String,
    hint: String,
    repaired: List[String],
    fallback: Option[List[String]]
  ) derives CanEqual

  /**
   * Every misuse on the command line, in the order the runner reports them, each repairing the line
   * the previous one left: an output global on a verb that never writes, the sheet global on a
   * [[sheetlessVerbs]] verb, an output global on `new`. The last one's `repaired` is the line with
   * every misuse removed.
   */
  def misuses(args: List[String]): Vector[Misuse] =
    val readOnlyOutput = (line: List[String]) =>
      outputOnReadOnlyVerb(line).map { (verb, flag) =>
        // view's raster and document exports DO write a file, through --raster-output — point at
        // it rather than saying the verb never writes (PR #679 review)
        val writes =
          if verb == "view" then
            s"$verb writes no workbook; png/jpeg/webp/pdf/svg/html exports take " +
              "--raster-output <path>"
          else s"$verb writes no file"
        Misuse(
          s"$verb is read-only and does not take $flag",
          s"drop $flag; $writes. Run `xl $verb --help` for the usage",
          dropGlobals(line, outputGlobals.keySet),
          None
        )
      }
    val sheetless = (line: List[String]) =>
      sheetOnSheetlessVerb(line).map { verb =>
        Misuse(
          s"$verb does not take -s/--sheet",
          s"drop -s/--sheet; ${sheetlessVerbs.getOrElse(verb, s"$verb takes no sheet")}. " +
            s"Run `xl $verb --help` for the usage",
          aroundVerb(line, sheetGlobals)(dropGlobals(_, _), _ ++ _),
          None
        )
      }
    val newOutput = (line: List[String]) =>
      outputOnNew(line).map { (flag, value) =>
        val dropped = dropGlobals(line, outputGlobals.keySet)
        // a line that already names its output positionally is repaired by the drop alone;
        // inserting the value too would make two positionals and blame the user's own
        val positional = for
          file <- value
          at <- verbIndex(dropped)
        yield
          val (upTo, after) = dropped.splitAt(at + 1)
          upTo ++ (file :: after)
        Misuse(
          s"new does not take $flag",
          s"new takes its output as a positional: `xl new ${value.getOrElse("<file>")}`",
          dropped,
          positional
        )
      }
    Vector(readOnlyOutput, sheetless, newOutput)
      .foldLeft((args, Vector.empty[Misuse])) { case ((line, found), rule) =>
        rule(line).fold((line, found))(misuse => (misuse.repaired, found :+ misuse))
      }
      ._2

  /**
   * The first of the named globals on the command line with its value, read as decline reads it
   * (the rule of [[wantsJson]]): a token that is the VALUE of another global is data, and nothing
   * behind `--` counts. The value is `None` for a flag that takes none or a dangling one.
   */
  private def globalAmong(
    args: List[String],
    names: Set[String]
  ): Option[(String, Option[String])] =
    @tailrec
    def scan(tokens: List[String]): Option[(String, Option[String])] = tokens match
      case Nil => None
      case token :: tail =>
        globalOf(token) match
          case Some((name, takesValue)) if names.contains(name) =>
            val value =
              if takesValue then tail.headOption
              else Option.when(token.contains('='))(token.substring(token.indexOf('=') + 1))
            Some((name, value))
          case Some((_, takesValue)) => scan(if takesValue then tail.drop(1) else tail)
          case None => scan(tail)
    scan(args.takeWhile(_ != "--"))

  /**
   * `f` over the command line with the named globals, split at the verb: in front of it every
   * spelling is the global, behind it the ones the verb owns ([[verbOwned]]) are the verb's own
   * option (`new out.xlsx --sheet Data`), so `f` sees only the rest there. `join` combines the two
   * halves' answers; a line without a verb is all "in front".
   */
  private def aroundVerb[A](args: List[String], names: Set[String])(
    f: (List[String], Set[String]) => A,
    join: (A, A) => A
  ): A =
    verbIndex(args) match
      case Some(at) =>
        val owned = args.lift(at).flatMap(verbOwned.get).getOrElse(Set.empty)
        val (front, behind) = args.splitAt(at)
        join(f(front, names), f(behind, names -- owned))
      case None => f(args, names)

  /** The command line without the named globals (and their values) before any `--`. */
  private def dropGlobals(args: List[String], names: Set[String]): List[String] =
    val (scanned, rest) = args.span(_ != "--")
    @tailrec
    def strip(tokens: List[String], kept: Vector[String]): Vector[String] = tokens match
      case Nil => kept
      case token :: tail =>
        globalOf(token) match
          case Some((name, takesValue)) if names.contains(name) =>
            strip(if takesValue then tail.drop(1) else tail, kept)
          case Some((_, true)) => strip(tail.drop(1), kept ++ (token :: tail.take(1)))
          case _ => strip(tail, kept :+ token)
    strip(scanned, Vector.empty).toList ++ rest

  /**
   * The command line without its globals (and their values), before any `--`: what decline should
   * see when help is asked for, since the globals a program adds (`--json`, `-f`) are already read
   * here and decline refuses `--help` behind an option it did not expect (GH-620).
   */
  def withoutGlobals(args: List[String]): List[String] =
    val (scanned, rest) = args.span(_ != "--")
    @tailrec
    def strip(tokens: List[String], kept: Vector[String]): Vector[String] = tokens match
      case Nil => kept
      case token :: tail =>
        globalOf(token) match
          case Some((_, true)) => strip(tail.drop(1), kept)
          case Some((_, false)) => strip(tail, kept)
          case None => strip(tail, kept :+ token)
    strip(scanned, Vector.empty).toList ++ rest

  /**
   * The command lines to try, in order, when `--help` was asked for (GH-620): the line without its
   * globals, then the verb path shortened from the right one word at a time — `view A1:B2 --help`,
   * `view --help`, `--help` — because decline honours `--help` only before a positional it did not
   * expect. The runner takes the first that parses to a help text without errors.
   */
  def helpCandidates(args: List[String]): Vector[List[String]] =
    val stripped = withoutGlobals(args).takeWhile(_ != "--")
    val words = stripped.filterNot(_.startsWith("-"))
    val prefixes = (words.length to 0 by -1).map(n => words.take(n) :+ "--help")
    (stripped +: prefixes.toVector).distinct

  private def looksLikeWorkbook(token: String): Boolean =
    val lower = token.toLowerCase(java.util.Locale.ROOT)
    workbookExtensions.exists(ext => lower.endsWith(ext) && lower.length > ext.length)

  /**
   * Flags a verb owns under a global's name: for that verb they are the verb's own and are not
   * hoisted. `--strict` after `view` gates `--eval` (Main's `strictOpt`); `--sheet`/`--backend`
   * after `new` are its repeatable sheet list and writer backend.
   */
  val verbOwned: Map[String, Set[String]] = Map(
    "view" -> Set("--strict"),
    "new" -> Set("--sheet", "--backend")
  )

  /**
   * Every top-level verb, in usage order — the names decline renders in `xl --help`, pinned against
   * that rendering by ArgvSpec so the two cannot drift.
   */
  val verbs: Vector[String] = Vector(
    "rasterizers",
    "functions",
    "schema",
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
                // a dangling value-taking global stays put, so decline reports it as it is
                case Nil => (hoisted, others :+ token)
            case Some((_, false)) => walk(tail, hoisted :+ token, others)
            case None => walk(tail, hoisted, others :+ token)
    val (hoisted, others) = walk(scanned, Vector.empty, Vector.empty)
    (hoisted ++ others).toList ++ rest

  /**
   * Whether `--json` is among the flags before any `--`. A token that is the VALUE of a
   * value-taking global is data, not a flag — `-o --json` names an output file called `--json` (the
   * same rule [[hoist]] applies) — so this reads what decline will read: a usage failure and a
   * successful parse of the same command line always agree on the output mode.
   */
  def wantsJson(args: List[String]): Boolean =
    @tailrec
    def scan(tokens: List[String]): Boolean = tokens match
      case Nil => false
      case "--json" :: _ => true
      case token :: tail =>
        globalOf(token) match
          case Some((_, true)) => scan(tail.drop(1))
          case _ => scan(tail)
    scan(args.takeWhile(_ != "--"))

  /** Whether `--stream` is among the flags before any `--` (the rule of [[wantsJson]]). */
  def wantsStream(args: List[String]): Boolean =
    @tailrec
    def scan(tokens: List[String]): Boolean = tokens match
      case Nil => false
      case "--stream" :: _ => true
      case token :: tail =>
        globalOf(token) match
          case Some((_, true)) => scan(tail.drop(1))
          case _ => scan(tail)
    scan(args.takeWhile(_ != "--"))

  /**
   * The verb the command line is heading for: the first token that is neither a global, a global's
   * value, nor a flag of any kind (`--help`, `--version`, an unknown `--frobnicate` — decline
   * reports those), before any `--`. It need not be a known verb — [[verbs]] tells — and is `None`
   * when no such token exists.
   */
  def verbOf(args: List[String]): Option[String] =
    verbIndex(args).flatMap(args.lift)

  /** The position of [[verbOf]]'s token in `args`. */
  private def verbIndex(args: List[String]): Option[Int] =
    @tailrec
    def find(tokens: List[String], at: Int): Option[Int] = tokens match
      case Nil => None
      case token :: tail =>
        globalOf(token) match
          case Some((_, true)) => find(tail.drop(1), at + 1 + tail.take(1).length)
          case Some((_, false)) => find(tail, at + 1)
          case None if token.startsWith("-") => find(tail, at + 1)
          case None => Some(at)
    find(args.takeWhile(_ != "--"), 0)
