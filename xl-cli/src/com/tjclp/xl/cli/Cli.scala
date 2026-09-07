package com.tjclp.xl.cli

import cats.effect.{ExitCode, IO}
import cats.implicits.*
import com.monovore.decline.{Command, Opts, Visibility}

import com.tjclp.xl.cli.Main.*
import com.tjclp.xl.cli.contract.{CliError, ErrorCode, ExitCodes, Outcome, OutputMode}

/**
 * The `xl` command line as a value.
 *
 * [[program]] is the complete verb tree with every handler wired to a [[CliIO]]; [[run]] is what
 * the binary does with argv: parse, then run the selected handler or render help. Keeping both
 * behind an injectable `CliIO` lets the contract suite (`CliHarness` under xl-cli/test) drive the
 * real parser and the real handlers in-process and pin exit code, stdout and stderr byte for byte.
 *
 * Help and version behaviour reproduce decline-effect's `CommandIOApp`, which this replaced because
 * its `run` is final and prints through an ambient console:
 *   - `--help` renders the help on STDERR and exits 0
 *   - a parse failure renders the errors plus help on stderr and exits 2 (usage — ADR-017 §2.3; the
 *     one-line usage instead of the full help arrives with the argv cluster). With `--json` among
 *     the arguments it renders the `ok:false` envelope instead (§2.4), so a program never has to
 *     parse help text out of a failed call
 *   - `--version` / `-v` prints the version on stdout and exits 0
 *
 * Help landing on stderr even when asked for is the current, pinned contract (golden `help`).
 * Command errors go to stderr as `Error: <message>` plus a `code:` line ([[contract.Diagnostics]]);
 * stdout is empty on every failure. Under the global `--json` every result — success or failure —
 * is one envelope on stdout ([[contract.Render.json]]).
 */
object Cli:

  val name: String = "xl"

  /** The `--help` header: what the tool is, then the exit-code table every pipeline branches on. */
  val header: String =
    """LLM-friendly Excel operations (stateless)
      |
      |Exit codes:
      |  0  ok
      |  1  completed with findings or a failed gate (diff differs, lint findings, --strict) — never a failure
      |  2  usage — the command line is wrong; nothing read, nothing written
      |  3  failed — the operation could not complete; nothing written
      |Results go to stdout; errors (Error: <message>, then code:/hint: lines) and warnings go to stderr.
      |--json wraps every result, success or failure, in one JSON envelope on stdout:
      |  {ok, exitCode, verb, version, data, warnings, error}""".stripMargin

  /**
   * Every top-level verb, in usage order — the best-effort `verb` of an envelope for a usage error
   * raised before dispatch ([[verbOf]]). Pinned against the parser's own list by the contract
   * suite.
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

  /** The verb tree. Options and handlers live in [[Main]]; this is the wiring between them. */
  def program(io: CliIO): Opts[IO[ExitCode]] =
    // Workbook-level: only --file (no --sheet)
    // Note: --stream not supported for workbook-level commands (need full metadata)
    val workbookSubcmds = namesCmd
    val workbookOpts = (fileOpt, maxSizeOpt, jsonOpt, workbookSubcmds).mapN {
      (file, maxSize, mode, cmd) =>
        Main.run(file, None, None, None, None, maxSize, false, cmd, io, mode)
    }

    // Sheets command: --file required, --output optional (required for hide/show, not for list)
    // This needs its own opts chain because list doesn't need output but hide/show do
    val sheetsOpts =
      (
        fileOpt,
        outputOpt.orNone,
        inPlaceOpt,
        backendOpt,
        maxSizeOpt,
        streamOpt,
        jsonOpt,
        sheetsCmd
      ).mapN { (file, outOpt, inPlace, backend, maxSize, stream, mode, cmd) =>
        runWithOutput(outOpt, inPlace, file, io, mode, cmd.verb) { (out, display) =>
          runResult(
            file,
            None,
            out,
            display,
            backend,
            maxSize,
            stream,
            cmd,
            strictFailureDiscardsOutput = inPlace,
            io = io,
            mode = mode
          )
        }
      }

    // Headless commands: --file is optional (for constant formulas like =1+1, =PI())
    // Note: --stream not supported for eval (needs formula analysis)
    // evala requires --file (array formulas need sheet context)
    val headlessOpts =
      (fileOpt.orNone, sheetOpt, maxSizeOpt, jsonOpt, evalCmd orElse evalArrayCmd).mapN {
        (fileOpt, sheet, maxSize, mode, cmd) =>
          runHeadless(fileOpt, sheet, maxSize, cmd, io, mode)
      }

    // Sheet-level read-only: --file and --sheet (no --output)
    val sheetReadOnlySubcmds =
      boundsCmd orElse viewCmd orElse cellCmd orElse searchCmd orElse statsCmd orElse filterCmd orElse describeCmd orElse auditCmd orElse depsCmd

    val sheetReadOnlyOpts =
      (fileOpt, sheetOpt, maxSizeOpt, streamOpt, jsonOpt, sheetReadOnlySubcmds).mapN {
        (file, sheet, maxSize, stream, mode, cmd) =>
          Main.run(file, sheet, None, None, None, maxSize, stream, cmd, io, mode)
      }

    // Sheet-level write: --file, --sheet, and --output (required)
    // --stream uses SAX/StAX workbook writes for modifying commands.
    val sheetWriteSubcmds =
      putCmd orElse putfCmd orElse styleCmd orElse rowCmd orElse colCmd orElse groupRowsCmd orElse groupColsCmd orElse ungroupRowsCmd orElse ungroupColsCmd orElse autoFitCmd orElse batchCmd orElse recalcCmd orElse importCmd orElse importMdCmd orElse addSheetCmd orElse removeSheetCmd orElse renameSheetCmd orElse moveSheetCmd orElse copySheetCmd orElse mergeCmd orElse unmergeCmd orElse commentCmd orElse removeCommentCmd orElse clearCmd orElse fillCmd orElse sortCmd orElse freezeCmd orElse unfreezeCmd orElse copyCmd orElse nameCmd orElse insertRowsCmd orElse deleteRowsCmd orElse insertColsCmd orElse deleteColsCmd orElse chartCmd orElse addImageCmd orElse sheetViewCmd orElse tabColorCmd orElse autoFilterCmd orElse pageSetupCmd orElse headerFooterCmd orElse cfCmd

    val sheetWriteOpts =
      (
        fileOpt,
        sheetOpt,
        outputOpt.orNone,
        inPlaceOpt,
        backendOpt,
        maxSizeOpt,
        streamOpt,
        writePolicyOpt,
        jsonOpt,
        sheetWriteSubcmds
      ).mapN { (file, sheet, outOpt, inPlace, backend, maxSize, stream, policy, mode, cmd) =>
        runWithOutput(outOpt, inPlace, file, io, mode, cmd.verb) { (out, display) =>
          runResult(
            file,
            sheet,
            out,
            display,
            backend,
            maxSize,
            stream,
            cmd,
            policy,
            strictFailureDiscardsOutput = inPlace,
            io = io,
            mode = mode
          )
        }
      }

    // Standalone: no --file required (creates new files)
    val standaloneOpts = (jsonOpt, newCmd).mapN {
      case (mode, (outPath, sheetName, sheets, backend)) =>
        runStandalone(outPath, sheetName, sheets, backend, io, mode)
    }

    // Diff: compares -f against -g (two inputs, no output); custom exit codes
    val diffOpts = (fileOpt, sheetOpt, maxSizeOpt, jsonOpt, diffCmd).mapN {
      (file, sheet, maxSize, mode, cmd) =>
        cmd match
          case CliCommand.Diff(file2, format) =>
            runDiff(file, file2, sheet, maxSize, format, io, mode)
          case other => internal("diff", s"Unexpected diff command: $other", io, mode)
    }

    // Lint: raw-zip structural validation (GH-397, no output file); custom exit codes.
    // The file arrives via -f or positionally (GH-422); exactly one form must be used.
    val lintOpts = (fileOpt.orNone, jsonOpt, lintCmd).mapN {
      case (flagFile, mode, (cmd, positional)) =>
        cmd match
          case CliCommand.Lint(format) =>
            resolveLintFile(flagFile, positional) match
              case Right(file) => runLint(file, format, io, mode)
              case Left(msg) => emit(Outcome.failed("lint", CliError.usage(msg, None)), mode, io)
          case other => internal("lint", s"Unexpected lint command: $other", io, mode)
    }

    // Info commands: no file required
    val infoOpts = (jsonOpt, functionsCmd).mapN((mode, _) => runInfo(io, mode))
    val rasterOpts = (jsonOpt, rasterizersCmd).mapN((mode, _) => runRasterizers(io, mode))

    // Batch dry-run: only needs batch source, no --file or --output. Same shape as the other
    // standalone runners: a bad source (invalid JSON, missing file) is a diagnostic on stderr with
    // the code's exit, never an escaped exception (which IOApp would print as a trace with exit 1).
    val dryRunFlag =
      Opts.flag("dry-run", "Validate batch JSON without writing")
    val batchDryRunOpts =
      (
        jsonOpt,
        Opts.subcommand("batch", batchHelp) {
          (batchArg, dryRunFlag).mapN((src, _) => src)
        }
      ).mapN { (mode, src) =>
        batchDryRunOutcome(src, io, mode).flatMap(emit(_, mode, io))
      }

    rasterOpts orElse infoOpts orElse standaloneOpts orElse diffOpts orElse lintOpts orElse headlessOpts orElse sheetsOpts orElse workbookOpts orElse sheetReadOnlyOpts orElse batchDryRunOpts orElse sheetWriteOpts

  /** The parser the binary runs: the program plus `--version`, under decline's `--help`. */
  def command(io: CliIO): Command[IO[ExitCode]] =
    Command(name, header, helpFlag = true)(versionFlag(io) orElse program(io))

  /**
   * argv to exit code: parse, then run the handler or render help (see the object doc). A parse
   * error with `--json` among the raw arguments is the `USAGE` envelope (exit 2) rather than help
   * text: the parser never reached a handler that could have seen the flag, so the flag is read
   * here.
   */
  def run(args: List[String], io: CliIO): IO[ExitCode] =
    IO(command(io).parse(args, sys.env)).flatMap {
      case Right(handler) => handler
      case Left(help) if help.errors.nonEmpty && wantsJson(args) =>
        val error = CliError.usage(
          help.errors.mkString("; "),
          Some("run `xl --help` (or `xl <verb> --help`) for the usage")
        )
        emit(Outcome.failed(verbOf(args), error), OutputMode.Json, io)
      case Left(help) =>
        io.err(help.toString)
          .as(if help.errors.nonEmpty then ExitCodes.usage else ExitCodes.ok)
    }

  /** `--json` among the arguments before any `--`: after it every token is data, not a flag. */
  private def wantsJson(args: List[String]): Boolean =
    args.takeWhile(_ != "--").contains("--json")

  /** The global options that take a value: the token after them is never the verb. */
  private val valueOptions: Set[String] =
    Set(
      "-f",
      "--file",
      "-s",
      "--sheet",
      "-o",
      "--output",
      "-g",
      "--file2",
      "--max-size",
      "--backend"
    )

  /**
   * The first argument that names a verb — skipping the values of global options and stopping at
   * `--` — else empty: what a failed parse was heading for. `--option=value` is one token and needs
   * no skip.
   */
  def verbOf(args: List[String]): String =
    @annotation.tailrec
    def find(rest: List[String]): String = rest match
      case Nil => ""
      case option :: _ :: tail if valueOptions.contains(option) => find(tail)
      case token :: tail => if verbs.contains(token) then token else find(tail)
    find(args.takeWhile(_ != "--"))

  /** A dispatch arm the parser cannot reach: a defect, reported like any other failure (exit 3). */
  private def internal(verb: String, message: String, io: CliIO, mode: OutputMode): IO[ExitCode] =
    emit(Outcome.failed(verb, CliError(ErrorCode.INTERNAL, message)), mode, io)

  /**
   * Exactly the flag decline-effect's `CommandIOApp` added: `--version`/`-v`, partially visible.
   */
  private def versionFlag(io: CliIO): Opts[IO[ExitCode]] =
    Opts
      .flag(
        long = "version",
        short = "v",
        help = "Print the version number and exit.",
        visibility = Visibility.Partial
      )
      .as(io.out(BuildInfo.version).as(ExitCode.Success))
