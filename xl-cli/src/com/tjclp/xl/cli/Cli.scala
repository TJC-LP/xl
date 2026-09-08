package com.tjclp.xl.cli

import cats.effect.{ExitCode, IO}
import cats.implicits.*
import com.monovore.decline.{Command, Opts, Visibility}

import com.tjclp.xl.cli.Main.*
import com.tjclp.xl.cli.contract.{
  Argv,
  CliError,
  Diagnostics,
  ErrorCode,
  ExitCodes,
  Outcome,
  OutputMode
}
import com.tjclp.xl.text.Suggest

/**
 * The `xl` command line as a value.
 *
 * [[program]] is the complete verb tree with every handler wired to a [[CliIO]]; [[run]] is what
 * the binary does with argv: hoist the globals ([[contract.Argv]]), parse, then run the selected
 * handler or render help. Keeping both behind an injectable `CliIO` lets the contract suite
 * (`CliHarness` under xl-cli/test) drive the real parser and the real handlers in-process and pin
 * exit code, stdout and stderr byte for byte.
 *
 * Help and version behaviour reproduce decline-effect's `CommandIOApp`, which this replaced because
 * its `run` is final and prints through an ambient console:
 *   - `--help` renders the help on STDERR and exits 0
 *   - an unknown verb is `UNKNOWN_VERB` (exit 2) with the nearest verb names as "did you mean"
 *   - any other parse failure is `USAGE` (exit 2): the first parser error with the verb's `--help`
 *     as the hint, then the one-line [[usage]] — never the full subcommand dump (ADR-017 §2.2).
 *     With `--json` among the arguments both render the `ok:false` envelope instead (§2.4), so a
 *     program never has to parse help text out of a failed call
 *   - `--version` / `-v` prints the version on stdout and exits 0
 *
 * Help landing on stderr even when asked for is the current, pinned contract (golden `help`).
 * Command errors go to stderr as `Error: <message>` plus a `code:` line ([[contract.Diagnostics]]);
 * stdout is empty on every failure. Under the global `--json` every result — success or failure —
 * is one envelope on stdout ([[contract.Render.json]]). Nothing escapes as a stack trace: a failure
 * no handler classified is rendered by [[run]]'s last-resort handler as `INTERNAL` (exit 3), in the
 * mode the command line asked for.
 */
object Cli:

  val name: String = "xl"

  /** The one-line usage a wrong command line gets; the verbs are listed by `xl --help`. */
  val usage: String = "usage: xl [-f FILE] [-s SHEET] [-o OUT | -i] [--json] <verb> …"

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
            runDiff(file, file2, sheet, maxSize, CliCommand.diffFormat(format, mode), io, mode)
          case other => internal("diff", s"Unexpected diff command: $other", io, mode)
    }

    // Lint: raw-zip structural validation (GH-397, no output file); custom exit codes.
    // The file arrives via -f or positionally (GH-422); exactly one form must be used.
    val lintOpts = (fileOpt.orNone, jsonOpt, lintCmd).mapN {
      case (flagFile, mode, (cmd, positional)) =>
        cmd match
          case CliCommand.Lint(format) =>
            resolveLintFile(flagFile, positional) match
              case Right(file) => runLint(file, CliCommand.lintFormat(format, mode), io, mode)
              case Left(msg) => emit(Outcome.failed("lint", CliError.usage(msg, None)), mode, io)
          case other => internal("lint", s"Unexpected lint command: $other", io, mode)
    }

    // Info commands: no file required
    val infoOpts = (jsonOpt, functionsCmd).mapN((mode, _) => runInfo(io, mode))
    val rasterOpts = (jsonOpt, rasterizersCmd).mapN((mode, _) => runRasterizers(io, mode))

    // Standalone batch forms — `--dry-run <source>` validates, `--schema` prints the document's
    // JSON Schema — need no --file or --output. Same shape as the other standalone runners: a bad
    // source (invalid JSON, missing file) is a diagnostic on stderr with the code's exit, never an
    // escaped exception (which IOApp would print as a trace with exit 1).
    val batchDryRunOpts =
      (jsonOpt, Opts.subcommand("batch", batchHelp)(batchStandaloneArgs)).mapN { (mode, form) =>
        batchStandaloneOutcome(form, io, mode).flatMap(emit(_, mode, io))
      }

    // The contract itself (ADR-017 §2.13): `xl schema [--json]`, no file required
    val contractOpts = schemaOpts(io)

    rasterOpts orElse infoOpts orElse contractOpts orElse standaloneOpts orElse diffOpts orElse lintOpts orElse headlessOpts orElse sheetsOpts orElse workbookOpts orElse sheetReadOnlyOpts orElse batchDryRunOpts orElse sheetWriteOpts

  /** The parser the binary runs: the program plus `--version`, under decline's `--help`. */
  def command(io: CliIO): Command[IO[ExitCode]] =
    Command(name, header, helpFlag = true)(versionFlag(io) orElse program(io))

  /**
   * argv to exit code (see the object doc): hoist the globals in front of the verb, refuse an
   * unknown verb with a suggestion, then parse and run the handler or render help. A usage failure
   * with `--json` among the arguments is the envelope (exit 2) rather than text: the parser never
   * reached a handler that could have seen the flag, so the flag is read here
   * ([[contract.Argv.wantsJson]] — a `--json` that is the value of `-o` is a file name, not the
   * flag, exactly as decline will read it).
   *
   * The last-resort handler is the contract's floor: whatever a handler lets escape (the staging
   * and commit steps run outside the handlers' own `attempt`) is still one diagnostic — or one
   * envelope — with the code [[contract.CliError.fromThrowable]] assigns, never a stack trace. Its
   * floor has a floor: cats-effect never delivers a fatal error (an `OutOfMemoryError` above all)
   * to this handler — it halts the runtime — so the in-memory load catches that one inside its own
   * thunk and re-raises it typed ([[MemoryGuard]], GH-636); every other `Error` stays fatal.
   */
  def run(args: List[String], io: CliIO): IO[ExitCode] =
    val argv = Argv.hoist(args)
    val mode = if Argv.wantsJson(argv) then OutputMode.Json else OutputMode.Text
    Argv.verbOf(argv) match
      case Some(word) if !Argv.verbs.contains(word) =>
        val error = CliError(
          ErrorCode.UNKNOWN_VERB,
          s"unknown verb '$word'",
          hint = Some("run `xl --help` for the list of verbs"),
          candidates = Suggest.closest(word, Argv.verbs)
        )
        usageFailure("", error, mode, io)
      case verb =>
        IO(command(io).parse(argv, sys.env))
          .flatMap {
            case Right(handler) => handler
            case Left(help) if help.errors.nonEmpty =>
              val error = CliError.usage(
                help.errors.headOption.fold("invalid command line") { first =>
                  if verb.isEmpty then compact(first) else first
                },
                Some(s"run `xl ${verb.fold("")(_ + " ")}--help` for the usage")
              )
              usageFailure(verb.getOrElse(""), error, mode, io)
            case Left(help) => io.err(help.toString).as(ExitCodes.ok)
          }
          .handleErrorWith { escaped =>
            emit(Outcome.failed(verb.getOrElse(""), CliError.fromThrowable(escaped)), mode, io)
          }

  /**
   * decline's first error, minus the one dump it embeds: with no verb at all it lists every
   * subcommand inside "Missing expected command (a or b or …)!", which is the help, not an error.
   * Applied only when no verb was given: a verb missing its sub-verb (`name`, `cf`) keeps decline's
   * own short list (`Missing expected command (add or rm)!`).
   */
  private def compact(error: String): String =
    if error.startsWith("Missing expected command") then "Missing expected command: no verb given"
    else error

  /**
   * A wrong command line (exit 2): under `--json` the `ok:false` envelope; in text mode the
   * diagnostics block — `Error:` first on stderr, as on every failure — then the one-line
   * [[usage]], nothing on stdout.
   */
  private def usageFailure(
    verb: String,
    error: CliError,
    mode: OutputMode,
    io: CliIO
  ): IO[ExitCode] =
    mode match
      case OutputMode.Json => emit(Outcome.failed(verb, error), mode, io)
      case OutputMode.Text =>
        io.err(s"${Diagnostics.render(error)}\n$usage").as(error.exitCode)

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
