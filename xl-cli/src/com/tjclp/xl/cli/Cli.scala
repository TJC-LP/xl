package com.tjclp.xl.cli

import cats.effect.{ExitCode, IO}
import cats.implicits.*
import com.monovore.decline.{Command, Opts, Visibility}

import com.tjclp.xl.cli.Main.*
import com.tjclp.xl.cli.contract.{CliError, Diagnostics, ErrorCode, ExitCodes}

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
 *     one-line usage instead of the full help arrives with the argv cluster)
 *   - `--version` / `-v` prints the version on stdout and exits 0
 *
 * Help landing on stderr even when asked for is the current, pinned contract (golden `help`).
 * Command errors go to stderr as `Error: <message>` plus a `code:` line ([[contract.Diagnostics]]);
 * stdout is empty on every failure.
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
      |Results go to stdout; errors (Error: <message>, then code:/hint: lines) and warnings go to stderr.""".stripMargin

  /** The verb tree. Options and handlers live in [[Main]]; this is the wiring between them. */
  def program(io: CliIO): Opts[IO[ExitCode]] =
    // Workbook-level: only --file (no --sheet)
    // Note: --stream not supported for workbook-level commands (need full metadata)
    val workbookSubcmds = namesCmd
    val workbookOpts = (fileOpt, maxSizeOpt, workbookSubcmds).mapN { (file, maxSize, cmd) =>
      Main.run(file, None, None, None, None, maxSize, false, cmd, io)
    }

    // Sheets command: --file required, --output optional (required for hide/show, not for list)
    // This needs its own opts chain because list doesn't need output but hide/show do
    val sheetsOpts =
      (fileOpt, outputOpt.orNone, inPlaceOpt, backendOpt, maxSizeOpt, streamOpt, sheetsCmd).mapN {
        (file, outOpt, inPlace, backend, maxSize, stream, cmd) =>
          runWithOutput(outOpt, inPlace, file, io) { (out, display) =>
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
              io = io
            )
          }
      }

    // Headless commands: --file is optional (for constant formulas like =1+1, =PI())
    // Note: --stream not supported for eval (needs formula analysis)
    // evala requires --file (array formulas need sheet context)
    val headlessOpts = (fileOpt.orNone, sheetOpt, maxSizeOpt, evalCmd orElse evalArrayCmd).mapN {
      (fileOpt, sheet, maxSize, cmd) =>
        runHeadless(fileOpt, sheet, maxSize, cmd, io)
    }

    // Sheet-level read-only: --file and --sheet (no --output)
    val sheetReadOnlySubcmds =
      boundsCmd orElse viewCmd orElse cellCmd orElse searchCmd orElse statsCmd orElse filterCmd

    val sheetReadOnlyOpts = (fileOpt, sheetOpt, maxSizeOpt, streamOpt, sheetReadOnlySubcmds).mapN {
      (file, sheet, maxSize, stream, cmd) =>
        Main.run(file, sheet, None, None, None, maxSize, stream, cmd, io)
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
        sheetWriteSubcmds
      ).mapN { (file, sheet, outOpt, inPlace, backend, maxSize, stream, policy, cmd) =>
        runWithOutput(outOpt, inPlace, file, io) { (out, display) =>
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
            io = io
          )
        }
      }

    // Standalone: no --file required (creates new files)
    val standaloneOpts = newCmd.map { case (outPath, sheetName, sheets, backend) =>
      runStandalone(outPath, sheetName, sheets, backend, io)
    }

    // Diff: compares -f against -g (two inputs, no output); custom exit codes
    val diffOpts = (fileOpt, sheetOpt, maxSizeOpt, diffCmd).mapN { (file, sheet, maxSize, cmd) =>
      cmd match
        case CliCommand.Diff(file2, format) => runDiff(file, file2, sheet, maxSize, format, io)
        case other => internal(s"Unexpected diff command: $other", io)
    }

    // Lint: raw-zip structural validation (GH-397, no output file); custom exit codes.
    // The file arrives via -f or positionally (GH-422); exactly one form must be used.
    val lintOpts = (fileOpt.orNone, lintCmd).mapN { case (flagFile, (cmd, positional)) =>
      cmd match
        case CliCommand.Lint(format) =>
          resolveLintFile(flagFile, positional) match
            case Right(file) => runLint(file, format, io)
            case Left(msg) =>
              val error = CliError.usage(msg, None)
              Diagnostics.report(error, io).as(error.exitCode)
        case other => internal(s"Unexpected lint command: $other", io)
    }

    // Info commands: no file required
    val infoOpts = functionsCmd.map(_ => runInfo(io))
    val rasterOpts = rasterizersCmd.map(_ => runRasterizers(io))

    // Batch dry-run: only needs batch source, no --file or --output
    val dryRunFlag =
      Opts.flag("dry-run", "Validate batch JSON without writing")
    val batchDryRunOpts =
      Opts
        .subcommand("batch", batchHelp) {
          (batchArg, dryRunFlag).mapN((src, _) => src)
        }
        .map(src => batchDryRun(src, io).flatMap(io.out).as(ExitCode.Success))

    rasterOpts orElse infoOpts orElse standaloneOpts orElse diffOpts orElse lintOpts orElse headlessOpts orElse sheetsOpts orElse workbookOpts orElse sheetReadOnlyOpts orElse batchDryRunOpts orElse sheetWriteOpts

  /** The parser the binary runs: the program plus `--version`, under decline's `--help`. */
  def command(io: CliIO): Command[IO[ExitCode]] =
    Command(name, header, helpFlag = true)(versionFlag(io) orElse program(io))

  /** argv to exit code: parse, then run the handler or render help (see the object doc). */
  def run(args: List[String], io: CliIO): IO[ExitCode] =
    IO(command(io).parse(args, sys.env)).flatMap {
      case Right(handler) => handler
      case Left(help) =>
        io.err(help.toString)
          .as(if help.errors.nonEmpty then ExitCodes.usage else ExitCodes.ok)
    }

  /** A dispatch arm the parser cannot reach: a defect, reported like any other failure (exit 3). */
  private def internal(message: String, io: CliIO): IO[ExitCode] =
    val error = CliError(ErrorCode.INTERNAL, message)
    Diagnostics.report(error, io).as(error.exitCode)

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
