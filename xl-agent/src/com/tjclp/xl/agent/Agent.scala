package com.tjclp.xl.agent

import cats.effect.{Clock, IO, Ref, Resource}
import cats.effect.std.Queue
import cats.syntax.all.*
import com.tjclp.xl.agent.anthropic.{AnthropicClientIO, CodeExecution}
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.{Files, Path}
import java.util.concurrent.TimeUnit

/** A task for the agent to execute */
case class AgentTask(
  instruction: String,
  inputFile: Path,
  outputFile: Path,
  answerPosition: Option[String] = None,
  binaryName: String = "xl",
  skillName: String = "xl-skill.zip"
)

/** Core agent API for executing LLM-driven Excel operations */
trait Agent:
  /** Run a task and return the result */
  def run(task: AgentTask): IO[AgentResult]

  /** Run a task with streaming callbacks */
  def runStreaming(
    task: AgentTask,
    onEvent: AgentEvent => IO[Unit]
  ): IO[AgentResult]

object Agent:

  /** Create an Agent instance with the given configuration */
  def create(
    client: AnthropicClientIO,
    config: AgentConfig,
    binaryFile: UploadedFile,
    skillFile: UploadedFile
  ): Agent = new Agent:

    override def run(task: AgentTask): IO[AgentResult] =
      runStreaming(task, _ => IO.unit)

    override def runStreaming(
      task: AgentTask,
      onEvent: AgentEvent => IO[Unit]
    ): IO[AgentResult] =
      for
        startTime <- Clock[IO].monotonic.map(_.toMillis)

        // Create event queue for streaming
        eventQueue <- Queue.unbounded[IO, AgentEvent]
        events <- Ref.of[IO, Vector[AgentEvent]](Vector.empty)

        // Upload input file
        inputFile <- client.uploadFile(task.inputFile)

        // Build prompts using actual uploaded filenames
        systemPrompt = buildSystemPrompt(binaryFile.filename, skillFile.filename)
        userPrompt = buildUserPrompt(task, binaryFile.filename)

        // Send request with streaming
        response <- CodeExecution.sendRequest(
          client.underlying,
          config,
          systemPrompt,
          userPrompt,
          containerUploads = List(binaryFile.id, skillFile.id, inputFile.id),
          eventQueue
        )

        // Drain event queue and collect events
        collectedEvents <- drainQueue(eventQueue, events, onEvent)

        // Extract response info
        responseText = CodeExecution.extractResponseText(response)
        outputFileId = CodeExecution.extractOutputFileId(response, config.verbose)
        usage = TokenUsage(response.usage().inputTokens(), response.usage().outputTokens())

        // Download output file if found
        outputPath <- outputFileId match
          case Some(fileId) =>
            IO.blocking(Files.createDirectories(task.outputFile.getParent)) *>
              client.downloadFile(fileId, task.outputFile) *>
              client.deleteFile(fileId).attempt *> // Cleanup, ignore errors
              IO.pure(Some(task.outputFile))
          case None =>
            IO.pure(None)

        // Cleanup input file
        _ <- client.deleteFile(inputFile.id).attempt

        endTime <- Clock[IO].monotonic.map(_.toMillis)
        latencyMs = endTime - startTime
      yield AgentResult(
        success = outputPath.isDefined,
        outputFileId = outputFileId,
        outputPath = outputPath,
        usage = usage,
        latencyMs = latencyMs,
        transcript = collectedEvents,
        responseText = Some(responseText),
        error = if outputPath.isEmpty then Some("No output file created") else None
      )

    private def drainQueue(
      queue: Queue[IO, AgentEvent],
      events: Ref[IO, Vector[AgentEvent]],
      onEvent: AgentEvent => IO[Unit]
    ): IO[Vector[AgentEvent]] =
      def loop: IO[Unit] =
        queue.tryTake.flatMap {
          case Some(event) =>
            events.update(_ :+ event) *> onEvent(event) *> loop
          case None =>
            IO.unit
        }
      loop *> events.get

  /** Create an Agent resource that manages uploaded files */
  def resource(config: AgentConfig): Resource[IO, Agent] =
    for
      client <- AnthropicClientIO.fromEnv

      // Resolve binary and skill paths
      binaryPath <- Resource.eval(resolveBinaryPath(config))
      skillPath <- Resource.eval(resolveSkillPath(config))

      // Upload binary and skill files
      binaryFile <- Resource.make(client.uploadFile(binaryPath))(f =>
        client.deleteFile(f.id).attempt.void
      )
      skillFile <- Resource.make(client.uploadFile(skillPath))(f =>
        client.deleteFile(f.id).attempt.void
      )
    yield create(client, config, binaryFile, skillFile)

  private def resolveBinaryPath(config: AgentConfig): IO[Path] =
    config.xlBinaryPath match
      case Some(p) => IO.pure(p)
      case None =>
        findFile(
          "xl-0.8.1-linux-amd64",
          List("../benchmark", "examples/anthropic-sdk/benchmark", ".")
        )

  private def resolveSkillPath(config: AgentConfig): IO[Path] =
    config.xlSkillPath match
      case Some(p) => IO.pure(p)
      case None =>
        findFile(
          "xl-skill-0.8.1.zip",
          List("../benchmark", "examples/anthropic-sdk/benchmark", ".")
        )

  private def findFile(name: String, dirs: List[String]): IO[Path] =
    IO.blocking {
      import java.nio.file.Paths
      dirs.view
        .map(d => Paths.get(d, name))
        .find(p => Files.exists(p))
        .getOrElse(Paths.get(dirs.head, name))
    }

  private def buildSystemPrompt(binaryName: String, skillName: String): String =
    s"""You are a spreadsheet expert assistant. You have access to the xl CLI tool for manipulating Excel files.

IMPORTANT REQUIREMENTS:
1. Use the xl CLI tool for ALL Excel operations (not Python libraries like openpyxl)
2. ALWAYS save output to $$OUTPUT_DIR/output.xlsx using the -o flag
3. When you finish modifying the spreadsheet, the final file MUST be at $$OUTPUT_DIR/output.xlsx
4. DO NOT modify the input data - only add formulas or values where specified

AVAILABLE FILES:
- $binaryName: The xl CLI binary (chmod +x it first, then use ./$binaryName or move to PATH)
- $skillName: Documentation for the xl CLI (unzip to read SKILL.md for command reference)
- Input xlsx file: The spreadsheet to work with

Your goal is to solve spreadsheet tasks efficiently and accurately. Always:
1. Explore the input file first to understand its structure
2. Plan your solution before executing
3. Use xl CLI commands with -o /files/output/output.xlsx for writing
4. Verify your solution by viewing the output cells
5. NEVER modify input data cells - only add to the cells specified in the task

The xl CLI supports viewing data, writing values, writing formulas, and batch operations."""

  private def buildUserPrompt(task: AgentTask, binaryName: String): String =
    val answerSection = task.answerPosition
      .map(pos => s"\nThe answer will be evaluated at position: $pos\n")
      .getOrElse("")

    s"""You are a spreadsheet expert with access to the xl CLI for Excel operations.

## Task
${task.instruction}

## File Locations
- Input file: $$INPUT_DIR/${task.inputFile.getFileName}
- XL binary: $$INPUT_DIR/$binaryName
$answerSection
## CRITICAL: Output Location
You MUST save your final output to: $$OUTPUT_DIR/output.xlsx
The $$OUTPUT_DIR environment variable is set by the system. Using it ensures the file is properly captured.

## Setup Commands (run these first)
```bash
chmod +x $$INPUT_DIR/$binaryName
XL="$$INPUT_DIR/$binaryName"
INPUT="$$INPUT_DIR/${task.inputFile.getFileName}"
OUTPUT="$$OUTPUT_DIR/output.xlsx"
```

## xl CLI Commands
```bash
# List sheets
$$XL -f $$INPUT sheets

# View data (use quotes for sheet names with spaces)
$$XL -f $$INPUT -s "<sheet>" view <range>

# Write a value (always use -o $$OUTPUT to save)
$$XL -f $$INPUT -o $$OUTPUT put <ref> <value>

# Write a formula (always use -o $$OUTPUT to save)
$$XL -f $$INPUT -o $$OUTPUT putf <ref> "=FORMULA"

# Batch operations (always use -o $$OUTPUT to save)
echo '[{"op":"put","ref":"A1","value":"Hello"}]' | $$XL -f $$INPUT -o $$OUTPUT batch -
```

## Instructions
1. Explore the input file to understand its structure
2. Implement the solution using xl CLI commands
3. IMPORTANT: Always use -o $$OUTPUT (which is $$OUTPUT_DIR/output.xlsx) when writing
4. Verify your solution by viewing the output cells${task.answerPosition
        .map(p => s" at $p")
        .getOrElse("")}
"""
