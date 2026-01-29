package com.tjclp.xl.agent

import cats.effect.{Clock, IO, Ref, Resource}
import cats.effect.std.Queue
import cats.syntax.all.*
import com.tjclp.xl.agent.anthropic.{AnthropicClientIO, CodeExecution, SkillsApi}
import com.tjclp.xl.agent.approach.{ApproachStrategy, XlApproachStrategy}
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

  /** Create an Agent instance with the given configuration and approach strategy */
  def create(
    client: AnthropicClientIO,
    config: AgentConfig,
    strategy: ApproachStrategy
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

        // Build prompts using strategy
        systemPrompt = strategy.systemPrompt
        userPrompt = strategy.userPrompt(task, task.inputFile.getFileName.toString)

        // Send request with streaming - pass onEvent for real-time tracing
        response <- CodeExecution.sendRequest(
          client.underlying,
          config,
          systemPrompt,
          userPrompt,
          containerUploads = strategy.containerUploads(inputFile.id),
          eventQueue,
          configureRequest = strategy.configureRequest,
          onEvent = onEvent
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
      // For file modification tasks, success = output file created.
      // For analysis tasks (no expected output), success = agent completed.
      // Don't set error for missing output files - let grading layer handle this.
      yield AgentResult(
        success = outputPath.isDefined || responseText.nonEmpty,
        outputFileId = outputFileId,
        outputPath = outputPath,
        usage = usage,
        latencyMs = latencyMs,
        transcript = collectedEvents,
        responseText = Some(responseText),
        error = None // Grading determines correctness, not output file presence
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

  /** Create an Agent resource that manages uploaded files (xl-cli approach) */
  def resource(config: AgentConfig): Resource[IO, Agent] =
    for
      client <- AnthropicClientIO.fromEnv

      // Resolve binary and skill paths
      binaryPath <- Resource.eval(resolveBinaryPath(config))
      skillPath <- Resource.eval(resolveSkillPath(config))

      // Upload binary file
      binaryFile <- Resource.make(client.uploadFile(binaryPath))(f =>
        client.deleteFile(f.id).attempt.void
      )

      // Register skill via Skills API (SKILL.md auto-indexed for token efficiency)
      apiKey <- Resource.eval(AnthropicClientIO.loadApiKey)
      skillId <- Resource.eval(SkillsApi.getOrCreateXlSkill(apiKey, skillPath))

      // Create xl-cli approach strategy with skill ID
      strategy = XlApproachStrategy(binaryFile, skillId)
    yield create(client, config, strategy)

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
