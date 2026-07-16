package com.tjclp.xl.agent

import cats.effect.{Clock, IO, Ref, Resource}
import cats.effect.std.Queue
import cats.syntax.all.*
import com.anthropic.models.beta.messages.BetaMessage
import com.tjclp.xl.agent.anthropic.{AnthropicClientIO, CodeExecution, SkillsApi}
import com.tjclp.xl.agent.approach.{ApproachStrategy, XlApproachStrategy}
import com.tjclp.xl.agent.benchmark.common.FileManager
import com.tjclp.xl.agent.error.AgentError

import java.nio.file.{Files, Path}
import java.util.concurrent.TimeUnit
import scala.jdk.OptionConverters.*

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

  /**
   * One logical turn's API responses: the pause_turn responses the API `paused` (oldest first)
   * followed by the `last` response that ended the resume loop. Every element was a separately
   * billed request, so result merging must span all of them.
   */
  private[agent] final case class TurnResponses(paused: Vector[BetaMessage], last: BetaMessage):
    def all: Vector[BetaMessage] = paused :+ last
    def resumesUsed: Int = paused.size

  /**
   * Drive one logical turn, auto-resuming while the API pauses it (stop_reason=pause_turn, issue
   * #344). `send` receives the paused turns accumulated so far — empty on the initial call — and
   * must re-send the original request with them appended verbatim (see
   * `CodeExecution.buildParams`). Caps at `maxResumes` re-sends; a response still paused at the cap
   * is returned as-is and surfaces through `StopReasonPolicy.errorFor(stopReason, resumesUsed)`.
   */
  private[agent] def sendWithPauseTurnResume(
    send: Vector[BetaMessage] => IO[BetaMessage],
    maxResumes: Int = StopReasonPolicy.MaxPauseTurnResumes
  ): IO[TurnResponses] =
    def go(paused: Vector[BetaMessage]): IO[TurnResponses] =
      send(paused).flatMap { response =>
        val stopReason = response.stopReason().toScala.map(_.asString)
        if stopReason.contains(StopReasonPolicy.PauseTurn) && paused.size < maxResumes then
          go(paused :+ response)
        else IO.pure(TurnResponses(paused, response))
      }
    go(Vector.empty)

  /** Token usage of a single API response */
  private[agent] def usageOf(response: BetaMessage): TokenUsage =
    TokenUsage(
      inputTokens = response.usage().inputTokens(),
      outputTokens = response.usage().outputTokens(),
      cacheCreationTokens =
        response.usage().cacheCreationInputTokens().toScala.map(Long.unbox).getOrElse(0L),
      cacheReadTokens =
        response.usage().cacheReadInputTokens().toScala.map(Long.unbox).getOrElse(0L)
    )

  /** Summed usage across a turn's responses (each pause_turn resume is billed separately) */
  private[agent] def mergedUsage(turn: TurnResponses): TokenUsage =
    turn.all.foldLeft(TokenUsage.zero)((acc, response) => acc + usageOf(response))

  /** Concatenated visible text across a turn's responses (a resume does not re-emit paused text) */
  private[agent] def mergedResponseText(turn: TurnResponses): String =
    turn.all.map(CodeExecution.extractResponseText).filter(_.nonEmpty).mkString("\n")

  /** Last output file id across a turn's responses (the file may pre-date the final resume) */
  private[agent] def mergedOutputFileId(turn: TurnResponses, verbose: Boolean): Option[String] =
    turn.all.flatMap(response => CodeExecution.extractOutputFileId(response, verbose)).lastOption

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

        // Send request with streaming - pass onEvent for real-time tracing. A turn the API
        // pauses (stop_reason=pause_turn) is auto-resumed with the paused assistant turns
        // appended verbatim, bounded by StopReasonPolicy.MaxPauseTurnResumes (issue #344).
        turn <- sendWithPauseTurnResume { resumeTurns =>
          CodeExecution.sendRequest(
            client.underlying,
            config,
            systemPrompt,
            userPrompt,
            containerUploads = strategy.containerUploads(inputFile.id),
            eventQueue,
            configureRequest = strategy.configureRequest,
            onEvent = onEvent,
            resumeTurns = resumeTurns
          )
        }

        // Drain event queue and collect events (onEvent already called during streaming)
        collectedEvents <- drainQueue(eventQueue, events)

        // Extract response info across the whole turn (paused responses + final response)
        responseText = mergedResponseText(turn)
        outputFileId = mergedOutputFileId(turn, config.verbose)
        // Final stop_reason of the accumulated message: max_tokens/refusal (or a pause_turn
        // that survived the resume loop) must not read as clean completions
        stopReason = turn.last.stopReason().toScala.map(_.asString)
        usage = mergedUsage(turn)

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
      // Don't set error for missing output files - let grading layer handle this;
      // error only reports non-clean stops (truncation/refusal/resume exhaustion), issue #340.
      yield AgentResult(
        success = outputPath.isDefined || responseText.nonEmpty,
        outputFileId = outputFileId,
        outputPath = outputPath,
        usage = usage,
        latencyMs = latencyMs,
        transcript = collectedEvents,
        responseText = Some(responseText),
        error = StopReasonPolicy.errorFor(stopReason, turn.resumesUsed),
        stopReason = stopReason
      )

    private def drainQueue(
      queue: Queue[IO, AgentEvent],
      events: Ref[IO, Vector[AgentEvent]]
    ): IO[Vector[AgentEvent]] =
      // Just collect events from queue - onEvent was already called during streaming
      def loop: IO[Unit] =
        queue.tryTake.flatMap {
          case Some(event) =>
            events.update(_ :+ event) *> loop
          case None =>
            IO.unit
        }
      loop *> events.get

  /** Create an Agent resource that manages uploaded files (xl-cli approach) */
  def resource(config: AgentConfig): Resource[IO, Agent] =
    for
      client <- AnthropicClientIO.fromEnv

      // Resolve binary and skill paths (version-agnostic, auto-downloads latest release)
      binaryPath <- Resource.eval(FileManager.resolveBinaryPath(config.xlBinaryPath))
      skillPath <- Resource.eval(FileManager.resolveSkillPath(config.xlSkillPath))

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
