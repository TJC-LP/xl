package com.tjclp.xl.agent.anthropic

import cats.effect.{IO, Ref}
import cats.effect.std.Queue
import cats.syntax.all.*
import com.anthropic.models.beta.messages.*
import com.tjclp.xl.agent.{AgentEvent, SubTurnUsage, TurnUsage}
import io.circe.{Json, parser as circeParser}

import java.time.Instant
import scala.collection.mutable
import scala.jdk.OptionConverters.*
import scala.util.Try

/** Per-turn token tracking state (immutable, used with Ref for thread safety) */
case class ProcessorState(
  turnNum: Int = 0,
  turnStartTime: Option[Instant] = None,
  prevCumulativeInput: Long = 0L,
  prevCumulativeOutput: Long = 0L,
  lastCumulativeInput: Long = 0L,
  lastCumulativeOutput: Long = 0L,
  // Sub-turn tracking for code execution (assistant → tool → result cycles)
  subTurnNum: Int = 0,
  subTurnStartTime: Option[Instant] = None,
  subTurnAnchorTime: Option[Instant] = None,
  pendingToolCall: Boolean = false
)

/** Processes streaming events from the Anthropic API and emits AgentEvents */
class StreamEventProcessor private (
  eventQueue: Queue[IO, AgentEvent],
  state: Ref[IO, ProcessorState],
  onEvent: AgentEvent => IO[Unit],
  verbose: Boolean
):
  // Track accumulated input JSON per content block index
  private val toolInputBuffers = mutable.Map[Long, StringBuilder]()
  // Track tool use IDs per content block index
  private val toolUseIds = mutable.Map[Long, String]()
  // Track current tool result's tool use ID
  private var lastToolUseId: String = ""

  private def startSubTurnIfNeeded(now: Instant, toolCall: Boolean): IO[Unit] =
    state.update { s =>
      if s.subTurnStartTime.isEmpty then
        val startTime = s.subTurnAnchorTime.getOrElse(now)
        s.copy(
          subTurnNum = s.subTurnNum + 1,
          subTurnStartTime = Some(startTime),
          subTurnAnchorTime = None,
          pendingToolCall = toolCall
        )
      else if toolCall && !s.pendingToolCall then s.copy(pendingToolCall = true)
      else s
    }

  private def emitSubTurnComplete(endTime: Instant): IO[Unit] =
    for
      s <- state.get
      _ <- s.subTurnStartTime match
        case Some(startTime) =>
          val durationMs = java.time.Duration.between(startTime, endTime).toMillis
          val usage = SubTurnUsage(
            subTurnNum = s.subTurnNum,
            durationMs = durationMs,
            hasToolCall = s.pendingToolCall
          )
          val subTurnEvent = AgentEvent.SubTurnComplete(usage)
          val emitSubTurn = eventQueue.offer(subTurnEvent)
          val callbackSubTurn = onEvent(subTurnEvent)
          val resetSubTurn =
            state.update(
              _.copy(
                subTurnStartTime = None,
                subTurnAnchorTime = Some(endTime),
                pendingToolCall = false
              )
            )
          val printSubTurnIO = IO.whenA(verbose) {
            IO(println(s"\n>>> Sub-turn ${s.subTurnNum} complete [${durationMs}ms]"))
          }
          emitSubTurn *> callbackSubTurn *> resetSubTurn *> printSubTurnIO
        case None => IO.unit
    yield ()

  /** Process a single streaming event, potentially emitting AgentEvents */
  def process(event: BetaRawMessageStreamEvent): IO[Unit] =
    val textIO = processTextDelta(event)
    val toolStartIO = processToolStart(event)
    val toolResultIO = processToolResult(event)
    val toolStopIO = processToolStop(event)
    val messageStartIO = processMessageStart(event)
    val messageDeltaIO = processMessageDelta(event)
    val messageStopIO = processMessageStop(event)
    textIO *> toolStartIO *> toolResultIO *> toolStopIO *>
      messageStartIO *> messageDeltaIO *> messageStopIO

  private def processTextDelta(event: BetaRawMessageStreamEvent): IO[Unit] =
    event.contentBlockDelta().toScala match
      case Some(delta) =>
        val index = delta.index()

        // Handle text deltas - emit to queue, call callback, AND print if verbose
        val textIO = delta.delta().text().toScala match
          case Some(textDelta) =>
            val text = textDelta.text()
            val event = AgentEvent.TextOutput(text, index.toInt)
            val printIO = IO.whenA(verbose)(IO(print(text)))
            val emitIO = eventQueue.offer(event)
            val callbackIO = onEvent(event)
            // Start a new sub-turn if not already in one (first text after a tool result)
            val startSubTurnIO =
              IO.realTimeInstant.flatMap(now => startSubTurnIfNeeded(now, toolCall = false))
            startSubTurnIO *> printIO *> emitIO *> callbackIO.void
          case None => IO.unit

        // Accumulate tool input JSON
        val jsonIO = IO {
          delta.delta().inputJson().toScala.foreach { jsonDelta =>
            val partial = jsonDelta.partialJson()
            toolInputBuffers.getOrElseUpdate(index, new StringBuilder()).append(partial)
          }
        }

        textIO *> jsonIO

      case None => IO.unit

  private def processToolStart(event: BetaRawMessageStreamEvent): IO[Unit] =
    event.contentBlockStart().toScala match
      case Some(start) =>
        val block = start.contentBlock()
        val index = start.index()

        // Server tool use (code_execution)
        val serverToolIO = block.serverToolUse().toScala match
          case Some(toolUse) =>
            for
              now <- IO.realTimeInstant
              _ <- startSubTurnIfNeeded(now, toolCall = true)
              _ <- IO {
                if verbose then println(s"\n>>> TOOL: ${toolUse.name()} [${toolUse.id()}]")
                toolInputBuffers(index) = new StringBuilder()
                toolUseIds(index) = toolUse.id()
                lastToolUseId = toolUse.id()
              }
            yield ()
          case None => IO.unit

        // MCP tool use (skill invocations) - marks that a tool call is pending
        val mcpToolIO = block.mcpToolUse().toScala match
          case Some(_) =>
            IO.realTimeInstant.flatMap(now => startSubTurnIfNeeded(now, toolCall = true))
          case None => IO.unit

        serverToolIO *> mcpToolIO

      case None => IO.unit

  private def processToolResult(event: BetaRawMessageStreamEvent): IO[Unit] =
    event.contentBlockStart().toScala match
      case Some(start) =>
        val block = start.contentBlock()

        import scala.jdk.CollectionConverters.*

        val bashResultOpt = block.bashCodeExecutionToolResult().toScala
        val mcpResultOpt = block.mcpToolResult().toScala
        val viewResultOpt = block.textEditorCodeExecutionToolResult().toScala
        val hasToolResult =
          bashResultOpt.isDefined || mcpResultOpt.isDefined || viewResultOpt.isDefined

        // Bash execution results
        val bashIO = bashResultOpt match
          case Some(result) =>
            result.content().betaBashCodeExecutionResultBlock().toScala match
              case Some(r) =>
                val stdout = r.stdout()
                val stderr = r.stderr()
                val fileIds = r.content().asScala.map(_.fileId()).toList
                val event = AgentEvent.ToolResult(lastToolUseId, stdout, stderr, None, fileIds)
                val emitResult = eventQueue.offer(event)
                val callbackIO = onEvent(event)

                val printIO = IO.whenA(verbose) {
                  IO {
                    println(s"<<< RESULT [${lastToolUseId.take(20)}...]")
                    if stdout.nonEmpty then
                      val lines = stdout.split("\n")
                      lines.take(20).foreach(line => println(s"    $line"))
                      if lines.length > 20 then
                        println(s"    ... (${lines.length - 20} more lines)")
                    if stderr.nonEmpty then
                      println("    STDERR:")
                      stderr.split("\n").take(5).foreach(line => println(s"    $line"))
                    if fileIds.nonEmpty then println(s"    FILES: ${fileIds.mkString(", ")}")
                  }
                }

                emitResult *> callbackIO *> printIO

              case None =>
                // Handle errors
                result.content().betaBashCodeExecutionToolResultError().toScala match
                  case Some(err) =>
                    val errMsg = err.errorCode().toString
                    val event = AgentEvent.Error(errMsg)
                    val emitError = eventQueue.offer(event)
                    val callbackIO = onEvent(event)
                    val printIO = IO.whenA(verbose)(IO(println(s"<<< ERROR: $errMsg")))
                    emitError *> callbackIO *> printIO
                  case None => IO.unit

          case None => IO.unit

        // MCP tool results (skill documentation, etc.)
        val mcpIO = mcpResultOpt match
          case Some(result) =>
            val content = extractMcpContent(result)
            val isError = result.isError()
            val toolUseId = result.toolUseId()
            val event = AgentEvent.McpToolResult(toolUseId, content, isError)
            val emitMcp = eventQueue.offer(event)
            val callbackIO = onEvent(event)
            val printIO = IO.whenA(verbose) {
              IO {
                val header = if isError then "<<< MCP ERROR" else "<<< MCP RESULT"
                println(s"$header [$toolUseId]")
                content.split("\n").take(10).foreach(line => println(s"    $line"))
                if content.split("\n").length > 10 then
                  println(s"    ... (${content.split("\n").length - 10} more lines)")
              }
            }
            emitMcp *> callbackIO *> printIO
          case None => IO.unit

        // Text editor code execution results (includes view results)
        val viewIO = viewResultOpt match
          case Some(toolResult) =>
            // Extract content from the result - it may have different structures
            val content = Try {
              toolResult
                ._content()
                .asKnown()
                .toScala
                .flatMap(_._json().toScala)
                .map(_.toString)
                .getOrElse("")
            }.getOrElse("")

            if content.nonEmpty then
              val event = AgentEvent.ViewResult(content)
              val emitView = eventQueue.offer(event)
              val callbackIO = onEvent(event)
              val printIO = IO.whenA(verbose) {
                IO {
                  println(s"<<< VIEW RESULT")
                  content.split("\n").take(10).foreach(line => println(s"    $line"))
                  if content.split("\n").length > 10 then
                    println(s"    ... (${content.split("\n").length - 10} more lines)")
                }
              }
              emitView *> callbackIO *> printIO
            else IO.unit
          case None => IO.unit

        val resultsIO = bashIO *> mcpIO *> viewIO

        if !hasToolResult then IO.unit
        else
          for
            endTime <- IO.realTimeInstant
            _ <- startSubTurnIfNeeded(endTime, toolCall = true)
            _ <- resultsIO
            _ <- emitSubTurnComplete(endTime)
          yield ()

      case None => IO.unit

  private def processToolStop(event: BetaRawMessageStreamEvent): IO[Unit] =
    event.contentBlockStop().toScala match
      case Some(stop) =>
        val index = stop.index()
        toolInputBuffers.get(index) match
          case Some(buffer) =>
            val jsonStr = buffer.toString
            val toolUseId = toolUseIds.getOrElse(index, "")
            toolInputBuffers.remove(index)
            toolUseIds.remove(index)

            if jsonStr.nonEmpty then
              // Parse full JSON input
              val fullInput = circeParser.parse(jsonStr).getOrElse(Json.Null)
              // Also extract command for convenience
              val command = extractCodeFromJson(jsonStr)

              val event = AgentEvent.ToolInvocation("code_execution", toolUseId, fullInput, command)
              val emitTool = eventQueue.offer(event)
              val callbackIO = onEvent(event)
              val printIO = IO.whenA(verbose) {
                IO {
                  command.foreach { code =>
                    println("--- CODE ---")
                    code.split("\n").take(30).foreach(line => println(s"    $line"))
                    if code.split("\n").length > 30 then
                      println(s"    ... (${code.split("\n").length - 30} more lines)")
                    println("------------")
                  }
                }
              }
              emitTool *> callbackIO *> printIO
            else IO.unit
          case None => IO.unit
      case None => IO.unit

  /** Track message start for per-turn timing and capture input tokens */
  private def processMessageStart(event: BetaRawMessageStreamEvent): IO[Unit] =
    event.messageStart().toScala match
      case Some(start) =>
        for
          now <- IO.realTimeInstant
          usage = start.message().usage()
          _ <- state.update { s =>
            s.copy(
              turnNum = s.turnNum + 1,
              turnStartTime = Some(now),
              prevCumulativeInput = s.lastCumulativeInput,
              prevCumulativeOutput = s.lastCumulativeOutput,
              lastCumulativeInput = usage.inputTokens(),
              subTurnStartTime = None,
              subTurnAnchorTime = Some(now),
              pendingToolCall = false
            )
          }
        yield ()
      case None => IO.unit

  /** Capture cumulative token usage from message_delta events */
  private def processMessageDelta(event: BetaRawMessageStreamEvent): IO[Unit] =
    event.messageDelta().toScala match
      case Some(delta) =>
        // Extract cumulative usage from the delta event
        // BetaMessageDeltaUsage provides cumulative tokens for the entire message
        val usage = delta.usage()
        state.update { s =>
          s.copy(
            lastCumulativeInput =
              usage.inputTokens().toScala.map(_.toLong).getOrElse(s.lastCumulativeInput),
            lastCumulativeOutput = usage.outputTokens()
          )
        }
      case None => IO.unit

  /** Emit TurnComplete event when a message (turn) ends */
  private def processMessageStop(event: BetaRawMessageStreamEvent): IO[Unit] =
    event.messageStop().toScala match
      case Some(_) =>
        for
          now <- IO.realTimeInstant
          _ <- emitSubTurnComplete(now)
          s <- state.get
          durationMs = s.turnStartTime
            .map { start =>
              java.time.Duration.between(start, now).toMillis
            }
            .getOrElse(0L)

          inputDelta = s.lastCumulativeInput - s.prevCumulativeInput
          outputDelta = s.lastCumulativeOutput - s.prevCumulativeOutput

          usage = TurnUsage(
            turnNum = s.turnNum,
            inputTokens = inputDelta,
            outputTokens = outputDelta,
            cumulativeInputTokens = s.lastCumulativeInput,
            cumulativeOutputTokens = s.lastCumulativeOutput,
            durationMs = durationMs
          )

          turnEvent = AgentEvent.TurnComplete(usage)
          _ <- eventQueue.offer(turnEvent)
          _ <- onEvent(turnEvent)
          _ <- IO.whenA(verbose) {
            IO {
              println(
                s"\n>>> TURN ${s.turnNum} complete: +$inputDelta in / +$outputDelta out (cumulative: ${s.lastCumulativeInput} / ${s.lastCumulativeOutput}) [${durationMs}ms]"
              )
            }
          }
        yield ()

      case None => IO.unit

  /** Extract the 'command' field from tool input JSON */
  private def extractCodeFromJson(json: String): Option[String] =
    circeParser.parse(json).toOption.flatMap { parsed =>
      parsed.hcursor.get[String]("command").toOption
    }

  /** Extract content from MCP tool result (handles String | List[BetaTextBlock] union) */
  private def extractMcpContent(result: BetaMcpToolResultBlock): String =
    import scala.jdk.CollectionConverters.*
    Try {
      // Try to get as string first (using Optional accessor)
      result.content().string().toScala.getOrElse {
        // Fall back to text blocks
        result
          .content()
          .betaMcpToolResultBlock()
          .toScala
          .map(_.asScala.map(_.text()).mkString("\n"))
          .getOrElse("")
      }
    }.getOrElse("")

object StreamEventProcessor:
  /** Create a new StreamEventProcessor with thread-safe atomic state */
  def create(
    eventQueue: Queue[IO, AgentEvent],
    onEvent: AgentEvent => IO[Unit] = _ => IO.unit,
    verbose: Boolean = false
  ): IO[StreamEventProcessor] =
    Ref.of[IO, ProcessorState](ProcessorState()).map { stateRef =>
      new StreamEventProcessor(eventQueue, stateRef, onEvent, verbose)
    }
