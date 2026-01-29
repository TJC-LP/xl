package com.tjclp.xl.agent.anthropic

import cats.effect.IO
import cats.effect.std.Queue
import cats.syntax.all.*
import com.anthropic.models.beta.messages.*
import com.tjclp.xl.agent.AgentEvent
import io.circe.{Json, parser as circeParser}

import scala.collection.mutable
import scala.jdk.OptionConverters.*
import scala.util.Try

/** Processes streaming events from the Anthropic API and emits AgentEvents */
class StreamEventProcessor(
  eventQueue: Queue[IO, AgentEvent],
  onEvent: AgentEvent => IO[Unit] = _ => IO.unit,
  verbose: Boolean = false
):
  // Track accumulated input JSON per content block index
  private val toolInputBuffers = mutable.Map[Long, StringBuilder]()
  // Track tool use IDs per content block index
  private val toolUseIds = mutable.Map[Long, String]()
  // Track current tool result's tool use ID
  private var lastToolUseId: String = ""

  /** Process a single streaming event, potentially emitting AgentEvents */
  def process(event: BetaRawMessageStreamEvent): IO[Unit] =
    val textIO = processTextDelta(event)
    val toolStartIO = processToolStart(event)
    val toolResultIO = processToolResult(event)
    val toolStopIO = processToolStop(event)
    textIO *> toolStartIO *> toolResultIO *> toolStopIO

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
            printIO *> emitIO *> callbackIO.void
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
        block.serverToolUse().toScala match
          case Some(toolUse) =>
            IO {
              if verbose then println(s"\n>>> TOOL: ${toolUse.name()} [${toolUse.id()}]")
              toolInputBuffers(index) = new StringBuilder()
              toolUseIds(index) = toolUse.id()
              lastToolUseId = toolUse.id()
            }
          case None => IO.unit

      case None => IO.unit

  private def processToolResult(event: BetaRawMessageStreamEvent): IO[Unit] =
    event.contentBlockStart().toScala match
      case Some(start) =>
        val block = start.contentBlock()

        import scala.jdk.CollectionConverters.*

        // Bash execution results
        val bashIO = block.bashCodeExecutionToolResult().toScala match
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
        val mcpIO = block.mcpToolResult().toScala match
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
        val viewIO = block.textEditorCodeExecutionToolResult().toScala match
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

        bashIO *> mcpIO *> viewIO

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

  /** Extract the 'command' field from tool input JSON */
  private def extractCodeFromJson(json: String): Option[String] =
    Try {
      val commandPattern = """"command"\s*:\s*"((?:[^"\\]|\\.)*)"""".r
      commandPattern.findFirstMatchIn(json).map { m =>
        m.group(1)
          .replace("\\n", "\n")
          .replace("\\t", "\t")
          .replace("\\\"", "\"")
          .replace("\\\\", "\\")
      }
    }.toOption.flatten

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
