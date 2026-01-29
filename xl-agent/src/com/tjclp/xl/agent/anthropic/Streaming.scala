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

        // Handle text deltas - emit to queue AND print if verbose
        val textIO = delta.delta().text().toScala match
          case Some(textDelta) =>
            val text = textDelta.text()
            val printIO = IO.whenA(verbose)(IO(print(text)))
            val emitIO = eventQueue.offer(AgentEvent.TextOutput(text, index.toInt))
            printIO *> emitIO.void
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

        // Bash execution results
        block.bashCodeExecutionToolResult().toScala match
          case Some(result) =>
            result.content().betaBashCodeExecutionResultBlock().toScala match
              case Some(r) =>
                val stdout = r.stdout()
                val stderr = r.stderr()
                // Extract file IDs from result content
                import scala.jdk.CollectionConverters.*
                val fileIds = r.content().asScala.map(_.fileId()).toList
                val emitResult = eventQueue.offer(
                  AgentEvent.ToolResult(lastToolUseId, stdout, stderr, None, fileIds)
                )

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

                emitResult *> printIO

              case None =>
                // Handle errors
                result.content().betaBashCodeExecutionToolResultError().toScala match
                  case Some(err) =>
                    val errMsg = err.errorCode().toString
                    val emitError = eventQueue.offer(AgentEvent.Error(errMsg))
                    val printIO = IO.whenA(verbose)(IO(println(s"<<< ERROR: $errMsg")))
                    emitError *> printIO
                  case None => IO.unit

          case None => IO.unit

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

              val emitTool = eventQueue.offer(
                AgentEvent.ToolInvocation("code_execution", toolUseId, fullInput, command)
              )
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
              emitTool *> printIO
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
