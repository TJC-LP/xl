package com.tjclp.xl.agent.anthropic

import cats.effect.IO
import cats.effect.std.Queue
import cats.syntax.all.*
import com.anthropic.models.beta.messages.*
import com.tjclp.xl.agent.AgentEvent

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

  /** Process a single streaming event, potentially emitting AgentEvents */
  def process(event: BetaRawMessageStreamEvent): IO[Unit] =
    val textIO = processTextDelta(event)
    val toolStartIO = processToolStart(event)
    val toolResultIO = processToolResult(event)
    val toolStopIO = processToolStop(event)
    textIO *> toolStartIO *> toolResultIO *> toolStopIO

  private def processTextDelta(event: BetaRawMessageStreamEvent): IO[Unit] =
    IO {
      event.contentBlockDelta().toScala.foreach { delta =>
        val index = delta.index()

        // Handle text deltas
        delta.delta().text().toScala.foreach { textDelta =>
          val text = textDelta.text()
          if verbose then print(text) // Real-time streaming to console
        }

        // Accumulate tool input JSON
        delta.delta().inputJson().toScala.foreach { jsonDelta =>
          val partial = jsonDelta.partialJson()
          toolInputBuffers.getOrElseUpdate(index, new StringBuilder()).append(partial)
        }
      }
    }

  private def processToolStart(event: BetaRawMessageStreamEvent): IO[Unit] =
    event.contentBlockStart().toScala match
      case Some(start) =>
        val block = start.contentBlock()
        val index = start.index()

        // Server tool use (code_execution)
        block.serverToolUse().toScala match
          case Some(toolUse) =>
            IO {
              if verbose then println(s"\n>>> TOOL: ${toolUse.name()}")
              toolInputBuffers(index) = new StringBuilder()
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
                val emitResult = eventQueue.offer(AgentEvent.ToolResult(stdout, stderr))

                val printIO = IO.whenA(verbose) {
                  IO {
                    println("<<< RESULT")
                    if stdout.nonEmpty then
                      val lines = stdout.split("\n")
                      lines.take(20).foreach(line => println(s"    $line"))
                      if lines.length > 20 then
                        println(s"    ... (${lines.length - 20} more lines)")
                    if stderr.nonEmpty then
                      println("    STDERR:")
                      stderr.split("\n").take(5).foreach(line => println(s"    $line"))
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
            val json = buffer.toString
            toolInputBuffers.remove(index)
            if json.nonEmpty then
              extractCodeFromJson(json) match
                case Some(code) =>
                  val emitTool = eventQueue.offer(AgentEvent.ToolInvocation("code_execution", code))
                  val printIO = IO.whenA(verbose) {
                    IO {
                      println("--- CODE ---")
                      code.split("\n").take(30).foreach(line => println(s"    $line"))
                      if code.split("\n").length > 30 then
                        println(s"    ... (${code.split("\n").length - 30} more lines)")
                      println("------------")
                    }
                  }
                  emitTool *> printIO
                case None => IO.unit
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
