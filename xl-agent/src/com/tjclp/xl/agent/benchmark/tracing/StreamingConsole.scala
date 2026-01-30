package com.tjclp.xl.agent.benchmark.tracing

import cats.effect.IO
import com.tjclp.xl.agent.AgentEvent

/** Real-time colored console output for streaming conversations */
object StreamingConsole:

  // ANSI color codes
  private val Reset = "\u001b[0m"
  private val Cyan = "\u001b[36m"
  private val Yellow = "\u001b[33m"
  private val Green = "\u001b[32m"
  private val Red = "\u001b[31m"
  private val Gray = "\u001b[90m"
  private val Bold = "\u001b[1m"

  /** Print a traced event with colors */
  def print(traced: TracedEvent): IO[Unit] =
    IO {
      val timeStr = f"[${traced.relativeMs / 1000.0}%6.2fs]"

      traced.event match
        case AgentEvent.TextOutput(text, _) =>
          // Stream text in real-time (cyan)
          Console.print(s"$Cyan$text$Reset")
          Console.flush()

        case AgentEvent.ToolInvocation(name, _, _, command) =>
          // Tool call (yellow)
          Console.println(s"\n$Yellow$timeStr TOOL: $name$Reset")
          command.foreach { cmd =>
            val truncated = truncateLines(cmd, 30)
            Console.println(s"$Yellow$$$$Reset $truncated")
          }

        case AgentEvent.ToolResult(_, stdout, stderr, exitCode, files) =>
          // Result (green for success, red for error)
          val success = exitCode.forall(_ == 0)
          val color = if success then Green else Red

          if stdout.nonEmpty then
            val truncated = truncateLines(stdout, 20)
            Console.println(s"$color$truncated$Reset")

          if stderr.nonEmpty then
            Console.println(s"${Red}STDERR: ${truncateLines(stderr, 5)}$Reset")

          if files.nonEmpty then Console.println(s"${Gray}FILES: ${files.mkString(", ")}$Reset")

        case AgentEvent.FileCreated(fileId, filename) =>
          Console.println(s"$Green$timeStr FILE: $filename ($fileId)$Reset")

        case AgentEvent.Error(message) =>
          Console.println(s"$Red$timeStr ERROR: $message$Reset")

        case AgentEvent.McpToolResult(_, content, isError) =>
          // MCP/skill result (skill documentation, etc.)
          val color = if isError then Red else Green
          val header = if isError then "MCP ERROR" else "MCP RESULT"
          Console.println(s"$color$timeStr $header$Reset")
          if content.nonEmpty then
            val truncated = truncateLines(content, 10)
            Console.println(s"$Gray$truncated$Reset")

        case AgentEvent.ViewResult(content) =>
          // Text editor view result
          Console.println(s"$Green$timeStr VIEW RESULT$Reset")
          if content.nonEmpty then
            val truncated = truncateLines(content, 10)
            Console.println(s"$Gray$truncated$Reset")

        case AgentEvent.TurnComplete(usage) =>
          // Turn complete with token usage
          Console.println(
            s"$Gray$timeStr Turn ${usage.turnNum}: +${usage.inputTokens} in / +${usage.outputTokens} out (cumulative: ${usage.cumulativeInputTokens}/${usage.cumulativeOutputTokens})$Reset"
          )
    }

  /** Print a skill header */
  def printSkillHeader(skillName: String): IO[Unit] =
    IO {
      val line = "=" * 60
      Console.println(s"\n$Bold$Cyan$line$Reset")
      Console.println(s"$Bold$Cyan  $skillName$Reset")
      Console.println(s"$Bold$Cyan$line$Reset\n")
    }

  /** Print a task header */
  def printTaskHeader(taskId: String, caseNum: Int): IO[Unit] =
    IO {
      Console.println(s"$Gray--- Task $taskId Case $caseNum ---$Reset")
    }

  /** Print completion message */
  def printComplete(
    skillName: String,
    taskId: String,
    caseNum: Int,
    passed: Boolean,
    durationMs: Long
  ): IO[Unit] =
    IO {
      val status = if passed then s"${Green}PASSED" else s"${Red}FAILED"
      val duration = f"${durationMs / 1000.0}%.1fs"
      Console.println(s"\n$status$Reset Task $taskId:$caseNum ($duration)")
    }

  private def truncateLines(s: String, maxLines: Int): String =
    val lines = s.split("\n")
    if lines.length <= maxLines then s
    else
      val shown = lines.take(maxLines).mkString("\n")
      val remaining = lines.length - maxLines
      s"$shown\n... ($remaining more lines)"
