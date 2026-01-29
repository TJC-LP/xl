package com.tjclp.xl.agent.benchmark.tracing

import com.tjclp.xl.agent.AgentEvent
import com.tjclp.xl.agent.benchmark.common.ModelPricing

/** Formatters for conversation traces */
object TraceFormat:

  /** Convert a conversation trace to readable markdown */
  def toMarkdown(trace: ConversationTrace): String =
    val sb = new StringBuilder()
    val meta = trace.metadata

    // Header
    sb.append(s"# Task ${meta.taskId} - ${meta.skillName} approach (Case ${meta.caseNum})\n\n")

    // Metadata
    sb.append(s"**Started:** ${meta.startTime}\n")
    meta.endTime.foreach(e => sb.append(s"**Ended:** $e\n"))
    sb.append(s"**Duration:** ${formatDuration(meta.durationMs)}\n")

    meta.usage.foreach { u =>
      sb.append(
        s"**Tokens:** ${formatNumber(u.inputTokens)} input / ${formatNumber(u.outputTokens)} output\n"
      )
      val pricing = ModelPricing.Opus4
      val cost =
        (u.inputTokens * pricing.inputPerMillion + u.outputTokens * pricing.outputPerMillion) / 1_000_000.0
      sb.append(f"**Cost:** $$$cost%.2f\n")
    }

    meta.passed.foreach { p =>
      val status = if p then "PASSED" else "FAILED"
      sb.append(s"**Result:** $status\n")
    }

    meta.error.foreach(e => sb.append(s"**Error:** $e\n"))

    sb.append("\n---\n\n")

    // Events
    var turnNum = 0
    var currentText = new StringBuilder()

    def flushText(): Unit =
      if currentText.nonEmpty then
        turnNum += 1
        sb.append(s"## Turn $turnNum\n\n")
        sb.append("**Assistant:**\n")
        sb.append(currentText.toString.trim)
        sb.append("\n\n---\n\n")
        currentText = new StringBuilder()

    trace.events.foreach { traced =>
      traced.event match
        case AgentEvent.TextOutput(text, _) =>
          currentText.append(text)

        case AgentEvent.ToolInvocation(name, toolUseId, _, command) =>
          flushText()
          turnNum += 1
          val timeStr = f"[${traced.relativeMs / 1000.0}%.2fs]"
          sb.append(s"## Turn $turnNum $timeStr\n\n")
          sb.append(s"**Tool Call:** `$name`\n")
          command.foreach { cmd =>
            sb.append("```bash\n")
            sb.append(truncate(cmd, 2000))
            sb.append("\n```\n")
          }
          sb.append("\n")

        case AgentEvent.ToolResult(_, stdout, stderr, exitCode, files) =>
          val exitStr = exitCode.map(c => s" (exit: $c)").getOrElse("")
          sb.append(s"**Result:**$exitStr\n")
          if stdout.nonEmpty then
            sb.append("```\n")
            sb.append(truncate(stdout, 1000))
            sb.append("\n```\n")
          if stderr.nonEmpty then
            sb.append("**STDERR:**\n```\n")
            sb.append(truncate(stderr, 500))
            sb.append("\n```\n")
          if files.nonEmpty then sb.append(s"**Files:** ${files.mkString(", ")}\n")
          sb.append("\n---\n\n")

        case AgentEvent.FileCreated(fileId, filename) =>
          sb.append(s"**File Created:** $filename (ID: $fileId)\n\n")

        case AgentEvent.Error(message) =>
          sb.append(s"**ERROR:** $message\n\n")

        case AgentEvent.McpToolResult(toolUseId, content, isError) =>
          val header = if isError then "**MCP Error:**" else "**MCP Result:**"
          sb.append(s"$header\n")
          if content.nonEmpty then
            sb.append("```\n")
            sb.append(truncate(content, 2000))
            sb.append("\n```\n")
          sb.append("\n---\n\n")

        case AgentEvent.ViewResult(content) =>
          sb.append("**View Result:**\n")
          if content.nonEmpty then
            sb.append("```markdown\n")
            sb.append(truncate(content, 2000))
            sb.append("\n```\n")
          sb.append("\n---\n\n")
    }

    // Flush any remaining text
    flushText()

    // Summary
    sb.append("## Summary\n\n")
    sb.append("| Metric | Value |\n")
    sb.append("|--------|-------|\n")
    sb.append(s"| Turns | $turnNum |\n")
    sb.append(s"| Tool Calls | ${trace.toolCallCount} |\n")
    sb.append(s"| Errors | ${trace.errorCount} |\n")

    // Extract xl commands if present
    val xlCommands = trace.events
      .collect {
        case TracedEvent(_, _, AgentEvent.ToolInvocation(_, _, _, Some(cmd)))
            if cmd.contains("xl") =>
          extractXlCommand(cmd)
      }
      .flatten
      .distinct

    if xlCommands.nonEmpty then sb.append(s"| xl commands | ${xlCommands.mkString(", ")} |\n")

    sb.toString

  private def formatDuration(ms: Long): String =
    if ms < 1000 then s"${ms}ms"
    else f"${ms / 1000.0}%.1fs"

  private def formatNumber(n: Long): String =
    if n >= 1_000_000 then f"${n / 1_000_000.0}%.1fM"
    else if n >= 1_000 then f"${n / 1_000.0}%.1fK"
    else n.toString

  private def truncate(s: String, maxLen: Int): String =
    if s.length <= maxLen then s
    else s.take(maxLen) + s"\n... (${s.length - maxLen} more chars)"

  private def extractXlCommand(cmd: String): Option[String] =
    // Extract xl subcommand from bash command
    val patterns = List(
      """\$XL\s+-f\s+\S+\s+(-o\s+\S+\s+)?(\w+)""".r,
      """xl\s+-f\s+\S+\s+(-o\s+\S+\s+)?(\w+)""".r,
      """\.\/xl[^\s]*\s+-f\s+\S+\s+(-o\s+\S+\s+)?(\w+)""".r
    )
    patterns.flatMap(_.findFirstMatchIn(cmd).map(_.group(2))).headOption
