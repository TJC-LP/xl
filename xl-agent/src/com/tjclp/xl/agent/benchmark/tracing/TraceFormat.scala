package com.tjclp.xl.agent.benchmark.tracing

import com.tjclp.xl.agent.{AgentEvent, SubTurnUsage, TurnUsage}
import com.tjclp.xl.agent.benchmark.common.ModelPricing

import scala.collection.mutable

/** Formatters for conversation traces */
object TraceFormat:

  /** Convert a conversation trace to readable markdown */
  def toMarkdown(trace: ConversationTrace): String =
    val sb = new StringBuilder()
    val meta = trace.metadata
    val subTurnUsages = trace.subTurnUsages
    val apiTurnCount = trace.turnCount.max(1) // At least 1 API turn if we have events

    // Header with condensed metadata
    sb.append(s"# Task ${meta.taskId} - ${meta.skillName} approach (Case ${meta.caseNum})\n\n")

    // Condensed metadata line
    sb.append(s"**Started:** ${meta.startTime}\n")
    val durationStr = formatDuration(meta.durationMs)
    val tokenStr = meta.usage
      .map(u =>
        s"**Tokens:** ${formatNumber(u.totalTokens)} (${formatNumber(u.inputTokens)} in / ${formatNumber(u.outputTokens)} out)"
      )
      .getOrElse("")
    val costStr = meta.usage
      .map { u =>
        val pricing = ModelPricing.Opus45
        val cost =
          (u.inputTokens * pricing.inputPerMillion + u.outputTokens * pricing.outputPerMillion) / 1_000_000.0
        f"**Cost:** $$$cost%.2f"
      }
      .getOrElse("")

    sb.append(s"**Duration:** $durationStr | $tokenStr | $costStr\n")

    meta.passed.foreach { p =>
      val status = if p then "PASSED" else "FAILED"
      sb.append(s"**Result:** $status\n")
    }

    meta.error.foreach(e => sb.append(s"**Error:** $e\n"))

    sb.append("\n---\n\n")

    // User Message section (if prompts are available)
    trace.prompts.foreach { case (systemPrompt, userPrompt) =>
      sb.append("## User Message\n\n")
      sb.append("**System Prompt:**\n")
      sb.append("```\n")
      sb.append(truncate(systemPrompt, 3000))
      sb.append("\n```\n\n")
      sb.append("**User Prompt:**\n")
      sb.append("```\n")
      sb.append(truncate(userPrompt, 2000))
      sb.append("\n```\n\n")
      sb.append("---\n\n")
    }

    // Assistant Response section header
    val subTurnCountStr =
      if subTurnUsages.nonEmpty then s", ${subTurnUsages.size} sub-turns" else ""
    sb.append(s"## Assistant Response ($apiTurnCount API turn$subTurnCountStr)\n\n")

    // Build sub-turn timing map for inline display
    // SubTurnComplete events are emitted after tool results, so timing[n] covers sub-turn n
    val subTurnTimings = subTurnUsages.map(u => u.subTurnNum -> u).toMap
    var currentSubTurn = 1
    var subTurnStartMs = 0L
    var subTurnHeaderWritten = false

    // Process events grouped by sub-turn
    // A sub-turn = text (optional) + tool call + tool result
    // Sub-turn ends after tool result; new sub-turn starts with next text
    var currentText = new StringBuilder()

    def getSubTurnTimeRange(subTurnNum: Int): String =
      subTurnTimings.get(subTurnNum) match
        case Some(u) =>
          val endMs = subTurnStartMs + u.durationMs
          val range = f"[${subTurnStartMs / 1000.0}%.1fs-${endMs / 1000.0}%.1fs]"
          subTurnStartMs = endMs // Update for next sub-turn
          range
        case None => ""

    def writeSubTurnHeader(): Unit =
      if !subTurnHeaderWritten then
        val timeRange = getSubTurnTimeRange(currentSubTurn)
        sb.append(s"### Sub-turn $currentSubTurn $timeRange\n\n")
        subTurnHeaderWritten = true

    def flushText(): Unit =
      if currentText.nonEmpty then
        sb.append(currentText.toString.trim)
        sb.append("\n\n")
        currentText = new StringBuilder()

    trace.events.foreach { traced =>
      traced.event match
        case AgentEvent.Prompts(_, _) =>
          // Already handled in User Message section
          ()

        case AgentEvent.TextOutput(text, _) =>
          // Text starts or continues a sub-turn
          writeSubTurnHeader()
          currentText.append(text)

        case AgentEvent.ToolInvocation(name, _, _, command) =>
          // Tool call is part of the current sub-turn
          writeSubTurnHeader()
          flushText()
          sb.append(s"**Tool:** `$name`\n")
          command.foreach { cmd =>
            sb.append("```bash\n")
            sb.append(truncate(cmd, 2000))
            sb.append("\n```\n")
          }
          sb.append("\n")

        case AgentEvent.ToolResult(_, stdout, stderr, exitCode, files) =>
          // Tool result completes the sub-turn
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
          sb.append("\n")
          // Sub-turn complete, next text starts a new sub-turn
          currentSubTurn += 1
          subTurnHeaderWritten = false

        case AgentEvent.McpToolResult(_, content, isError) =>
          val header = if isError then "**MCP Error:**" else "**Result:**"
          sb.append(s"$header\n")
          if content.nonEmpty then
            sb.append("```\n")
            sb.append(truncate(content, 2000))
            sb.append("\n```\n")
          sb.append("\n")
          currentSubTurn += 1
          subTurnHeaderWritten = false

        case AgentEvent.ViewResult(content) =>
          sb.append("**View Result:**\n")
          if content.nonEmpty then
            sb.append("```markdown\n")
            sb.append(truncate(content, 2000))
            sb.append("\n```\n")
          sb.append("\n")
          currentSubTurn += 1
          subTurnHeaderWritten = false

        case AgentEvent.FileCreated(fileId, filename) =>
          sb.append(s"**File Created:** $filename (ID: $fileId)\n\n")

        case AgentEvent.Error(message) =>
          sb.append(s"**ERROR:** $message\n\n")

        case AgentEvent.TurnComplete(_) =>
          // API turn complete - don't need to display separately
          ()

        case AgentEvent.SubTurnComplete(_) =>
          // Sub-turn timing is used via subTurnTimings map
          ()
    }

    // Flush any remaining text as final sub-turn (text with no tool call)
    if currentText.nonEmpty then flushText()

    sb.append("---\n\n")

    // Summary section
    sb.append("## Summary\n\n")
    sb.append("| Metric | Value |\n")
    sb.append("|--------|-------|\n")
    sb.append(s"| API Turns | $apiTurnCount |\n")
    if subTurnUsages.nonEmpty then sb.append(s"| Sub-turns | ${subTurnUsages.size} |\n")
    sb.append(s"| Tool Calls | ${trace.toolCallCount} |\n")
    meta.usage.foreach { u =>
      sb.append(
        s"| Total Tokens | ${formatNumber(u.totalTokens)} (${formatNumber(u.inputTokens)} in / ${formatNumber(u.outputTokens)} out) |\n"
      )
    }
    if trace.errorCount > 0 then sb.append(s"| Errors | ${trace.errorCount} |\n")

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
