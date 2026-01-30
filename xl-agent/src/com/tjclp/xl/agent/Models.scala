package com.tjclp.xl.agent

import io.circe.*
import io.circe.generic.semiauto.*
import java.nio.file.Path

// ============================================================================
// Core Agent Types
// ============================================================================

/** Configuration for the agent */
case class AgentConfig(
  model: String = "claude-opus-4-5-20251101",
  maxTokens: Int = 8192,
  verbose: Boolean = false,
  xlBinaryPath: Option[Path] = None,
  xlSkillPath: Option[Path] = None
)

/** Token usage from API */
case class TokenUsage(
  inputTokens: Long,
  outputTokens: Long
):
  def totalTokens: Long = inputTokens + outputTokens
  def +(other: TokenUsage): TokenUsage =
    TokenUsage(inputTokens + other.inputTokens, outputTokens + other.outputTokens)

object TokenUsage:
  val zero: TokenUsage = TokenUsage(0, 0)
  given Encoder[TokenUsage] = deriveEncoder
  given Decoder[TokenUsage] = deriveDecoder

/** Per-turn token usage tracking */
case class TurnUsage(
  turnNum: Int,
  inputTokens: Long,
  outputTokens: Long,
  cumulativeInputTokens: Long,
  cumulativeOutputTokens: Long,
  durationMs: Long
):
  def totalTokens: Long = inputTokens + outputTokens
  def cumulativeTotalTokens: Long = cumulativeInputTokens + cumulativeOutputTokens

object TurnUsage:
  given Encoder[TurnUsage] = deriveEncoder
  given Decoder[TurnUsage] = deriveDecoder

/**
 * Sub-turn tracking for code execution.
 *
 * Code execution runs as a single API message but has multiple internal "sub-turns" (assistant text
 * → tool call → result cycles). The API only reports total tokens at the end, so we track timing
 * per sub-turn but can't split tokens.
 */
case class SubTurnUsage(
  subTurnNum: Int,
  durationMs: Long,
  hasToolCall: Boolean
)

object SubTurnUsage:
  given Encoder[SubTurnUsage] = deriveEncoder
  given Decoder[SubTurnUsage] = deriveDecoder

/** Events emitted during agent execution */
enum AgentEvent:
  /** Emitted at the start with the system and user prompts for tracing */
  case Prompts(systemPrompt: String, userPrompt: String)

  case TextOutput(text: String, contentIndex: Int = 0)
  case ToolInvocation(
    name: String,
    toolUseId: String,
    input: Json,
    command: Option[String] = None
  )
  case ToolResult(
    toolUseId: String,
    stdout: String,
    stderr: String,
    exitCode: Option[Int] = None,
    files: List[String] = Nil
  )

  /** MCP/skill tool result (e.g., skill documentation from view command) */
  case McpToolResult(toolUseId: String, content: String, isError: Boolean = false)

  /** Text editor view result */
  case ViewResult(content: String)
  case FileCreated(fileId: String, filename: String)
  case Error(message: String)

  /** Emitted when a turn (assistant response cycle) completes */
  case TurnComplete(usage: TurnUsage)

  /**
   * Emitted when a sub-turn completes within code execution.
   *
   * A sub-turn is one assistant-text → tool-call → result cycle. Multiple sub-turns can occur
   * within a single API message during code execution.
   */
  case SubTurnComplete(usage: SubTurnUsage)

object AgentEvent:
  import io.circe.syntax.*

  given Encoder[AgentEvent] = Encoder.instance {
    case Prompts(systemPrompt, userPrompt) =>
      Json.obj(
        "type" -> "Prompts".asJson,
        "systemPrompt" -> systemPrompt.asJson,
        "userPrompt" -> userPrompt.asJson
      )
    case TextOutput(text, idx) =>
      Json.obj(
        "type" -> "TextOutput".asJson,
        "text" -> text.asJson,
        "contentIndex" -> idx.asJson
      )
    case ToolInvocation(name, toolUseId, input, command) =>
      Json.obj(
        "type" -> "ToolInvocation".asJson,
        "name" -> name.asJson,
        "toolUseId" -> toolUseId.asJson,
        "input" -> input,
        "command" -> command.asJson
      )
    case ToolResult(toolUseId, stdout, stderr, exitCode, files) =>
      Json.obj(
        "type" -> "ToolResult".asJson,
        "toolUseId" -> toolUseId.asJson,
        "stdout" -> stdout.asJson,
        "stderr" -> stderr.asJson,
        "exitCode" -> exitCode.asJson,
        "files" -> files.asJson
      )
    case McpToolResult(toolUseId, content, isError) =>
      Json.obj(
        "type" -> "McpToolResult".asJson,
        "toolUseId" -> toolUseId.asJson,
        "content" -> content.asJson,
        "isError" -> isError.asJson
      )
    case ViewResult(content) =>
      Json.obj(
        "type" -> "ViewResult".asJson,
        "content" -> content.asJson
      )
    case FileCreated(fileId, filename) =>
      Json.obj(
        "type" -> "FileCreated".asJson,
        "fileId" -> fileId.asJson,
        "filename" -> filename.asJson
      )
    case Error(message) =>
      Json.obj(
        "type" -> "Error".asJson,
        "message" -> message.asJson
      )
    case TurnComplete(usage) =>
      Json.obj(
        "type" -> "TurnComplete".asJson,
        "turnNum" -> usage.turnNum.asJson,
        "inputTokens" -> usage.inputTokens.asJson,
        "outputTokens" -> usage.outputTokens.asJson,
        "cumulativeInputTokens" -> usage.cumulativeInputTokens.asJson,
        "cumulativeOutputTokens" -> usage.cumulativeOutputTokens.asJson,
        "durationMs" -> usage.durationMs.asJson
      )
    case SubTurnComplete(usage) =>
      Json.obj(
        "type" -> "SubTurnComplete".asJson,
        "subTurnNum" -> usage.subTurnNum.asJson,
        "durationMs" -> usage.durationMs.asJson,
        "hasToolCall" -> usage.hasToolCall.asJson
      )
  }

/** Result of an agent execution */
case class AgentResult(
  success: Boolean,
  outputFileId: Option[String],
  outputPath: Option[Path],
  usage: TokenUsage,
  latencyMs: Long,
  transcript: Vector[AgentEvent],
  responseText: Option[String] = None,
  error: Option[String] = None
)

object AgentResult:
  given Encoder[AgentResult] = Encoder.instance { r =>
    Json.obj(
      "success" -> Json.fromBoolean(r.success),
      "outputFileId" -> r.outputFileId.fold(Json.Null)(Json.fromString),
      "outputPath" -> r.outputPath.fold(Json.Null)(p => Json.fromString(p.toString)),
      "usage" -> Encoder[TokenUsage].apply(r.usage),
      "latencyMs" -> Json.fromLong(r.latencyMs),
      "responseText" -> r.responseText.fold(Json.Null)(Json.fromString),
      "error" -> r.error.fold(Json.Null)(Json.fromString)
    )
  }

// ============================================================================
// Uploaded File Reference
// ============================================================================

/** Reference to an uploaded file in the Anthropic Files API */
case class UploadedFile(
  id: String,
  filename: String
)
