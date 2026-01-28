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

/** Events emitted during agent execution */
enum AgentEvent:
  case TextOutput(text: String)
  case ToolInvocation(name: String, command: String)
  case ToolResult(stdout: String, stderr: String, exitCode: Option[Int] = None)
  case FileCreated(fileId: String, filename: String)
  case Error(message: String)

object AgentEvent:
  given Encoder[AgentEvent] = Encoder.instance {
    case TextOutput(text) =>
      Json.obj("type" -> Json.fromString("text"), "text" -> Json.fromString(text))
    case ToolInvocation(name, command) =>
      Json.obj(
        "type" -> Json.fromString("tool_invocation"),
        "name" -> Json.fromString(name),
        "command" -> Json.fromString(command)
      )
    case ToolResult(stdout, stderr, exitCode) =>
      Json.obj(
        "type" -> Json.fromString("tool_result"),
        "stdout" -> Json.fromString(stdout),
        "stderr" -> Json.fromString(stderr),
        "exitCode" -> exitCode.fold(Json.Null)(Json.fromInt)
      )
    case FileCreated(fileId, filename) =>
      Json.obj(
        "type" -> Json.fromString("file_created"),
        "fileId" -> Json.fromString(fileId),
        "filename" -> Json.fromString(filename)
      )
    case Error(message) =>
      Json.obj("type" -> Json.fromString("error"), "message" -> Json.fromString(message))
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
