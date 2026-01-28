package com.tjclp.xl.agent.error

import scala.util.control.NoStackTrace

/** Errors that can occur during agent execution */
enum AgentError extends Exception with NoStackTrace:
  case ApiKeyMissing(message: String)
  case FileUploadFailed(path: String, cause: String)
  case FileDownloadFailed(fileId: String, cause: String)
  case RequestFailed(cause: String)
  case StreamingError(cause: String)
  case OutputNotFound(message: String)
  case EvaluationFailed(cause: String)
  case ConfigError(message: String)
  case TaskLoadError(cause: String)
  case ParseError(json: String, cause: String)

  override def getMessage: String = this match
    case ApiKeyMissing(msg) => s"API key missing: $msg"
    case FileUploadFailed(path, c) => s"Failed to upload file '$path': $c"
    case FileDownloadFailed(fid, c) => s"Failed to download file '$fid': $c"
    case RequestFailed(c) => s"API request failed: $c"
    case StreamingError(c) => s"Streaming error: $c"
    case OutputNotFound(msg) => s"Output not found: $msg"
    case EvaluationFailed(c) => s"Evaluation failed: $c"
    case ConfigError(msg) => s"Configuration error: $msg"
    case TaskLoadError(c) => s"Failed to load tasks: $c"
    case ParseError(json, c) => s"JSON parse error: $c (json: ${json.take(100)})"
