package com.tjclp.xl.agent.approach

import com.anthropic.models.beta.messages.MessageCreateParams
import com.tjclp.xl.agent.AgentTask

/** Strategy for different benchmark approaches (xl-cli vs xlsx skill) */
trait ApproachStrategy:
  /** Name of this approach for logging/reporting */
  def name: String

  /** File IDs to include as container_upload blocks */
  def containerUploads(inputFileId: String): List[String]

  /** Build the system prompt */
  def systemPrompt: String

  /** Build the user prompt for a task */
  def userPrompt(task: AgentTask, inputFilename: String): String

  /** Configure the MessageCreateParams with tools/skills/betas */
  def configureRequest(builder: MessageCreateParams.Builder): MessageCreateParams.Builder
