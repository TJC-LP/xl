package com.tjclp.xl.agent.approach

import com.anthropic.models.beta.messages.MessageCreateParams
import com.tjclp.xl.agent.AgentTask
import com.tjclp.xl.agent.benchmark.prompt.BenchmarkPrompts

/**
 * Strategy for different benchmark approaches (xl-cli vs xlsx skill).
 *
 * Prompts are composed from:
 *   - BenchmarkPrompts.systemPreamble/taskFooter (behavioral guidance)
 *   - toolSystemPrompt/toolUserPrompt (tool-specific instructions)
 */
trait ApproachStrategy:
  /** Name of this approach for logging/reporting */
  def name: String

  /** File IDs to include as container_upload blocks */
  def containerUploads(inputFileId: String): List[String]

  /** Tool-specific system prompt (no behavioral guidance - that's in preamble) */
  def toolSystemPrompt: String

  /** Tool-specific user prompt (setup commands, examples) */
  def toolUserPrompt(task: AgentTask, inputFilename: String): String

  /** Build the full system prompt with behavioral preamble */
  final def systemPrompt: String =
    s"${BenchmarkPrompts.systemPreamble}\n\n${toolSystemPrompt}"

  /** Build the full user prompt with task footer */
  final def userPrompt(task: AgentTask, inputFilename: String): String =
    s"${toolUserPrompt(task, inputFilename)}\n\n${BenchmarkPrompts.taskFooter}"

  /** Configure the MessageCreateParams with tools/skills/betas */
  def configureRequest(builder: MessageCreateParams.Builder): MessageCreateParams.Builder
