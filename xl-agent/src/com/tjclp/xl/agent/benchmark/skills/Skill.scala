package com.tjclp.xl.agent.benchmark.skills

import cats.effect.IO
import com.tjclp.xl.agent.{Agent, AgentConfig, AgentTask}
import com.tjclp.xl.agent.anthropic.AnthropicClientIO

/** Context created during skill setup, containing resources to clean up */
case class SkillContext(
  fileIds: List[String] = Nil,
  skillId: Option[String] = None,
  metadata: Map[String, String] = Map.empty
)

object SkillContext:
  val empty: SkillContext = SkillContext()

/**
 * Skill abstraction for different benchmark approaches
 *
 * Skills encapsulate:
 *   - Setup: Upload files, register APIs, etc.
 *   - Teardown: Clean up uploaded resources
 *   - Agent creation: Configure agent with skill-specific prompts and tools
 *
 * This allows N-skill comparisons without hardcoding approaches.
 */
trait Skill:
  /** Unique identifier for this skill (e.g., "xl", "xlsx", "raw-python") */
  def name: String

  /** Human-readable name (e.g., "xl-cli", "OpenPyXL", "Raw Python") */
  def displayName: String

  /** One-time setup (upload files, register skill via API, etc.) */
  def setup(client: AnthropicClientIO, config: AgentConfig): IO[SkillContext]

  /** Cleanup resources (delete uploaded files, etc.) */
  def teardown(client: AnthropicClientIO, ctx: SkillContext): IO[Unit]

  /** Create an agent configured for this skill */
  def createAgent(
    client: AnthropicClientIO,
    ctx: SkillContext,
    config: AgentConfig
  ): Agent

  /** Description for --list-skills output */
  def description: String = displayName
