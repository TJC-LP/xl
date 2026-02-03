package com.tjclp.xl.agent.approach

import com.anthropic.models.beta.AnthropicBeta
import com.anthropic.models.beta.messages.{
  BetaCodeExecutionTool20250825,
  BetaContainerParams,
  BetaSkillParams,
  MessageCreateParams
}
import com.tjclp.xl.agent.{AgentTask, UploadedFile}

/** Approach using the custom xl-cli skill with uploaded binary and Skills API */
class XlApproachStrategy(
  binaryFile: UploadedFile,
  skillId: String
) extends ApproachStrategy:

  override def name: String = "xl"

  // Only binary + input file (skill is registered via Skills API)
  override def containerUploads(inputFileId: String): List[String] =
    List(binaryFile.id, inputFileId)

  override def toolSystemPrompt: String =
    "You have the xl CLI for Excel operations. Use the xl-cli skill documentation for command reference."

  override def toolUserPrompt(task: AgentTask, inputFilename: String): String =
    val answerSection = task.answerPosition
      .map(pos => s"\nThe answer will be evaluated at position: $pos")
      .getOrElse("")

    s"""## Task
${task.instruction}
$answerSection

## File Locations
- Input: $$INPUT_DIR/$inputFilename
- XL binary: $$INPUT_DIR/${binaryFile.filename}
- Output: $$OUTPUT_DIR/output.xlsx (or /tmp/output.xlsx then copy)

## Setup
```bash
chmod +x $$INPUT_DIR/${binaryFile.filename}
XL="$$INPUT_DIR/${binaryFile.filename}"
INPUT="$$INPUT_DIR/$inputFilename"
```"""

  override def configureRequest(builder: MessageCreateParams.Builder): MessageCreateParams.Builder =
    // Build custom skill container (xl-cli registered via Skills API)
    val xlSkill = BetaSkillParams
      .builder()
      .skillId(skillId)
      .`type`(BetaSkillParams.Type.CUSTOM)
      .version("latest")
      .build()

    val container = BetaContainerParams
      .builder()
      .addSkill(xlSkill)
      .build()

    builder
      .container(container)
      .addTool(BetaCodeExecutionTool20250825.builder().build())
      .addBeta(AnthropicBeta.of("code-execution-2025-08-25"))
      .addBeta(AnthropicBeta.SKILLS_2025_10_02)
      .addBeta(AnthropicBeta.FILES_API_2025_04_14)
