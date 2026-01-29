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
    s"""You have access to the xl CLI tool for Excel operations.

AVAILABLE FILES:
- ${binaryFile.filename}: The xl CLI binary (chmod +x first)
- Input xlsx file in $$INPUT_DIR

xl CLI COMMANDS:
- $$XL -f $$INPUT sheets                    # List sheets
- $$XL -f $$INPUT -s "Sheet" view A1:B10    # View range
- $$XL -f $$INPUT -o $$OUTPUT put A1 value  # Write value
- $$XL -f $$INPUT -o $$OUTPUT putf A1 "=SUM(B:B)"  # Write formula
- echo '[{"op":"put","ref":"A1","value":123}]' | $$XL -f $$INPUT -o $$OUTPUT batch -

Use -o $$OUTPUT_DIR/output.xlsx or save to /tmp then copy."""

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
