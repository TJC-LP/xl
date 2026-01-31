package com.tjclp.xl.agent.approach

import com.anthropic.models.beta.AnthropicBeta
import com.anthropic.models.beta.messages.{
  BetaCodeExecutionTool20250825,
  BetaContainerParams,
  BetaSkillParams,
  MessageCreateParams
}
import com.tjclp.xl.agent.AgentTask

/** Approach using Anthropic's built-in xlsx skill (openpyxl) */
class XlsxApproachStrategy extends ApproachStrategy:

  override def name: String = "xlsx"

  override def containerUploads(inputFileId: String): List[String] =
    List(inputFileId)

  override def toolSystemPrompt: String =
    "You have Python with openpyxl for Excel operations. Use the xlsx skill documentation for reference."

  override def toolUserPrompt(task: AgentTask, inputFilename: String): String =
    val answerSection = task.answerPosition
      .map(pos => s"\nThe answer will be evaluated at position: $pos")
      .getOrElse("")

    s"""## Task
${task.instruction}
$answerSection

## File Locations
- Input: $$INPUT_DIR/$inputFilename
- Output: $$OUTPUT_DIR/output.xlsx (or /tmp/output.xlsx then copy)

## Setup
```python
import os, shutil
from openpyxl import load_workbook

INPUT = os.path.join(os.environ['INPUT_DIR'], '$inputFilename')
wb = load_workbook(INPUT)
ws = wb.active
```"""

  override def configureRequest(builder: MessageCreateParams.Builder): MessageCreateParams.Builder =
    val xlsxSkill = BetaSkillParams
      .builder()
      .skillId("xlsx")
      .`type`(BetaSkillParams.Type.ANTHROPIC)
      .version("latest")
      .build()

    val container = BetaContainerParams
      .builder()
      .addSkill(xlsxSkill)
      .build()

    builder
      .container(container)
      .addTool(BetaCodeExecutionTool20250825.builder().build())
      .addBeta(AnthropicBeta.of("code-execution-2025-08-25"))
      .addBeta(AnthropicBeta.SKILLS_2025_10_02)
      .addBeta(AnthropicBeta.FILES_API_2025_04_14)
