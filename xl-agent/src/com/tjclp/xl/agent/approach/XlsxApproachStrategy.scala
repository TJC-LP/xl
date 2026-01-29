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

  override def systemPrompt: String =
    """You are a spreadsheet expert assistant. You have access to the xlsx skill for manipulating Excel files using Python and openpyxl.

IMPORTANT REQUIREMENTS:
1. Use Python with openpyxl for ALL Excel operations
2. ALWAYS save output to $OUTPUT_DIR/output.xlsx
3. When you finish modifying the spreadsheet, the final file MUST be at $OUTPUT_DIR/output.xlsx
4. DO NOT modify the input data - only add formulas or values where specified

Your goal is to solve spreadsheet tasks efficiently and accurately. Always:
1. Explore the input file first to understand its structure
2. Plan your solution before executing
3. Save your output to $OUTPUT_DIR/output.xlsx
4. Verify your solution by reading back the output cells
5. NEVER modify input data cells - only add to the cells specified in the task"""

  override def userPrompt(task: AgentTask, inputFilename: String): String =
    val answerSection = task.answerPosition
      .map(pos => s"\nThe answer will be evaluated at position: $pos\n")
      .getOrElse("")

    s"""You are a spreadsheet expert with access to Python and openpyxl for Excel operations.

## Task
${task.instruction}

## File Locations
- Input file: $$INPUT_DIR/$inputFilename
$answerSection
## CRITICAL: Output Location
You MUST save your final output to: $$OUTPUT_DIR/output.xlsx
The $$OUTPUT_DIR environment variable is set by the system. Using it ensures the file is properly captured.

## Python Setup
```python
import os
from openpyxl import load_workbook

# Load input file
INPUT = os.path.join(os.environ['INPUT_DIR'], '$inputFilename')
OUTPUT = os.path.join(os.environ['OUTPUT_DIR'], 'output.xlsx')

wb = load_workbook(INPUT)
ws = wb.active  # or wb['Sheet1'] for specific sheet
```

## Instructions
1. Explore the input file to understand its structure
2. Implement the solution using openpyxl
3. IMPORTANT: Always save to OUTPUT (which is $$OUTPUT_DIR/output.xlsx)
4. Verify your solution by reading the output cells${task.answerPosition
        .map(p => s" at $p")
        .getOrElse("")}
"""

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
