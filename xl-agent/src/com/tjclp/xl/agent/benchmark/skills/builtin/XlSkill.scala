package com.tjclp.xl.agent.benchmark.skills.builtin

import cats.effect.IO
import cats.syntax.all.*
import com.anthropic.models.beta.AnthropicBeta
import com.anthropic.models.beta.messages.{
  BetaCodeExecutionTool20250825,
  BetaContainerParams,
  BetaSkillParams,
  MessageCreateParams
}
import com.tjclp.xl.agent.{Agent, AgentConfig, AgentTask, UploadedFile}
import com.tjclp.xl.agent.anthropic.{AnthropicClientIO, SkillsApi}
import com.tjclp.xl.agent.approach.ApproachStrategy
import com.tjclp.xl.agent.benchmark.common.FileManager
import com.tjclp.xl.agent.benchmark.skills.{Skill, SkillContext, SkillRegistry}

import java.nio.file.Path

/** Skill using the custom xl-cli tool with the Skills API */
object XlSkill extends Skill:

  override val name: String = "xl"
  override val displayName: String = "xl-cli"
  override val description: String = "Custom xl-cli tool for Excel operations"

  override def setup(client: AnthropicClientIO, config: AgentConfig): IO[SkillContext] =
    for
      // Resolve paths
      binaryPath <- FileManager.resolveBinaryPath(config.xlBinaryPath)
      skillPath <- FileManager.resolveSkillPath(config.xlSkillPath)

      _ <- IO.println(s"  [$name] Uploading binary: ${binaryPath.getFileName}")
      binaryFile <- client.uploadFile(binaryPath)

      _ <- IO.println(s"  [$name] Registering skill via Skills API...")
      apiKey <- AnthropicClientIO.loadApiKey
      skillId <- SkillsApi.getOrCreateXlSkill(apiKey, skillPath)
      _ <- IO.println(s"  [$name] Skill ID: $skillId")
    yield SkillContext(
      fileIds = List(binaryFile.id),
      skillId = Some(skillId),
      metadata = Map(
        "binaryFilename" -> binaryFile.filename,
        "binaryFileId" -> binaryFile.id
      )
    )

  override def teardown(client: AnthropicClientIO, ctx: SkillContext): IO[Unit] =
    ctx.fileIds.traverse_(id => client.deleteFile(id).attempt.void)

  override def createAgent(
    client: AnthropicClientIO,
    ctx: SkillContext,
    config: AgentConfig
  ): Agent =
    val strategy = XlSkillStrategy(
      binaryFilename = ctx.metadata.getOrElse("binaryFilename", "xl"),
      binaryFileId = ctx.metadata.getOrElse("binaryFileId", ""),
      skillId = ctx.skillId.getOrElse("")
    )
    Agent.create(client, config, strategy)

/** Internal strategy for XlSkill */
private class XlSkillStrategy(
  binaryFilename: String,
  binaryFileId: String,
  skillId: String
) extends ApproachStrategy:

  override def name: String = "xl"

  override def containerUploads(inputFileId: String): List[String] =
    List(binaryFileId, inputFileId)

  override def systemPrompt: String =
    s"""You are a spreadsheet expert assistant. You have access to the xl CLI tool for manipulating Excel files.

IMPORTANT REQUIREMENTS:
1. Use the xl CLI tool for ALL Excel operations (not Python libraries like openpyxl)
2. ALWAYS save output to $$OUTPUT_DIR/output.xlsx using the -o flag
3. When you finish modifying the spreadsheet, the final file MUST be at $$OUTPUT_DIR/output.xlsx
4. DO NOT modify the input data - only add formulas or values where specified

CRITICAL: OUTPUT_DIR CHANGES BETWEEN TOOL CALLS
The $$OUTPUT_DIR path changes with each tool invocation. To work around this:
- Save to /tmp/output.xlsx for intermediate work and verification
- In your FINAL tool call, copy to $$OUTPUT_DIR: cp /tmp/output.xlsx "$$OUTPUT_DIR/output.xlsx"
- Or do all work (write + verify + save) in a SINGLE tool call

AVAILABLE FILES:
- $binaryFilename: The xl CLI binary (chmod +x it first, then use ./$binaryFilename or move to PATH)
- Input xlsx file: The spreadsheet to work with

The xl-cli skill documentation is available. Use xl commands to complete the task.

Your goal is to solve spreadsheet tasks efficiently and accurately. Always:
1. Explore the input file first to understand its structure
2. Plan your solution before executing
3. Use xl CLI commands with -o for writing (save to /tmp first, then copy to $$OUTPUT_DIR)
4. Verify your solution by viewing the output cells
5. NEVER modify input data cells - only add to the cells specified in the task

The xl CLI supports viewing data, writing values, writing formulas, and batch operations."""

  override def userPrompt(task: AgentTask, inputFilename: String): String =
    val answerSection = task.answerPosition
      .map(pos => s"\nThe answer will be evaluated at position: $pos\n")
      .getOrElse("")

    s"""You are a spreadsheet expert with access to the xl CLI for Excel operations.

## Task
${task.instruction}

## File Locations
- Input file: $$INPUT_DIR/$inputFilename
- XL binary: $$INPUT_DIR/$binaryFilename
$answerSection
## CRITICAL: Output Location
You MUST save your final output to: $$OUTPUT_DIR/output.xlsx
The $$OUTPUT_DIR environment variable is set by the system. Using it ensures the file is properly captured.

## Setup Commands (run these first)
```bash
chmod +x $$INPUT_DIR/$binaryFilename
XL="$$INPUT_DIR/$binaryFilename"
INPUT="$$INPUT_DIR/$inputFilename"
OUTPUT="$$OUTPUT_DIR/output.xlsx"
```

## xl CLI Commands
```bash
# List sheets
$$XL -f $$INPUT sheets

# View data (use quotes for sheet names with spaces)
$$XL -f $$INPUT -s "<sheet>" view <range>

# Write a value (always use -o $$OUTPUT to save)
$$XL -f $$INPUT -o $$OUTPUT put <ref> <value>

# Write a formula (always use -o $$OUTPUT to save)
$$XL -f $$INPUT -o $$OUTPUT putf <ref> "=FORMULA"

# Batch operations (always use -o $$OUTPUT to save)
echo '[{"op":"put","ref":"A1","value":"Hello"}]' | $$XL -f $$INPUT -o $$OUTPUT batch -
```

## Instructions
1. Explore the input file to understand its structure
2. Implement the solution using xl CLI commands
3. IMPORTANT: Always use -o $$OUTPUT (which is $$OUTPUT_DIR/output.xlsx) when writing
4. Verify your solution by viewing the output cells${task.answerPosition
        .map(p => s" at $p")
        .getOrElse("")}
"""

  override def configureRequest(builder: MessageCreateParams.Builder): MessageCreateParams.Builder =
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
