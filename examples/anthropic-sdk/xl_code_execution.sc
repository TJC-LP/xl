#!/usr/bin/env -S scala-cli shebang
//> using scala 3.7.4
//> using dep com.anthropic:anthropic-java:2.11.1
//> using dep io.github.cdimascio:dotenv-java:3.2.0

/**
 * XL CLI + Anthropic Code Execution Example
 *
 * This script demonstrates using the Anthropic API's code execution tool
 * with the XL CLI for Excel operations. The code execution sandbox has
 * no network access, so we upload the pre-built Linux binary.
 *
 * Prerequisites:
 *   1. Copy .env.example to .env and add your API key
 *   2. Download the Linux binary:
 *      gh release download --repo TJC-LP/xl --pattern "xl-*-linux-amd64" -D examples/anthropic-sdk
 *      mv examples/anthropic-sdk/xl-*-linux-amd64 examples/anthropic-sdk/xl-linux-amd64
 *
 * Usage:
 *   scala-cli run examples/anthropic-sdk/xl_code_execution.sc
 */

import com.anthropic.client.AnthropicClient
import com.anthropic.client.okhttp.AnthropicOkHttpClient
import com.anthropic.core.JsonValue
import com.anthropic.models.messages.*
import com.anthropic.models.beta.AnthropicBeta
import com.anthropic.models.beta.files.{FileMetadata, FileUploadParams}
import io.github.cdimascio.dotenv.Dotenv

import java.nio.file.{Files, Path, Paths}
import scala.jdk.CollectionConverters.*
import scala.util.Try

// ============================================================================
// Configuration
// ============================================================================

val scriptDir = Paths.get("examples/anthropic-sdk")
val binaryPath = scriptDir.resolve("xl-linux-amd64")
val samplePath = scriptDir.resolve("sample.xlsx")
val envPath = scriptDir.resolve(".env")

// ============================================================================
// Load Environment
// ============================================================================

// Load .env file if it exists
val dotenv = Dotenv.configure()
  .directory(scriptDir.toString)
  .ignoreIfMissing()
  .load()

// Set API key from .env if not already in environment
Option(dotenv.get("ANTHROPIC_API_KEY")).foreach { key =>
  if System.getenv("ANTHROPIC_API_KEY") == null then
    // Can't set env var, but the SDK client builder accepts it directly
    System.setProperty("ANTHROPIC_API_KEY", key)
}

val apiKey = Option(System.getenv("ANTHROPIC_API_KEY"))
  .orElse(Option(dotenv.get("ANTHROPIC_API_KEY")))
  .getOrElse {
    System.err.println("""
      |Error: ANTHROPIC_API_KEY not found
      |
      |Set it in examples/anthropic-sdk/.env or as an environment variable.
      |See .env.example for the format.
      |""".stripMargin)
    System.exit(1)
    ""
  }

// ============================================================================
// Skill Instructions
// ============================================================================

val skillInstructions = """
You have access to the `xl` CLI tool for Excel operations.

The xl binary has been uploaded and is available at /mnt/user/xl-linux-amd64
First, make it executable: chmod +x /mnt/user/xl-linux-amd64

Key commands:
- xl -f <file> sheets                    # List sheets
- xl -f <file> -s <sheet> bounds         # Get used range
- xl -f <file> -s <sheet> view <range>   # View as table
- xl -f <file> -s <sheet> stats <range>  # Calculate statistics
- xl -f <file> -s <sheet> view <range> --eval  # Show computed values
- xl -f <file> -s <sheet> view <range> --formulas  # Show formulas

The sample Excel file is at /mnt/user/sample.xlsx
""".trim

// ============================================================================
// Main
// ============================================================================

// Validate prerequisites
if !Files.exists(binaryPath) then
  System.err.println(s"""
    |Error: Linux binary not found at $binaryPath
    |
    |Download it with:
    |  gh release download --repo TJC-LP/xl --pattern "xl-*-linux-amd64" -D examples/anthropic-sdk
    |  mv examples/anthropic-sdk/xl-*-linux-amd64 examples/anthropic-sdk/xl-linux-amd64
    |""".stripMargin)
  System.exit(1)

if !Files.exists(samplePath) then
  System.err.println(s"""
    |Error: Sample file not found at $samplePath
    |
    |Generate it with:
    |  scala-cli run examples/anthropic-sdk/create_sample.sc
    |""".stripMargin)
  System.exit(1)

println("=== XL CLI + Anthropic Code Execution Example ===\n")

// Initialize client with API key
val client = AnthropicOkHttpClient.builder()
  .apiKey(apiKey)
  .build()

// Step 1: Upload the Linux binary
println("1. Uploading xl binary...")
val binaryFile = uploadFile(client, binaryPath)
println(s"   Uploaded: ${binaryFile.id()} (${Files.size(binaryPath)} bytes)")

// Step 2: Upload the sample Excel file
println("2. Uploading sample.xlsx...")
val excelFile = uploadFile(client, samplePath)
println(s"   Uploaded: ${excelFile.id()} (${Files.size(samplePath)} bytes)")

// Step 3: Create message with code execution
println("3. Sending request to Claude with code execution tool...")
println()

val userPrompt = s"""
Analyze the uploaded Excel file. Please:
1. List all sheets
2. View the Sales sheet data
3. Calculate statistics on the quarterly data (B2:E4)
4. Show the formulas in column F
5. Summarize your findings
""".trim

// Build message with code execution tool and container uploads
val params = MessageCreateParams.builder()
  .model(Model.CLAUDE_SONNET_4_5_20250929)
  .maxTokens(4096L)
  .system(skillInstructions)
  // Need to provide messages (will be overridden by additionalBodyProperty)
  .addUserMessage("placeholder")
  // Override messages to include container_upload blocks
  .putAdditionalBodyProperty("messages", JsonValue.from(java.util.List.of(
    java.util.Map.of(
      "role", "user",
      "content", java.util.List.of(
        java.util.Map.of("type", "text", "text", userPrompt),
        java.util.Map.of("type", "container_upload", "file_id", binaryFile.id()),
        java.util.Map.of("type", "container_upload", "file_id", excelFile.id())
      )
    )
  )))
  // Add code execution tool
  .putAdditionalBodyProperty("tools", JsonValue.from(java.util.List.of(
    java.util.Map.of(
      "type", "code_execution_20250825",
      "name", "code_execution"
    )
  )))
  // Add beta headers for code execution and files API
  .putAdditionalHeader("anthropic-beta", "code-execution-2025-08-25,files-api-2025-04-14")
  .build()

val response = client.messages().create(params)

// Step 4: Process the response
println("=== Claude's Response ===\n")

// Extract content from the response
// Note: server_tool_use and tool_result blocks are not exposed via the typed API
// but the text blocks contain Claude's analysis
response.content().asScala.foreach { block =>
  if block.isText then
    println(block.asText().text())
  // Tool use/results are handled internally by code execution
}

println("\n=== Done ===")
println(s"Stop reason: ${response.stopReason()}")
println(s"Usage: ${response.usage().inputTokens()} input, ${response.usage().outputTokens()} output")

// ============================================================================
// Helper Functions
// ============================================================================

def uploadFile(client: AnthropicClient, path: Path): FileMetadata =
  val params = FileUploadParams.builder()
    .file(path)
    .addBeta(AnthropicBeta.FILES_API_2025_04_14)
    .build()

  client.beta().files().upload(params)
