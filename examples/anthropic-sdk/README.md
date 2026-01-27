# Anthropic SDK + XL CLI Code Execution Examples

This directory contains examples demonstrating how to use the [Anthropic API's code execution tool](https://docs.anthropic.com/en/docs/agents-and-tools/tool-use/code-execution-tool) with the XL CLI for Excel operations.

## Overview

The code execution sandbox runs in an isolated Linux container with no network access. To use the `xl` CLI, we:

1. Upload the pre-built Linux binary via the Files API
2. Upload Excel files for analysis
3. Include skill instructions in the system prompt
4. Enable the code execution tool
5. Claude makes the binary executable and runs xl commands

## Prerequisites

1. **API Key**: Copy `.env.example` to `.env` and add your API key
2. **GitHub CLI**: Install `gh` for downloading release binaries
3. **XL Library**: Publish locally with `./mill __.publishLocal`

## Quick Start

### 1. Set Up API Key

```bash
cp examples/anthropic-sdk/.env.example examples/anthropic-sdk/.env
# Edit .env and add your ANTHROPIC_API_KEY
```

### 2. Download the Linux Binary

```bash
# Download the latest Linux x86_64 binary
gh release download --repo TJC-LP/xl --pattern "xl-*-linux-amd64" -D examples/anthropic-sdk

# Rename to standard name
mv examples/anthropic-sdk/xl-*-linux-amd64 examples/anthropic-sdk/xl-linux-amd64
```

### 3. Generate Sample Data

```bash
# Publish xl locally (if not already done)
./mill __.publishLocal

# Generate sample.xlsx
scala-cli run examples/anthropic-sdk/create_sample.sc
```

### 4. Run the Examples

**Scala (using Java SDK):**
```bash
scala-cli run examples/anthropic-sdk/xl_code_execution.sc
```

**Python:**
```bash
pip install anthropic python-dotenv
python examples/anthropic-sdk/xl_code_execution.py
```

## Files

| File | Description |
|------|-------------|
| `xl_code_execution.sc` | Scala-CLI example using Anthropic Java SDK |
| `xl_code_execution.py` | Python example using Anthropic Python SDK |
| `create_sample.sc` | Script to generate sample.xlsx |
| `sample.xlsx` | Sample spreadsheet with formulas and styling |
| `.env.example` | Template for API key configuration |
| `.env` | Your API key (gitignored) |
| `xl-linux-amd64` | Linux binary (downloaded from releases) |
| `benchmark/` | Token efficiency benchmark (xl CLI vs Anthropic xlsx skill) |

## How It Works

1. **Environment**: API key loaded from `.env` file using dotenv
2. **File Upload**: Both the xl binary (~54MB) and Excel files are uploaded via the Files API
3. **Code Execution**: The `code_execution_20250825` tool gives Claude access to bash and file editing
4. **Skill Instructions**: System prompt explains how to use xl CLI with uploaded file paths
5. **Execution**: Claude makes the binary executable and runs analysis commands

## Example Output

```
=== XL CLI + Anthropic Code Execution Example ===

1. Uploading xl binary...
   Uploaded: file_abc123 (54528264 bytes)
2. Uploading sample.xlsx...
   Uploaded: file_def456 (3649 bytes)
3. Sending request to Claude with code execution tool...

=== Claude's Response ===

I'll analyze the Excel file for you...

## Summary of Findings

### 1. Sheets in the Workbook
The Excel file contains 2 sheets:
- Sales - Contains 36 cells with 13 formulas
- Summary - Contains 9 cells with 3 formulas

### 2. Sales Sheet Data
...

=== Done ===
Stop reason: Optional[end_turn]
Usage: 30541 input, 1664 output
```

## Customization

### Using Your Own Excel Files

Replace the file paths in the examples:

```python
# Python
SAMPLE_PATH = SCRIPT_DIR / "your-file.xlsx"
```

```scala
// Scala
val samplePath = scriptDir.resolve("your-file.xlsx")
```

### Modifying Skill Instructions

Edit `SKILL_INSTRUCTIONS` / `skillInstructions` to:
- Add more xl commands
- Include domain-specific guidance
- Reference additional uploaded files

## Troubleshooting

### Binary Not Found
```bash
ls -la examples/anthropic-sdk/xl-linux-amd64
```

### Sample File Missing
```bash
scala-cli run examples/anthropic-sdk/create_sample.sc
```

### API Key Not Set
Check `.env` file exists and contains valid key:
```bash
cat examples/anthropic-sdk/.env
```

## Dependencies

**Scala:**
- `com.anthropic:anthropic-java:2.11.1`
- `io.github.cdimascio:dotenv-java:3.2.0`

**Python:**
- `anthropic`
- `python-dotenv`

## Token Efficiency Benchmark

The `benchmark/` directory contains a benchmark comparing xl CLI vs Anthropic's built-in xlsx skill:

```bash
cd benchmark
uv run benchmark.py  # Parallel execution, auto-downloads assets from GitHub
```

See [benchmark/README.md](benchmark/README.md) for details.

## Related Documentation

- [XL CLI Skill](../../plugin/skills/xl-cli/SKILL.md) - Full xl CLI reference
- [Anthropic Code Execution](https://docs.anthropic.com/en/docs/agents-and-tools/tool-use/code-execution-tool) - API documentation
- [Anthropic Skills API](https://docs.anthropic.com/en/docs/agents-and-tools/agent-skills/using-skills) - Skills API documentation
- [Anthropic Java SDK](https://github.com/anthropics/anthropic-sdk-java) - Java/Scala SDK
