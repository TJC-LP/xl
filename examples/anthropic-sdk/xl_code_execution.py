#!/usr/bin/env python3
"""
XL CLI + Anthropic Code Execution Example

This script demonstrates using the Anthropic API's code execution tool
with the XL CLI for Excel operations. The code execution sandbox has
no network access, so we upload the pre-built Linux binary.

Prerequisites:
    1. Copy .env.example to .env and add your API key
    2. Install dependencies: pip install anthropic python-dotenv
    3. Download the Linux binary:
       gh release download --repo TJC-LP/xl --pattern "xl-*-linux-amd64" -D examples/anthropic-sdk
       mv examples/anthropic-sdk/xl-*-linux-amd64 examples/anthropic-sdk/xl-linux-amd64

Usage:
    python examples/anthropic-sdk/xl_code_execution.py
"""

import os
import sys
from pathlib import Path

from dotenv import load_dotenv

# ============================================================================
# Configuration
# ============================================================================

SCRIPT_DIR = Path(__file__).parent
BINARY_PATH = SCRIPT_DIR / "xl-linux-amd64"
SAMPLE_PATH = SCRIPT_DIR / "sample.xlsx"
ENV_PATH = SCRIPT_DIR / ".env"

# Load environment from .env file
load_dotenv(ENV_PATH)

# ============================================================================
# Skill Instructions
# ============================================================================

SKILL_INSTRUCTIONS = """
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
""".strip()


def main():
    # Validate prerequisites
    if not BINARY_PATH.exists():
        print(f"""
Error: Linux binary not found at {BINARY_PATH}

Download it with:
  gh release download --repo TJC-LP/xl --pattern "xl-*-linux-amd64" -D examples/anthropic-sdk
  mv examples/anthropic-sdk/xl-*-linux-amd64 examples/anthropic-sdk/xl-linux-amd64
""", file=sys.stderr)
        sys.exit(1)

    if not SAMPLE_PATH.exists():
        print(f"""
Error: Sample file not found at {SAMPLE_PATH}

Generate it with:
  scala-cli run examples/anthropic-sdk/create_sample.sc
""", file=sys.stderr)
        sys.exit(1)

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("""
Error: ANTHROPIC_API_KEY not found

Set it in examples/anthropic-sdk/.env or as an environment variable.
See .env.example for the format.
""", file=sys.stderr)
        sys.exit(1)

    # Import anthropic after validation to give better error messages
    import anthropic

    print("=== XL CLI + Anthropic Code Execution Example ===\n")

    # Initialize client
    client = anthropic.Anthropic(api_key=api_key)

    # Step 1: Upload the Linux binary
    print("1. Uploading xl binary...")
    with open(BINARY_PATH, "rb") as f:
        binary_file = client.beta.files.upload(file=f)
    print(f"   Uploaded: {binary_file.id} ({BINARY_PATH.stat().st_size} bytes)")

    # Step 2: Upload the sample Excel file
    print("2. Uploading sample.xlsx...")
    with open(SAMPLE_PATH, "rb") as f:
        excel_file = client.beta.files.upload(file=f)
    print(f"   Uploaded: {excel_file.id} ({SAMPLE_PATH.stat().st_size} bytes)")

    # Step 3: Create message with code execution
    print("3. Sending request to Claude with code execution tool...")
    print()

    user_prompt = """Analyze the uploaded Excel file. Please:
1. List all sheets
2. View the Sales sheet data
3. Calculate statistics on the quarterly data (B2:E4)
4. Show the formulas in column F
5. Summarize your findings"""

    response = client.beta.messages.create(
        model="claude-sonnet-4-5-20250929",
        betas=["code-execution-2025-08-25", "files-api-2025-04-14"],
        max_tokens=4096,
        system=SKILL_INSTRUCTIONS,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": user_prompt},
                    {"type": "container_upload", "file_id": binary_file.id},
                    {"type": "container_upload", "file_id": excel_file.id},
                ],
            }
        ],
        tools=[{"type": "code_execution_20250825", "name": "code_execution"}],
    )

    # Step 4: Process the response
    print("=== Claude's Response ===\n")

    for block in response.content:
        if block.type == "text":
            print(block.text)
        # server_tool_use and tool_result blocks are handled internally

    print("\n=== Done ===")
    print(f"Stop reason: {response.stop_reason}")
    print(f"Usage: {response.usage.input_tokens} input, {response.usage.output_tokens} output")


if __name__ == "__main__":
    main()
