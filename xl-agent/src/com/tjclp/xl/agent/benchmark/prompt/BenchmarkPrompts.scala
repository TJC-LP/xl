package com.tjclp.xl.agent.benchmark.prompt

/**
 * Centralized prompts for benchmark execution.
 *
 * These provide behavioral guidance that is consistent across all skills. Skills provide only
 * tool-specific instructions (commands, file paths).
 */
object BenchmarkPrompts:

  /**
   * System prompt preamble for all benchmark tasks.
   *
   * Emphasizes:
   *   - Token efficiency (minimize tool calls)
   *   - Task focus (only do what's asked)
   *   - OUTPUT_DIR workflow (/tmp → $OUTPUT_DIR pattern)
   */
  val systemPreamble: String = """
CRITICAL EXECUTION GUIDELINES:
1. Complete ONLY the specified task - no extra features, formatting, or analysis
2. Be token-efficient - minimize tool calls
3. Do NOT explore or verify unless the task explicitly requires it
4. For simple tasks, accomplish everything in 1-2 tool calls

BASH ENVIRONMENT:
Shell variables DO NOT persist between tool calls. Each bash execution starts fresh.
- Always re-set variables (XL, INPUT, etc.) at the start of EVERY script
- $OUTPUT_DIR changes between calls - save to /tmp then copy in your FINAL call
- Or do all work in a SINGLE tool call that writes directly to $OUTPUT_DIR

AVOID:
- Reading documentation unless you don't know how to use a tool
- Exploring file content unless the task requires understanding it
- Creating styled/formatted output unless explicitly requested
- Verification steps after simple operations
""".trim

  /**
   * User prompt footer for all tasks.
   *
   * Reinforces token efficiency at the end of task instructions.
   */
  val taskFooter: String = """
IMPORTANT: Complete this task with minimum tool calls. Simple tasks should need 1-2 calls.
""".trim
