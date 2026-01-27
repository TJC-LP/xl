# Token Efficiency Benchmark: xl CLI vs Anthropic xlsx Skill

This benchmark compares token usage between two approaches for Excel operations via the Anthropic API:

1. **xl CLI (Custom Skill)** - Our LLM-optimized Excel CLI uploaded as a custom skill
2. **Anthropic xlsx Skill (Built-in)** - Anthropic's built-in xlsx skill using openpyxl

## Hypothesis

xl CLI should demonstrate token efficiency due to:
- **Compact output**: Markdown tables vs verbose Python data structures
- **Single commands**: `xl -f file.xlsx sheets` vs multi-line Python code
- **Purpose-built for LLMs**: Output format designed for model consumption

## Features

- **Token Efficiency**: Compare input/output token counts between approaches
- **Correctness Grading**: Opus 4.5 grades each response against expected answers (A/B/C/D/F) using [structured outputs](https://platform.claude.com/docs/en/build-with-claude/structured-outputs)
- **Parallel Execution**: Run all tasks concurrently with anyio
- **Thread-Safe Logging**: Proper async logging for parallel runs

## Quick Start

### 1. Prerequisites

```bash
# Install uv (if not already installed)
# macOS/Linux: curl -LsSf https://astral.sh/uv/install.sh | sh
# Windows: powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Install GitHub CLI (for auto-download of xl assets)
# macOS: brew install gh
# Linux: https://github.com/cli/cli#installation

# Set API key as environment variable
export ANTHROPIC_API_KEY="your-key-here"
```

### 2. Generate Sample Files

```bash
# From repository root:
./mill __.publishLocal  # If not already done

# Small sample (in parent directory)
scala-cli run ../create_sample.sc

# Large sample (1000 rows) - optional
scala-cli run create_large_sample.sc
```

### 3. Run Benchmark

The benchmark uses inline dependencies (PEP 723) - no manual install needed. It auto-downloads the xl binary and skill zip from GitHub releases if not present.

```bash
# Run all tasks in parallel with grading (default)
uv run benchmark.py

# Skip correctness grading (faster, cheaper)
uv run benchmark.py --no-grade

# Run sequentially (easier to follow output)
uv run benchmark.py --sequential

# Include large file tasks
uv run benchmark.py --large

# Run specific task
uv run benchmark.py --task list_sheets

# Only run one approach
uv run benchmark.py --xl-only
uv run benchmark.py --xlsx-only
```

## Benchmark Tasks

| Task | Description |
|------|-------------|
| `list_sheets` | List all sheets with basic info |
| `view_range` | Display a range of cells as a table |
| `statistics` | Calculate statistics on a numeric range |
| `show_formulas` | Display formulas instead of computed values |
| `search` | Find cells containing a pattern |
| `cell_details` | Get details about a cell including dependencies |
| `cross_sheet_ref` | Understand cross-sheet formula references |

### Large File Tasks (optional)

| Task | Description |
|------|-------------|
| `large_view` | Display first 20 rows of 1000-row dataset |
| `large_stats` | Compute statistics on 1000 values |
| `large_search` | Search in large file |

## Output

### Console Output

```
========================================================================================================================
TOKEN EFFICIENCY COMPARISON: xl CLI vs Anthropic xlsx Skill
========================================================================================================================
Task                 | xl Tokens  | xl Grade | xlsx Tokens | xlsx Grade | Winner   | Savings
------------------------------------------------------------------------------------------------------------------------
List Sheets          |     37,837 |        A |      44,783 |          A |       xl |    -15%
View Range           |     38,457 |        A |      44,308 |          B |       xl |    -13%
Statistics           |     65,797 |        A |      71,563 |          A |       xl |     -8%
...
------------------------------------------------------------------------------------------------------------------------
TOTAL                |    312,304 |        A |     523,813 |          B |       xl |  -40.4%
========================================================================================================================

Summary: xl wins 7 tasks, xlsx wins 0 tasks
Total tokens: xl=312,304, xlsx=523,813
xl CLI uses 40.4% fewer tokens overall

Grade distribution:
  xl: A:7
  xlsx: A:5, B:2
```

### JSON Results

Results are saved to `results/benchmark_YYYYMMDD_HHMMSS.json`:

```json
{
  "timestamp": "20250127_154523",
  "model": "claude-sonnet-4-5-20250929",
  "sample_file": "sample.xlsx",
  "results": [
    {
      "task_id": "list_sheets",
      "task_name": "List Sheets",
      "approach": "xl",
      "success": true,
      "input_tokens": 37319,
      "output_tokens": 518,
      "total_tokens": 37837,
      "latency_ms": 22501,
      "grade": "A",
      "grade_reasoning": "Correctly identified both sheets with accurate cell counts."
    },
    ...
  ]
}
```

## Architecture

### xl CLI Approach

```
┌─────────────────────────────────────────────────┐
│ Custom Skill (xl-skill-0.8.0.zip)               │
│ - SKILL.md with CLI reference                   │
│ - Installation instructions                      │
│ - Command examples                              │
└─────────────────────────────────────────────────┘
                      +
┌─────────────────────────────────────────────────┐
│ Uploaded Binary (xl-linux-amd64)                │
│ - Native binary, no JDK required                │
│ - ~54MB upload                                   │
└─────────────────────────────────────────────────┘
                      +
┌─────────────────────────────────────────────────┐
│ Excel File (sample.xlsx)                        │
└─────────────────────────────────────────────────┘
```

### Anthropic xlsx Approach

```
┌─────────────────────────────────────────────────┐
│ Built-in Skill (anthropic/xlsx)                 │
│ - Pre-installed openpyxl                        │
│ - Anthropic-maintained instructions              │
└─────────────────────────────────────────────────┘
                      +
┌─────────────────────────────────────────────────┐
│ Excel File (sample.xlsx)                        │
└─────────────────────────────────────────────────┘
```

## Metrics Captured

| Metric | Description |
|--------|-------------|
| `input_tokens` | Tokens in system prompt + user message + skill definitions |
| `output_tokens` | Tokens in Claude's response including tool use |
| `total_tokens` | Sum of input + output |
| `latency_ms` | Wall-clock time for the API call |
| `success` | Whether the task completed without errors |
| `grade` | Letter grade (A/B/C/D/F) from Opus 4.5 correctness grading |
| `grade_reasoning` | Brief explanation from the grader |

## Files

| File | Description |
|------|-------------|
| `benchmark.py` | Main benchmark runner |
| `tasks.py` | Task definitions for both approaches |
| `create_large_sample.sc` | Generate 1000-row test file |
| `results/` | JSON output from benchmark runs |
| `xl-*-linux-amd64` | Linux binary (auto-downloaded) |
| `xl-skill-*.zip` | Skill package (auto-downloaded) |
| `large_sample.xlsx` | Generated large test file |

## Expected Results

Based on xl CLI design principles:

| Metric | xl CLI Expected | xlsx Skill Expected |
|--------|-----------------|---------------------|
| Output tokens | Lower (compact tables) | Higher (Python print output) |
| Input tokens | Higher (skill docs + binary) | Lower (built-in skill) |
| Net efficiency | Depends on task | Depends on task |

**Key insight**: xl CLI's advantage grows with output complexity. For simple tasks (list sheets), overhead may not pay off. For complex analysis with multiple ranges, compact output wins.

## Contributing

To add new benchmark tasks, edit `tasks.py`:

```python
{
    "id": "my_task",
    "name": "My Task",
    "description": "What this task tests",
    "xl_prompt": "Prompt for xl CLI approach",
    "xlsx_prompt": "Prompt for xlsx skill approach",
    "expected_answer": """Ground truth for grading.
Include the key facts that should be in the response.""",
}
```
