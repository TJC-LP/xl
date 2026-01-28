# Token Efficiency Benchmark: xl CLI vs Anthropic xlsx Skill

This benchmark compares token usage between two approaches for Excel operations via the Anthropic API:

1. **xl CLI (Custom Skill)** - Our LLM-optimized Excel CLI uploaded as a custom skill
2. **Anthropic xlsx Skill (Built-in)** - Anthropic's built-in xlsx skill using openpyxl

> **Note**: The benchmark has been migrated to the `xl-agent` module for better integration and maintainability. This directory now only contains the binary assets.

## Quick Start

```bash
# From repository root:
./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- [options]

# Run all tasks with grading
./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner

# Skip grading (faster, cheaper)
./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- --no-grade

# Run specific task
./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- --task list_sheets

# Only run xl CLI tests
./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- --xl-only

# Only run xlsx skill tests
./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- --xlsx-only

# Run sequentially with verbose output
./mill xl-agent.runMain com.tjclp.xl.agent.benchmark.token.TokenRunner -- --sequential --verbose
```

## CLI Options

```
Task Selection:
  --task <id>         Run a specific task by ID
  --large             Include large file tasks

Approach Selection:
  --xl-only           Only run xl CLI tests (skip xlsx skill)
  --xlsx-only         Only run xlsx skill tests (skip xl)

Execution:
  --sequential        Run tasks sequentially (default: parallel)
  --no-grade          Skip correctness grading with Opus 4.5

Output:
  --output <path>     Output file for JSON results
  --verbose           Show detailed output

Development:
  --xl-binary <path>  Use local xl binary instead of released version
  --xl-skill <path>   Use local xl skill zip instead of released version
  --sample <path>     Custom sample.xlsx path
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

### Large File Tasks (with `--large`)

| Task | Description |
|------|-------------|
| `large_view` | Display first 20 rows of 1000-row dataset |
| `large_stats` | Compute statistics on 1000 values |
| `large_search` | Search in large file |

## Files in This Directory

| File | Description |
|------|-------------|
| `xl-*-linux-amd64` | Linux binary (auto-downloaded from GitHub releases) |
| `xl-skill-*.zip` | Skill package (auto-downloaded from GitHub releases) |
| `create_large_sample.sc` | Generate 1000-row test file |
| `.gitignore` | Ignores results and binaries |

## Source Code Location

The benchmark implementation is in the `xl-agent` module:

- `xl-agent/src/com/tjclp/xl/agent/benchmark/token/TokenRunner.scala` - Main entrypoint
- `xl-agent/src/com/tjclp/xl/agent/benchmark/token/TokenTasks.scala` - Task definitions
- `xl-agent/src/com/tjclp/xl/agent/benchmark/token/TokenGrader.scala` - Opus 4.5 grading
- `xl-agent/src/com/tjclp/xl/agent/benchmark/token/TokenReporter.scala` - Output formatting
- `xl-agent/src/com/tjclp/xl/agent/benchmark/token/TokenModels.scala` - Data types

## Expected Results

Based on xl CLI design principles:

| Metric | xl CLI Expected | xlsx Skill Expected |
|--------|-----------------|---------------------|
| Output tokens | Lower (compact tables) | Higher (Python print output) |
| Input tokens | Higher (skill docs + binary) | Lower (built-in skill) |
| Net efficiency | Depends on task | Depends on task |

**Key insight**: xl CLI's advantage grows with output complexity. For simple tasks (list sheets), overhead may not pay off. For complex analysis with multiple ranges, compact output wins.
