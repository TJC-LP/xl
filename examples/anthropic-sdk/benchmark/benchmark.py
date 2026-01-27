# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "anthropic",
#     "anyio",
#     "pydantic",
#     "python-dotenv",
# ]
# ///
# Run with: uv run benchmark.py (or PYTHONUNBUFFERED=1 uv run benchmark.py for live output)
"""
Token Efficiency Benchmark: xl CLI (Custom Skill) vs Anthropic xlsx Skill

Compares token usage between:
1. xl CLI - Custom skill with uploaded binary
2. Anthropic xlsx - Built-in skill (openpyxl-based)

Usage:
    uv run benchmark.py                    # Run all tasks (parallel)
    uv run benchmark.py --task list_sheets # Run specific task
    uv run benchmark.py --xl-only          # Only run xl CLI tests
    uv run benchmark.py --sequential       # Disable parallel execution
"""

import argparse
import json
import logging
import os
import subprocess
import sys
import time
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Literal

import anyio
import anthropic
from dotenv import load_dotenv
from pydantic import BaseModel

from tasks import TASKS, LARGE_FILE_TASKS

# Load .env from script directory
load_dotenv(Path(__file__).parent / ".env")

# ============================================================================
# Logging (thread-safe)
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ============================================================================
# Configuration
# ============================================================================

SCRIPT_DIR = Path(__file__).parent
PARENT_DIR = SCRIPT_DIR.parent
RESULTS_DIR = SCRIPT_DIR / "results"

SAMPLE_PATH = PARENT_DIR / "sample.xlsx"
LARGE_SAMPLE_PATH = SCRIPT_DIR / "large_sample.xlsx"
BINARY_PATH = SCRIPT_DIR / "xl-linux-amd64"
SKILL_ZIP_PATH = SCRIPT_DIR / "xl-skill-0.8.0.zip"

# ============================================================================
# Data Classes
# ============================================================================


@dataclass
class TaskResult:
    task_id: str
    task_name: str
    approach: str  # "xl" or "xlsx"
    success: bool
    input_tokens: int
    output_tokens: int
    total_tokens: int
    latency_ms: int
    error: str | None = None
    response_text: str | None = None
    grade: str | None = None  # A/B/C/D/F letter grade
    grade_reasoning: str | None = None  # Brief explanation from grader


@dataclass
class BenchmarkResults:
    timestamp: str
    model: str
    sample_file: str
    results: list[TaskResult]

    def to_dict(self):
        return {
            "timestamp": self.timestamp,
            "model": self.model,
            "sample_file": self.sample_file,
            "results": [asdict(r) for r in self.results],
        }


# ============================================================================
# Grading (Opus 4.5 with Structured Output)
# ============================================================================


class GradeResult(BaseModel):
    """Structured grading result from Opus 4.5."""
    grade: Literal["A", "B", "C", "D", "F"]
    reason: str


async def grade_response(
    client: anthropic.AsyncAnthropic,
    task: dict,
    response_text: str,
) -> tuple[str, str]:
    """Grade a task response using Opus 4.5 with structured output.

    Returns: (letter_grade, reasoning)
    """
    grading_prompt = f"""You are grading an AI's response to an Excel analysis task.

TASK: {task['name']}
PROMPT: {task.get('xl_prompt', task.get('xlsx_prompt', ''))}

EXPECTED ANSWER (ground truth):
{task.get('expected_answer', 'No expected answer provided')}

AI'S ACTUAL RESPONSE:
{response_text}

Grade the response on correctness:
- A: Fully correct, all key information present
- B: Mostly correct, minor details missing or slight inaccuracies
- C: Partially correct, some key information wrong or missing
- D: Mostly incorrect, but shows some understanding
- F: Completely wrong or didn't answer the question"""

    try:
        response = await client.beta.messages.parse(
            model="claude-opus-4-5-20251101",
            max_tokens=256,
            betas=["structured-outputs-2025-11-13"],
            messages=[{"role": "user", "content": grading_prompt}],
            output_format=GradeResult,
        )

        result = response.parsed_output
        return result.grade, result.reason
    except Exception as e:
        log.warning(f"Grading failed: {e}")
        return "?", f"Grading error: {e}"


# ============================================================================
# Skill Management
# ============================================================================


def get_or_create_xl_skill(client: anthropic.Anthropic, skill_zip_path: Path) -> str:
    """Get existing xl-cli skill ID or create a new one."""
    import tempfile
    import zipfile

    skills = client.beta.skills.list(source="custom", betas=["skills-2025-10-02"])

    for skill in skills.data:
        if skill.display_title == "xl-cli":
            print(f"   Found existing xl-cli skill: {skill.id}")
            return skill.id

    print(f"   Creating xl-cli skill from {skill_zip_path.name}...")

    # Extract zip and upload files with proper path prefix
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(skill_zip_path, "r") as zf:
            zf.extractall(tmpdir)

        # Build file list with xl-cli/ prefix
        files = []
        tmppath = Path(tmpdir)
        for filepath in tmppath.rglob("*"):
            if filepath.is_file():
                relpath = filepath.relative_to(tmppath)
                # Add xl-cli/ prefix to satisfy "common root directory" requirement
                prefixed_path = f"xl-cli/{relpath}"
                files.append((prefixed_path, open(filepath, "rb")))

        try:
            skill = client.beta.skills.create(
                display_title="xl-cli",
                files=files,
                betas=["skills-2025-10-02"],
            )
        finally:
            # Close all file handles
            for _, f in files:
                f.close()

    print(f"   Created skill: {skill.id}")
    return skill.id


def download_release_assets() -> tuple[Path, Path]:
    """Download xl binary and skill zip from GitHub if not present."""
    binary_path = BINARY_PATH
    if not binary_path.exists():
        binaries = list(SCRIPT_DIR.glob("xl-*-linux-amd64"))
        if binaries:
            binary_path = binaries[0]
        else:
            print("   Downloading xl binary from GitHub...")
            result = subprocess.run(
                ["gh", "release", "download", "--repo", "TJC-LP/xl",
                 "--pattern", "xl-*-linux-amd64", "-D", str(SCRIPT_DIR)],
                capture_output=True, text=True
            )
            if result.returncode != 0:
                raise RuntimeError(f"Failed to download binary: {result.stderr}")
            binaries = list(SCRIPT_DIR.glob("xl-*-linux-amd64"))
            binary_path = binaries[0] if binaries else None
            if not binary_path:
                raise RuntimeError("Binary download succeeded but file not found")

    skill_path = SKILL_ZIP_PATH
    if not skill_path.exists():
        zips = list(SCRIPT_DIR.glob("xl-skill-*.zip"))
        if zips:
            skill_path = zips[0]
        else:
            print("   Downloading xl skill from GitHub...")
            result = subprocess.run(
                ["gh", "release", "download", "--repo", "TJC-LP/xl",
                 "--pattern", "xl-skill-*.zip", "-D", str(SCRIPT_DIR)],
                capture_output=True, text=True
            )
            if result.returncode != 0:
                raise RuntimeError(f"Failed to download skill: {result.stderr}")
            zips = list(SCRIPT_DIR.glob("xl-skill-*.zip"))
            skill_path = zips[0] if zips else None
            if not skill_path:
                raise RuntimeError("Skill download succeeded but file not found")

    return binary_path, skill_path


# ============================================================================
# Benchmark Runners (Async)
# ============================================================================


async def run_xl_task(
    client: anthropic.AsyncAnthropic,
    skill_id: str,
    binary_file_id: str,
    excel_file_id: str,
    task: dict,
) -> TaskResult:
    """Run a task using xl CLI custom skill."""
    start_time = time.time()

    system_prompt = """You have access to the xl CLI tool for Excel operations.

The xl binary has been uploaded to /mnt/user/xl-linux-amd64 - make it executable first.
The Excel file is at /mnt/user/sample.xlsx

Use xl commands to complete the task. Be concise in your response."""

    try:
        response = await client.beta.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=2048,
            betas=["code-execution-2025-08-25", "skills-2025-10-02", "files-api-2025-04-14"],
            system=system_prompt,
            container={
                "skills": [{"type": "custom", "skill_id": skill_id, "version": "latest"}]
            },
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": task["xl_prompt"]},
                        {"type": "container_upload", "file_id": binary_file_id},
                        {"type": "container_upload", "file_id": excel_file_id},
                    ],
                }
            ],
            tools=[{"type": "code_execution_20250825", "name": "code_execution"}],
        )

        latency_ms = int((time.time() - start_time) * 1000)
        response_text = "".join(b.text for b in response.content if b.type == "text")

        return TaskResult(
            task_id=task["id"],
            task_name=task["name"],
            approach="xl",
            success=True,
            input_tokens=response.usage.input_tokens,
            output_tokens=response.usage.output_tokens,
            total_tokens=response.usage.input_tokens + response.usage.output_tokens,
            latency_ms=latency_ms,
            response_text=response_text or None,
        )

    except Exception as e:
        latency_ms = int((time.time() - start_time) * 1000)
        return TaskResult(
            task_id=task["id"],
            task_name=task["name"],
            approach="xl",
            success=False,
            input_tokens=0,
            output_tokens=0,
            total_tokens=0,
            latency_ms=latency_ms,
            error=str(e),
        )


async def run_xlsx_task(
    client: anthropic.AsyncAnthropic,
    excel_file_id: str,
    task: dict,
) -> TaskResult:
    """Run a task using Anthropic's built-in xlsx skill."""
    start_time = time.time()

    system_prompt = """You have access to the xlsx skill for Excel operations.

The Excel file is at /mnt/user/sample.xlsx

Use Python with openpyxl to complete the task. Be concise in your response."""

    try:
        response = await client.beta.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=2048,
            betas=["code-execution-2025-08-25", "skills-2025-10-02", "files-api-2025-04-14"],
            system=system_prompt,
            container={
                "skills": [{"type": "anthropic", "skill_id": "xlsx", "version": "latest"}]
            },
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": task["xlsx_prompt"]},
                        {"type": "container_upload", "file_id": excel_file_id},
                    ],
                }
            ],
            tools=[{"type": "code_execution_20250825", "name": "code_execution"}],
        )

        latency_ms = int((time.time() - start_time) * 1000)
        response_text = "".join(b.text for b in response.content if b.type == "text")

        return TaskResult(
            task_id=task["id"],
            task_name=task["name"],
            approach="xlsx",
            success=True,
            input_tokens=response.usage.input_tokens,
            output_tokens=response.usage.output_tokens,
            total_tokens=response.usage.input_tokens + response.usage.output_tokens,
            latency_ms=latency_ms,
            response_text=response_text or None,
        )

    except Exception as e:
        latency_ms = int((time.time() - start_time) * 1000)
        return TaskResult(
            task_id=task["id"],
            task_name=task["name"],
            approach="xlsx",
            success=False,
            input_tokens=0,
            output_tokens=0,
            total_tokens=0,
            latency_ms=latency_ms,
            error=str(e),
        )


# ============================================================================
# Results Analysis
# ============================================================================


def print_comparison_table(results: list[TaskResult]):
    """Print a formatted comparison table with grades."""
    tasks = {}
    for r in results:
        if r.task_id not in tasks:
            tasks[r.task_id] = {"name": r.task_name}
        tasks[r.task_id][r.approach] = r

    # Check if any results have grades
    has_grades = any(r.grade for r in results)

    print("\n" + "=" * 120)
    print("TOKEN EFFICIENCY COMPARISON: xl CLI vs Anthropic xlsx Skill")
    print("=" * 120)

    if has_grades:
        print(
            f"{'Task':<20} | {'xl Tokens':>10} | {'xl Grade':>8} | {'xlsx Tokens':>11} | {'xlsx Grade':>10} | {'Winner':>8} | {'Savings':>8}"
        )
    else:
        print(
            f"{'Task':<20} | {'xl Input':>10} | {'xl Output':>10} | {'xlsx Input':>11} | {'xlsx Output':>12} | {'Winner':>8} | {'Savings':>8}"
        )
    print("-" * 120)

    total_xl_input = total_xl_output = total_xlsx_input = total_xlsx_output = 0
    xl_wins = xlsx_wins = 0
    xl_grades = []
    xlsx_grades = []

    for task_id, data in tasks.items():
        xl = data.get("xl")
        xlsx = data.get("xlsx")

        if xl and xlsx and xl.success and xlsx.success:
            total_xl_input += xl.input_tokens
            total_xl_output += xl.output_tokens
            total_xlsx_input += xlsx.input_tokens
            total_xlsx_output += xlsx.output_tokens

            if xl.grade:
                xl_grades.append(xl.grade)
            if xlsx.grade:
                xlsx_grades.append(xlsx.grade)

            if xl.total_tokens < xlsx.total_tokens:
                winner, savings = "xl", f"-{((xlsx.total_tokens - xl.total_tokens) / xlsx.total_tokens * 100):.0f}%"
                xl_wins += 1
            elif xlsx.total_tokens < xl.total_tokens:
                winner, savings = "xlsx", f"-{((xl.total_tokens - xlsx.total_tokens) / xl.total_tokens * 100):.0f}%"
                xlsx_wins += 1
            else:
                winner, savings = "tie", "0%"

            if has_grades:
                xl_grade_str = xl.grade or "-"
                xlsx_grade_str = xlsx.grade or "-"
                print(
                    f"{data['name']:<20} | {xl.total_tokens:>10,} | {xl_grade_str:>8} | "
                    f"{xlsx.total_tokens:>11,} | {xlsx_grade_str:>10} | {winner:>8} | {savings:>8}"
                )
            else:
                print(
                    f"{data['name']:<20} | {xl.input_tokens:>10,} | {xl.output_tokens:>10,} | "
                    f"{xlsx.input_tokens:>11,} | {xlsx.output_tokens:>12,} | {winner:>8} | {savings:>8}"
                )
        else:
            xl_tok = xl.total_tokens if xl and xl.success else "ERR"
            xlsx_tok = xlsx.total_tokens if xlsx and xlsx.success else "ERR"
            if has_grades:
                print(
                    f"{data['name']:<20} | {str(xl_tok):>10} | {'-':>8} | "
                    f"{str(xlsx_tok):>11} | {'-':>10} | {'N/A':>8} | {'N/A':>8}"
                )
            else:
                xl_in = xl.input_tokens if xl and xl.success else "ERR"
                xl_out = xl.output_tokens if xl and xl.success else "ERR"
                xlsx_in = xlsx.input_tokens if xlsx and xlsx.success else "ERR"
                xlsx_out = xlsx.output_tokens if xlsx and xlsx.success else "ERR"
                print(
                    f"{data['name']:<20} | {str(xl_in):>10} | {str(xl_out):>10} | "
                    f"{str(xlsx_in):>11} | {str(xlsx_out):>12} | {'N/A':>8} | {'N/A':>8}"
                )

    print("-" * 120)

    total_xl = total_xl_input + total_xl_output
    total_xlsx = total_xlsx_input + total_xlsx_output

    if total_xl > 0 and total_xlsx > 0:
        if total_xl < total_xlsx:
            overall_winner = "xl"
            overall_savings = f"-{((total_xlsx - total_xl) / total_xlsx * 100):.1f}%"
        else:
            overall_winner = "xlsx"
            overall_savings = f"-{((total_xl - total_xlsx) / total_xl * 100):.1f}%"

        if has_grades:
            # Calculate average grade (A=4, B=3, C=2, D=1, F=0)
            def grade_to_num(g: str) -> int:
                return {"A": 4, "B": 3, "C": 2, "D": 1, "F": 0}.get(g, -1)

            def avg_grade(grades: list[str]) -> str:
                valid = [grade_to_num(g) for g in grades if grade_to_num(g) >= 0]
                if not valid:
                    return "-"
                avg = sum(valid) / len(valid)
                if avg >= 3.5:
                    return "A"
                elif avg >= 2.5:
                    return "B"
                elif avg >= 1.5:
                    return "C"
                elif avg >= 0.5:
                    return "D"
                return "F"

            xl_avg = avg_grade(xl_grades)
            xlsx_avg = avg_grade(xlsx_grades)
            print(
                f"{'TOTAL':<20} | {total_xl:>10,} | {xl_avg:>8} | "
                f"{total_xlsx:>11,} | {xlsx_avg:>10} | {overall_winner:>8} | {overall_savings:>8}"
            )
        else:
            print(
                f"{'TOTAL':<20} | {total_xl_input:>10,} | {total_xl_output:>10,} | "
                f"{total_xlsx_input:>11,} | {total_xlsx_output:>12,} | {overall_winner:>8} | {overall_savings:>8}"
            )

    print("=" * 120)
    print(f"\nSummary: xl wins {xl_wins} tasks, xlsx wins {xlsx_wins} tasks")
    print(f"Total tokens: xl={total_xl:,}, xlsx={total_xlsx:,}")

    if total_xl < total_xlsx:
        print(f"xl CLI uses {((total_xlsx - total_xl) / total_xlsx * 100):.1f}% fewer tokens overall")
    elif total_xlsx < total_xl:
        print(f"Anthropic xlsx uses {((total_xl - total_xlsx) / total_xl * 100):.1f}% fewer tokens overall")

    # Print grade summary if available
    if has_grades and xl_grades and xlsx_grades:
        print(f"\nGrade distribution:")
        for approach, grades in [("xl", xl_grades), ("xlsx", xlsx_grades)]:
            counts = {g: grades.count(g) for g in "ABCDF" if grades.count(g) > 0}
            grade_str = ", ".join(f"{g}:{c}" for g, c in sorted(counts.items()))
            print(f"  {approach}: {grade_str}")


# ============================================================================
# Main
# ============================================================================


async def run_benchmark(
    tasks: list[dict],
    xl_only: bool,
    xlsx_only: bool,
    parallel: bool,
    enable_grading: bool,
) -> list[TaskResult]:
    """Run benchmark tasks, optionally in parallel."""
    # Sync client for setup (file uploads)
    sync_client = anthropic.Anthropic()

    log.info("Setting up...")

    # Upload Excel file
    log.info("Uploading sample.xlsx...")
    with open(SAMPLE_PATH, "rb") as f:
        excel_file = sync_client.beta.files.upload(file=f, betas=["files-api-2025-04-14"])
    log.info(f"Uploaded: {excel_file.id}")

    xl_skill_id = None
    binary_file_id = None

    if not xlsx_only:
        binary_path, skill_path = download_release_assets()
        log.info("Setting up xl-cli skill...")
        xl_skill_id = get_or_create_xl_skill(sync_client, skill_path)

        log.info(f"Uploading {binary_path.name}...")
        with open(binary_path, "rb") as f:
            binary_file = sync_client.beta.files.upload(file=f, betas=["files-api-2025-04-14"])
        binary_file_id = binary_file.id
        log.info(f"Uploaded: {binary_file_id}")

    # Async client for benchmark runs and grading
    async_client = anthropic.AsyncAnthropic()

    log.info(f"Running benchmarks ({'parallel' if parallel else 'sequential'}, grading={'on' if enable_grading else 'off'})...")

    async def run_and_grade(task: dict, approach: str, runner_coro) -> TaskResult:
        """Run a task and optionally grade the response."""
        result = await runner_coro

        # Grade the response if enabled and task succeeded
        if enable_grading and result.success and result.response_text:
            grade, reason = await grade_response(async_client, task, result.response_text)
            result.grade = grade
            result.grade_reasoning = reason
            grade_str = f" [{grade}]"
        else:
            grade_str = ""

        status = "OK" if result.success else f"ERR: {result.error}"
        log.info(f"[{task['name']}] {approach}: {result.input_tokens:,} in / {result.output_tokens:,} out ({status}){grade_str}")
        return result

    async def run_task_pair(task: dict) -> list[TaskResult]:
        """Run both xl and xlsx for a single task."""
        results = []
        if not xlsx_only:
            result = await run_and_grade(
                task, "xl",
                run_xl_task(async_client, xl_skill_id, binary_file_id, excel_file.id, task)
            )
            results.append(result)

        if not xl_only:
            result = await run_and_grade(
                task, "xlsx",
                run_xlsx_task(async_client, excel_file.id, task)
            )
            results.append(result)

        return results

    all_results = []

    if parallel:
        # Run all tasks in parallel
        async with anyio.create_task_group() as tg:
            results_list: list[list[TaskResult]] = [[] for _ in tasks]

            async def run_and_store(idx: int, task: dict):
                results_list[idx] = await run_task_pair(task)

            for i, task in enumerate(tasks):
                tg.start_soon(run_and_store, i, task)

        for results in results_list:
            all_results.extend(results)
    else:
        # Run sequentially
        for i, task in enumerate(tasks, 1):
            log.info(f"[{i}/{len(tasks)}] {task['name']}")
            results = await run_task_pair(task)
            all_results.extend(results)

    return all_results


def main():
    parser = argparse.ArgumentParser(description="Token efficiency benchmark: xl CLI vs Anthropic xlsx skill")
    parser.add_argument("--large", action="store_true", help="Include large file tasks")
    parser.add_argument("--task", help="Run specific task by ID")
    parser.add_argument("--xl-only", action="store_true", help="Only run xl CLI tests")
    parser.add_argument("--xlsx-only", action="store_true", help="Only run xlsx skill tests")
    parser.add_argument("--sequential", action="store_true", help="Run tasks sequentially (default: parallel)")
    parser.add_argument("--no-grade", action="store_true", help="Skip correctness grading with Opus 4.5")
    parser.add_argument("--output", help="Output file for JSON results")
    args = parser.parse_args()

    # Validate prerequisites
    if not SAMPLE_PATH.exists():
        log.error(f"Sample file not found: {SAMPLE_PATH}")
        log.error("  Generate it: scala-cli run examples/anthropic-sdk/create_sample.sc")
        sys.exit(1)

    if not os.getenv("ANTHROPIC_API_KEY"):
        log.error("ANTHROPIC_API_KEY not found")
        log.error("  Set it as environment variable or in .env file")
        sys.exit(1)

    print("=" * 60)
    print("Token Efficiency Benchmark: xl CLI vs Anthropic xlsx Skill")
    print("=" * 60)

    # Select tasks
    tasks = TASKS.copy()
    if args.large and LARGE_SAMPLE_PATH.exists():
        tasks.extend(LARGE_FILE_TASKS)
    if args.task:
        tasks = [t for t in tasks if t["id"] == args.task]
        if not tasks:
            log.error(f"Task '{args.task}' not found")
            sys.exit(1)

    grading_status = "off" if args.no_grade else "on (Opus 4.5)"
    log.info(f"Running {len(tasks)} tasks, grading: {grading_status}")

    # Run benchmark
    results = anyio.run(
        run_benchmark,
        tasks,
        args.xl_only,
        args.xlsx_only,
        not args.sequential,
        not args.no_grade,  # enable_grading
    )

    # Print comparison
    if not args.xl_only and not args.xlsx_only:
        print_comparison_table(results)

    # Save results
    RESULTS_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = Path(args.output) if args.output else RESULTS_DIR / f"benchmark_{timestamp}.json"

    benchmark_results = BenchmarkResults(
        timestamp=timestamp,
        model="claude-sonnet-4-5-20250929",
        sample_file=str(SAMPLE_PATH),
        results=results,
    )

    with open(output_path, "w") as f:
        json.dump(benchmark_results.to_dict(), f, indent=2)

    print(f"\nResults saved to: {output_path}")


if __name__ == "__main__":
    main()
