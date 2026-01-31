package com.tjclp.xl.agent.benchmark.execution

import com.tjclp.xl.agent.benchmark.skills.{Skill, SkillContext}
import com.tjclp.xl.agent.benchmark.task.{BenchmarkTask, TaskId, TestCaseFile}

/**
 * A single unit of work: one test case for one skill on one task.
 *
 * By flattening (task, skill, case) into independent work units, the engine can schedule all work
 * with a single `parTraverseN`, providing:
 *   - Single parallelism control point
 *   - Better load balancing across cases
 *   - More accurate progress tracking
 *   - No nested parallelism (avoids thread explosion)
 *
 * @param task
 *   The benchmark task this case belongs to
 * @param skill
 *   The skill to use for execution
 * @param skillContext
 *   Context from skill setup (uploaded files, skill IDs, etc.)
 * @param testCase
 *   The specific test case to execute
 * @param caseIndex
 *   Zero-based index of this case within the task-skill combination
 * @param totalCases
 *   Total number of cases for this task-skill combination
 */
case class WorkUnit(
  task: BenchmarkTask,
  skill: Skill,
  skillContext: SkillContext,
  testCase: TestCaseFile,
  caseIndex: Int,
  totalCases: Int
):
  /** Task ID for grouping results */
  def taskId: TaskId = task.id

  /** Skill name for grouping results */
  def skillName: String = skill.name

  /** Human-readable identifier for logging */
  def identifier: String = s"${task.taskIdValue}/${skill.name}/${testCase.caseNum}"

/**
 * Result of executing a single work unit.
 *
 * These are aggregated back into ExecutionResult by (taskId, skillName) after parallel execution.
 *
 * @param taskId
 *   The task this result belongs to
 * @param skillName
 *   The skill that executed this case
 * @param caseResult
 *   The individual case result
 */
case class WorkUnitResult(
  taskId: TaskId,
  skillName: String,
  caseResult: CaseResult
)
