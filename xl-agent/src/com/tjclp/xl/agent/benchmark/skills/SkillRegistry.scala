package com.tjclp.xl.agent.benchmark.skills

import scala.collection.mutable

/**
 * Registry for benchmark skills
 *
 * Allows looking up skills by name and iterating over all registered skills. Built-in skills are
 * registered at startup; custom skills can be added dynamically.
 */
object SkillRegistry:
  private val skills: mutable.Map[String, Skill] = mutable.Map.empty
  @volatile private var initialized: Boolean = false

  /** Initialize built-in skills (called once at startup) */
  def initialize(): Unit =
    if !initialized then
      synchronized {
        if !initialized then
          import com.tjclp.xl.agent.benchmark.skills.builtin.{XlSkill, XlsxSkill}
          register(XlSkill)
          register(XlsxSkill)
          initialized = true
      }

  /** Register a skill */
  def register(skill: Skill): Unit =
    skills(skill.name) = skill

  /** Get a skill by name */
  def get(name: String): Option[Skill] =
    skills.get(name)

  /** Get a skill by name, throwing if not found */
  def apply(name: String): Skill =
    skills.getOrElse(
      name,
      throw new IllegalArgumentException(
        s"Unknown skill: $name. Available skills: ${names.mkString(", ")}"
      )
    )

  /** All registered skills */
  def all: List[Skill] = skills.values.toList.sortBy(_.name)

  /** All skill names */
  def names: List[String] = skills.keys.toList.sorted

  /** Check if a skill is registered */
  def contains(name: String): Boolean = skills.contains(name)

  /** Parse a comma-separated list of skill names */
  def parseSkillList(input: String): Either[String, List[Skill]] =
    val skillNames = input.split(",").map(_.trim).filter(_.nonEmpty).toList
    val (found, missing) = skillNames.partition(contains)
    if missing.nonEmpty then
      Left(s"Unknown skills: ${missing.mkString(", ")}. Available: ${names.mkString(", ")}")
    else Right(skillNames.map(apply))
