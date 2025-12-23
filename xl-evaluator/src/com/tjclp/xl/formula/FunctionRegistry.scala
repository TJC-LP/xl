package com.tjclp.xl.formula

object FunctionRegistry:
  inline def all: List[FunctionSpec[?]] =
    ${ FunctionRegistryMacro.collect[FunctionSpecs.type] }

  private lazy val byName: Map[String, FunctionSpec[?]] =
    all.map(spec => spec.name.toUpperCase -> spec).toMap

  def lookup(name: String): Option[FunctionSpec[?]] =
    byName.get(name.toUpperCase)

  def allNames: List[String] =
    byName.keys.toList.sorted
