package com.tjclp.xl.formula.functions

/**
 * GH-713: the names of the functions that return a reference, as plain constants.
 *
 * A leaf object on purpose: it references nothing, so the parser can consult it (every `:`, union
 * and intersection asks whether an operand is reference-shaped) without forcing `FunctionRegistry`
 * or `FunctionSpecs` to initialize, and no object-initialization cycle can reach it. The drift law
 * in `ReferenceOperatorLawsSpec` ties [[computing]] to the registry: it must equal the names of
 * every registered `FunctionSpec.Referencing`, so a new reference-computing spec fails that test
 * until it is added here.
 */
private[formula] object ReferenceFunctionNames:

  /**
   * Excel's functions that return the reference they select. They evaluate their branches lazily
   * rather than through `FunctionSpec.referencing`, so no spec marks them.
   */
  val selecting: Set[String] = Set("IF", "IFS", "CHOOSE", "SWITCH")

  /** The registered functions that compute a reference: every `FunctionSpec.Referencing`. */
  val computing: Set[String] = Set("OFFSET", "INDIRECT", "INDEX", "XLOOKUP")
