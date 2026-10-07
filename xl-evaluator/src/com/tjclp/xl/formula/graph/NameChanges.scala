package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{Evaluator, NameWalk, SheetEvaluator}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.workbooks.{DefinedName, Workbook}

/** Formula roots invalidated by a change to the defined-name table, including aliases and scope. */
private[xl] object NameChanges:
  def readers(before: Workbook, after: Workbook): Set[QualifiedRef] =
    if before.metadata.definedNames == after.metadata.definedNames then Set.empty
    else
      // A name reference as looked up: (lookup sheet, the sheet an unqualified name in its target
      // resolves from, name). The memo keys on the exact spelling; the index applies the case fold.
      type NameKey = (SheetName, SheetName, String)
      val canonicalBefore = DependencyGraph.sheetCanonicaliser(before)
      val canonicalAfter = DependencyGraph.sheetCanonicaliser(after)
      val positionsBefore =
        before.sheets.zipWithIndex.reverseIterator.map((s, i) => s.name -> i).toMap
      val positionsAfter =
        after.sheets.zipWithIndex.reverseIterator.map((s, i) => s.name -> i).toMap
      // Invocation-local memoization: many formulas share a name/alias chain. No state escapes.
      val parsed = scala.collection.mutable.HashMap.empty[String, Option[TExpr[?]]]

      def binding(wb: Workbook, dn: DefinedName): (String, Option[SheetName]) =
        (dn.formula, Evaluator.definedNameScope(wb, dn).map(_.name))

      // Whether `text`, read from `current`, may read a changed name, each name reference it makes
      // answered by `matches(name, lookup sheet, fallback sheet)`. A dynamic call or text the
      // parser rejects (scanned instead) answers true on its own: neither certifies independence.
      def mayRead(
        text: String,
        current: SheetName,
        matches: (String, SheetName, SheetName) => Boolean
      ): Boolean =
        parsed.getOrElseUpdate(text, FormulaParser.parse(text).toOption) match
          case Some(expr) =>
            DependencyGraph.referencesMatching(
              expr,
              (name, scope) => matches(name, scope.getOrElse(current), current),
              includeDynamicCalls = true
            )
          case None => ReferenceScan.mayReferenceName(after, current, text, matches)

      // A reference's own verdict: its binding differs between the two tables, or its target's
      // text certifies nothing by itself; then the references that target makes. Stepped once per
      // reference by the walk.
      def step(key: NameKey): NameWalk.Step[NameKey] =
        val (lookup, fallback, name) = key
        val old = positionsBefore
          .get(canonicalBefore(lookup))
          .flatMap(i => Evaluator.lookupDefinedNameAt(before, Some(i), name))
        val fresh = positionsAfter
          .get(lookup)
          .flatMap(i => Evaluator.lookupDefinedNameAt(after, Some(i), name))
        val rebound = old.map(binding(before, _)) != fresh.map(binding(after, _))
        fresh.fold(NameWalk.Step(rebound, Vector.empty[NameKey])) { dn =>
          val current = Evaluator.definedNameScope(after, dn).map(_.name).getOrElse(fallback)
          val refs = Vector.newBuilder[NameKey]
          val opaque = mayRead(
            dn.formula,
            current,
            (name, lookupFrom, from) =>
              refs += ((canonicalAfter(lookupFrom), from, name))
              false
          )
          val next = refs.result()
          NameWalk.Step(rebound || opaque, if next.sizeIs > 1 then next.distinct else next)
        }

      // #691: a [[NameWalk.Walk]] over the references a name reaches — a loop, so a chain of any
      // length leaves the stack alone, and each reference is stepped once however many readers and
      // paths reach it. A cycle, or a chain more than NameWalk.MaxDepth names long, cannot certify
      // independence: conservatively changed.
      val walk = NameWalk.Walk(step)

      def nameChanged(name: String, lookupFrom: SheetName, fallback: SheetName): Boolean =
        val verdict = walk((canonicalAfter(lookupFrom), fallback, name))
        verdict.found || verdict.cyclic || verdict.tooDeep

      after.sheets.iterator.flatMap { sheet =>
        sheet.cells.iterator.collect {
          case (ref, cell) if (cell.value match
                case CellValue.Formula(_, _, _: FormulaKind.DataTable) => false
                case CellValue.Formula(text, _, _) =>
                  SheetEvaluator.pinnedExternalCache(cell.value).isEmpty &&
                  mayRead(text, sheet.name, nameChanged)
                case _ => false
              ) =>
            QualifiedRef(sheet.name, ref)
        }
      }.toSet
