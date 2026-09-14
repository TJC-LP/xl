package com.tjclp.xl.formula.graph

import com.tjclp.xl.addressing.SheetName
import com.tjclp.xl.cells.{CellValue, FormulaKind}
import com.tjclp.xl.formula.ast.TExpr
import com.tjclp.xl.formula.eval.{Evaluator, SheetEvaluator}
import com.tjclp.xl.formula.graph.DependencyGraph.QualifiedRef
import com.tjclp.xl.formula.parser.FormulaParser
import com.tjclp.xl.workbooks.{DefinedName, Workbook}

/** Formula roots invalidated by a change to the defined-name table, including aliases and scope. */
private[xl] object NameChanges:
  def readers(before: Workbook, after: Workbook): Set[QualifiedRef] =
    if before.metadata.definedNames == after.metadata.definedNames then Set.empty
    else
      type NameKey = (SheetName, SheetName, String)
      val canonicalBefore = DependencyGraph.sheetCanonicaliser(before)
      val canonicalAfter = DependencyGraph.sheetCanonicaliser(after)
      val positionsBefore =
        before.sheets.zipWithIndex.reverseIterator.map((s, i) => s.name -> i).toMap
      val positionsAfter =
        after.sheets.zipWithIndex.reverseIterator.map((s, i) => s.name -> i).toMap
      // Invocation-local memoization: many formulas share a name/alias chain. No state escapes.
      val memo = scala.collection.mutable.HashMap.empty[NameKey, Boolean]
      val parsed = scala.collection.mutable.HashMap.empty[String, Option[TExpr[?]]]

      def binding(wb: Workbook, dn: DefinedName): (String, Option[SheetName]) =
        (dn.formula, Evaluator.definedNameScope(wb, dn).map(_.name))

      def expressionChanged(text: String, current: SheetName, visiting: Set[NameKey]): Boolean =
        val matches = (name: String, lookup: SheetName, fallback: SheetName) =>
          nameChanged(name, lookup, fallback, visiting)
        parsed.getOrElseUpdate(text, FormulaParser.parse(text).toOption) match
          case Some(expr) =>
            DependencyGraph.referencesMatching(
              expr,
              (name, scope) => matches(name, scope.getOrElse(current), current),
              includeDynamicCalls = true
            )
          case None => ReferenceScan.mayReferenceName(after, current, text, matches)

      def nameChanged(
        name: String,
        lookupFrom: SheetName,
        fallback: SheetName,
        visiting: Set[NameKey]
      ): Boolean =
        val lookup = canonicalAfter(lookupFrom)
        val key = (lookup, fallback, name)
        memo.get(key) match
          case Some(changed) => changed
          // Cycles/depth limits cannot certify independence. Conservatively invalidate.
          case None if visiting.contains(key) || visiting.size >= 100 => true
          case None =>
            val old = positionsBefore
              .get(canonicalBefore(lookupFrom))
              .flatMap(i => Evaluator.lookupDefinedNameAt(before, Some(i), name))
            val fresh = positionsAfter
              .get(lookup)
              .flatMap(i => Evaluator.lookupDefinedNameAt(after, Some(i), name))
            val changed =
              old.map(binding(before, _)) != fresh.map(binding(after, _)) || fresh.exists { dn =>
                val current = Evaluator.definedNameScope(after, dn).map(_.name).getOrElse(fallback)
                expressionChanged(dn.formula, current, visiting + key)
              }
            memo(key) = changed
            changed

      after.sheets.iterator.flatMap { sheet =>
        sheet.cells.iterator.collect {
          case (ref, cell) if (cell.value match
                case CellValue.Formula(_, _, _: FormulaKind.DataTable) => false
                case CellValue.Formula(text, _, _) =>
                  SheetEvaluator.pinnedExternalCache(cell.value).isEmpty &&
                  expressionChanged(text, sheet.name, Set.empty)
                case _ => false
              ) =>
            QualifiedRef(sheet.name, ref)
        }
      }.toSet
