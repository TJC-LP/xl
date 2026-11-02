package com.tjclp.xl.ooxml

import com.tjclp.xl.workbooks.DefinedName

/** Reportable non-worksheet kind; the model still contains worksheets only (GH-788). */
enum PreservedSheetKind derives CanEqual:
  case Macro, InternationalMacro, Chart, Dialog

  def slug: String = this match
    case Macro => "macro"
    case InternationalMacro => "international-macro"
    case Chart => "chart"
    case Dialog => "dialog"

  def isMacro: Boolean = this == Macro || this == InternationalMacro

/**
 * The `<sheet>` kinds of `xl/workbook.xml` that are not worksheets (GH-785).
 *
 * A `<sheets>` entry names its part through a workbook relationship, and the relationship type says
 * what the part is. Chartsheets and dialog sheets (ECMA-376 Part 1, §12.3) and Excel 4.0 macro
 * sheets, plain and international (the `xlMacrosheet` and `xlIntlMacrosheet` relationships of
 * [MS-XLSX]), sit in `<sheets>` next to worksheets but hold no worksheet. The domain model holds
 * worksheets only: these entries, their parts and their relationships are kept byte for byte by the
 * writer, and every index the model exposes (sheet position, `localSheetId`, active tab) counts
 * worksheets only. `localSheetId` and `activeTab` index `<sheets>` in the file (ECMA-376 Part 1,
 * §18.2.5, §18.2.30), so both are renumbered at the boundary.
 */
object SheetKinds:

  val relTypeIntlMacrosheet: String =
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet"

  private val preservedKinds: Map[String, PreservedSheetKind] = Map(
    XmlUtil.relTypeChartsheet -> PreservedSheetKind.Chart,
    XmlUtil.relTypeDialogsheet -> PreservedSheetKind.Dialog,
    XmlUtil.relTypeMacrosheet -> PreservedSheetKind.Macro,
    relTypeIntlMacrosheet -> PreservedSheetKind.InternationalMacro,
    // ISO/IEC 29500 Strict spellings of the two standard kinds
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet" -> PreservedSheetKind.Chart,
    "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet" -> PreservedSheetKind.Dialog
  )

  /** Relationship types of `<sheet>` targets that are not worksheets. */
  val nonWorksheetRelTypes: Set[String] = preservedKinds.keySet

  /**
   * Whether a `<sheet>` whose `r:id` resolves to a relationship of `relType` is a worksheet. An
   * entry with no relationship stays a worksheet (the `sheet<sheetId>.xml` fallback).
   */
  def isWorksheet(relType: Option[String]): Boolean =
    !relType.exists(nonWorksheetRelTypes.contains)

  /** Stable reporting kind; does not make the entry addressable as a worksheet (GH-788). */
  def preservedKind(relType: String): Option[PreservedSheetKind] = preservedKinds.get(relType)

  /** Per `<sheet>` entry, in `<sheets>` order: whether it is a worksheet. */
  def worksheetMask(refs: Seq[SheetRef], rels: Relationships): Vector[Boolean] =
    val typeById = rels.relationships.iterator.map(rel => rel.id -> rel.`type`).toMap
    refs.toVector.map(ref => isWorksheet(typeById.get(ref.relationshipId)))

/**
 * The read-side map from `<sheets>` positions to model (worksheet-only) positions (GH-785).
 *
 * `mask(i)` is true when `<sheets>` entry `i` is a worksheet.
 */
final case class SourceSheetIndex(mask: Vector[Boolean]):

  private val nonWorksheetCount: Int = mask.count(!_)

  /** `rank(i)`: the number of worksheets before `<sheets>` entry `i`. */
  private val rank: Vector[Int] = mask.scanLeft(0)((n, isWs) => if isWs then n + 1 else n)

  /** True when every entry is a worksheet: the maps are then the identity. */
  def allWorksheets: Boolean = nonWorksheetCount == 0

  /**
   * The model position of `<sheets>` entry `idx`: its rank among worksheets; None for a
   * non-worksheet entry. An index past the end (a malformed `localSheetId`) stays past the end by
   * the same distance, so the writer maps it back to the value it was.
   */
  def toModel(idx: Int): Option[Int] =
    if idx < 0 then Some(idx)
    else if idx >= mask.size then Some(idx - nonWorksheetCount)
    else if mask(idx) then Some(rank(idx))
    else None

  /**
   * The model's active sheet for a source `activeTab`: the tab itself when it is a worksheet, else
   * the first worksheet after it (the caller clamps past the end).
   */
  def activeToModel(idx: Int): Int =
    toModel(idx).getOrElse(rank(idx))

  /**
   * Defined names in model positions: a `localSheetId` naming a worksheet is renumbered to its
   * model position; a name scoped to a non-worksheet entry is left out (the writer restores it from
   * the preserved `<definedNames>`).
   */
  def namesToModel(names: Vector[DefinedName]): Vector[DefinedName] =
    if allWorksheets then names
    else
      names.flatMap { dn =>
        dn.localSheetId match
          case None => Some(dn)
          case Some(idx) => toModel(idx).map(m => dn.copy(localSheetId = Some(m)))
      }

/**
 * The write-side `<sheets>` order (GH-785): the model's worksheets with the source's non-worksheet
 * entries kept in place.
 *
 * Each kept entry follows the worksheet that preceded it in the source; when that worksheet is gone
 * it follows the nearest earlier source worksheet still present, and with none it goes first.
 * Worksheets are identified by their source part path, so renames and reorders carry the kept
 * entries with them.
 */
final case class SheetLayout(
  slots: Vector[SheetLayout.Slot],
  source: SourceSheetIndex,
  modelOfSource: Map[Int, Int] = Map.empty
):
  import SheetLayout.Slot

  private val modelCount: Int = slots.count {
    case Slot.Model(_) => true
    case Slot.Kept(_, _) => false
  }

  private val positionOfModel: Map[Int, Int] =
    slots.zipWithIndex.collect { case (Slot.Model(m), pos) => m -> pos }.toMap

  private val positionOfSource: Map[Int, Int] =
    slots.zipWithIndex.collect { case (Slot.Kept(_, src), pos) => src -> pos }.toMap

  /** The kept non-worksheet entries, in output order. */
  def kept: Vector[SheetRef] = slots.collect { case Slot.Kept(ref, _) => ref }

  /** The `<sheets>` position of model sheet `m`; past the model's end stays past the end. */
  def toOutput(m: Int): Int =
    positionOfModel.getOrElse(m, if m < 0 then m else m - modelCount + slots.size)

  /** The model's sheet refs (one per model sheet, in model order) with the kept entries placed. */
  def interleave(modelRefs: Vector[SheetRef]): Vector[SheetRef] =
    slots.flatMap {
      case Slot.Model(m) => modelRefs.lift(m)
      case Slot.Kept(ref, _) => Some(ref)
    }

  /**
   * The `<definedNames>` content in output positions: the model's names renumbered, followed by the
   * source's names scoped to a kept entry, renumbered to where that entry now sits. The text of a
   * name scoped to a kept entry is written as the source has it: a worksheet it refers to that was
   * renamed or deleted is not rewritten.
   */
  def namesToOutput(
    model: Vector[DefinedName],
    preserved: Vector[DefinedName]
  ): Vector[DefinedName] =
    if kept.isEmpty then model
    else
      val modelNames = model.map(dn => dn.copy(localSheetId = dn.localSheetId.map(toOutput)))
      val keptNames = preserved.flatMap { dn =>
        dn.localSheetId.flatMap(positionOfSource.get).map(p => dn.copy(localSheetId = Some(p)))
      }
      modelNames ++ keptNames

  /**
   * The output `activeTab`. A source tab on a kept entry stays on that entry while the model's
   * active sheet is still the worksheet the reader made active for it (the next worksheet, else the
   * last), found by its source part so a reorder keeps the match; otherwise the model's sheet wins.
   * The model cannot tell an unchanged active sheet from one selected again, so selecting that
   * worksheet explicitly also keeps the tab on the kept entry.
   */
  def activeTab(modelActive: Int, sourceActive: Option[Int]): Int =
    val clampedModel = OoxmlWorkbook.clampActiveTab(modelActive, modelCount)
    sourceActive
      .flatMap { src =>
        val mask = source.mask
        val derivedSource =
          (src + 1 until mask.size).find(mask).orElse((src - 1 to 0 by -1).find(mask))
        positionOfSource
          .get(src)
          .filter(_ => derivedSource.flatMap(modelOfSource.get).contains(clampedModel))
      }
      .getOrElse(toOutput(clampedModel))

object SheetLayout:

  enum Slot:
    /** Model sheet at this model position. */
    case Model(index: Int)

    /** A non-worksheet entry kept verbatim, with its `<sheets>` position in the source. */
    case Kept(ref: SheetRef, sourceIndex: Int)

  /** The layout of a write with no kept entries: model order, identity index maps. */
  def identity(modelCount: Int): SheetLayout =
    SheetLayout(
      Vector.tabulate(modelCount)(Slot.Model(_)),
      SourceSheetIndex(Vector.fill(modelCount)(true))
    )

  /**
   * Place the source's non-worksheet entries among the model's sheets.
   *
   * @param sourceRefs
   *   the source `<sheets>` entries
   * @param sourceRels
   *   the source workbook relationships (kinds and worksheet part paths)
   * @param modelSourcePaths
   *   per model sheet, its source worksheet part path (None = a sheet the source did not have)
   */
  def plan(
    sourceRefs: Seq[SheetRef],
    sourceRels: Relationships,
    modelSourcePaths: Vector[Option[String]]
  ): SheetLayout =
    val refs = sourceRefs.toVector
    val mask = SheetKinds.worksheetMask(refs, sourceRels)
    if SourceSheetIndex(mask).allWorksheets then
      SheetLayout(Vector.tabulate(modelSourcePaths.size)(Slot.Model(_)), SourceSheetIndex(mask))
    else
      val targetById = sourceRels.relationships.iterator.map(rel => rel.id -> rel.target).toMap
      def sourcePath(ref: SheetRef): String =
        targetById
          .get(ref.relationshipId)
          .map(Relationships.resolveWorkbookTarget)
          .getOrElse(s"xl/worksheets/sheet${ref.sheetId}.xml")
      val paths: Vector[String] = refs.map(sourcePath)
      val modelIndexByPath: Map[String, Int] =
        modelSourcePaths.zipWithIndex.collect { case (Some(path), m) => path -> m }.toMap

      // Per kept entry: the model sheet it follows (-1 = first), from the nearest earlier source
      // worksheet still present in the model.
      val anchored: Vector[(Int, Slot.Kept)] = refs.indices.toVector.collect {
        case i if !mask(i) =>
          val anchor = (i - 1 to 0 by -1).iterator
            .filter(mask)
            .flatMap(j => modelIndexByPath.get(paths(j)))
            .nextOption()
            .getOrElse(-1)
          anchor -> Slot.Kept(refs(i), i)
      }
      def keptAfter(m: Int): Vector[Slot] = anchored.collect { case (a, k) if a == m => k }
      val slots = keptAfter(-1) ++ modelSourcePaths.indices.toVector.flatMap(m =>
        Slot.Model(m) +: keptAfter(m)
      )
      val modelOfSource: Map[Int, Int] =
        paths.indices.filter(mask).flatMap(j => modelIndexByPath.get(paths(j)).map(j -> _)).toMap
      SheetLayout(slots, SourceSheetIndex(mask), modelOfSource)
