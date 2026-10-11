package com.tjclp.xl.ooxml

import scala.xml.*

import com.tjclp.xl.cells.ArrayMode

/**
 * The workbook's cell-metadata part, `xl/metadata.xml` (GH-714), as far as dynamic arrays need it.
 *
 * Excel 365 marks a dynamic-array anchor with the cell attribute `cm="n"`: the 1-based index of a
 * `cellMetadata/bk` whose `rc` names the `XLDAPR` metadata type (`t`, 1-based into `metadataTypes`)
 * and a block of that type's `futureMetadata` (`v`, 0-based), which carries
 * `xda:dynamicArrayProperties fDynamic="1"`. The same part also holds rich-value metadata (images
 * in cells, `vm`), so the writer never rewrites what it finds: it reuses a matching block, appends
 * the missing ones (existing indices never move), or generates the part when the package has none.
 * Everything here is pure and total — a malformed part reads as "no dynamic arrays" and plans as
 * untouched.
 */
private[xl] object SheetMetadataPart:

  val relType: String = XmlUtil.relTypeSheetMetadata
  val contentType: String = XmlUtil.ctSheetMetadata

  /** Where a generated part goes (Excel's own name). */
  val defaultPath: String = "xl/metadata.xml"

  private val nsXda = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
  private val xldapr = "XLDAPR"
  private val extUri = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"

  /** Excel's XLDAPR metadataType attributes, in Excel's order. */
  private val xldaprTypeAttrs: List[(String, String)] = List(
    "name" -> xldapr,
    "minSupportedVersion" -> "120000",
    "copy" -> "1",
    "pasteAll" -> "1",
    "pasteValues" -> "1",
    "merge" -> "1",
    "splitFirst" -> "1",
    "rowColShift" -> "1",
    "clearFormats" -> "1",
    "clearComments" -> "1",
    "assign" -> "1",
    "coerce" -> "1",
    "cellMeta" -> "1"
  )

  /** The part Excel writes for one plain dynamic array, byte for byte. */
  val canonicalXml: String = generate(Vector(ArrayMode.Dynamic()))

  /** The fixed order variants are allocated in, so a plan never depends on cell order. */
  private val variantOrder: Vector[ArrayMode.Dynamic] =
    Vector(ArrayMode.Dynamic(), ArrayMode.Dynamic(collapsed = true))

  /** What the writer does with the part. */
  enum MetadataOutput derives CanEqual:
    /** Ship the source part as it is (or no part at all when none is needed). */
    case Untouched

    /** No source part: write this one, registering its Override and relationship. */
    case Generated(xml: String)

    /** The source part lacks a needed block: replace its bytes with these (indices kept). */
    case Appended(xml: String)

  /** The `cm` index each needed dynamic variant is written with, and the part to ship. */
  final case class CellMetadataPlan(
    cmOf: Map[ArrayMode.Dynamic, Int],
    output: MetadataOutput
  ):
    def lookup: ArrayMode.Dynamic => Option[Int] = cmOf.get

  object CellMetadataPlan:
    val none: CellMetadataPlan = CellMetadataPlan(Map.empty, MetadataOutput.Untouched)

  /** The package path of the metadata part, found by relationship type (never by name). */
  def locate(workbookRels: Relationships): Option[String] =
    workbookRels.relationships
      .find(_.`type` == relType)
      .map(rel => Relationships.resolveWorkbookTarget(rel.target))

  /** Resolve every dynamic-array `cm` index of a part; malformed XML reads as none. */
  def parse(xml: String): CellMetadataIndex =
    XmlSecurity.parseSafe(xml, defaultPath).toOption match
      case Some(root) => index(root)
      case None => CellMetadataIndex.empty

  private def children(e: Elem, label: String): Vector[Elem] =
    e.child.collect { case c: Elem if c.label == label => c }.toVector

  private def attr(e: Elem, name: String): Option[String] =
    e.attributes.collectFirst {
      case a: UnprefixedAttribute if a.key == name => a.value.text
    }

  /** 1-based positions of the XLDAPR metadataType. */
  private def xldaprTypeIndices(root: Elem): Set[Int] =
    children(root, "metadataTypes")
      .take(1)
      .flatMap(children(_, "metadataType"))
      .zipWithIndex
      .collect { case (t, i) if attr(t, "name").contains(xldapr) => i + 1 }
      .toSet

  /** The XLDAPR future-metadata blocks, 0-based, as the dynamic mode each one describes. */
  private def xldaprBlocks(root: Elem): Vector[Option[ArrayMode.Dynamic]] =
    children(root, "futureMetadata")
      .filter(fm => attr(fm, "name").contains(xldapr))
      .take(1)
      .flatMap(children(_, "bk"))
      .map { bk =>
        (bk \\ "dynamicArrayProperties").collectFirst { case p: Elem => p }.flatMap { p =>
          val dynamic = attr(p, "fDynamic").exists(FormulaKindCodec.parseBoolean)
          val collapsed = attr(p, "fCollapsed").exists(FormulaKindCodec.parseBoolean)
          Option.when(dynamic)(ArrayMode.Dynamic(collapsed))
        }
      }

  private def cellBlocks(root: Elem): Vector[Elem] =
    children(root, "cellMetadata").take(1).flatMap(children(_, "bk"))

  private def index(root: Elem): CellMetadataIndex =
    val types = xldaprTypeIndices(root)
    val blocks = xldaprBlocks(root)
    val byCm = cellBlocks(root).zipWithIndex.flatMap { case (bk, i) =>
      children(bk, "rc").headOption.flatMap { rc =>
        for
          t <- attr(rc, "t").flatMap(_.trim.toIntOption)
          if types.contains(t)
          v <- attr(rc, "v").flatMap(_.trim.toIntOption)
          mode <- blocks.lift(v).flatten
        yield (i + 1) -> mode
      }
    }
    CellMetadataIndex(byCm.toMap)

  /**
   * Plan the part for a write whose regenerated sheets hold dynamic records of the `needed`
   * variants. A pure function of the source bytes and the set: never of cell order.
   */
  def plan(source: Option[String], needed: Set[ArrayMode.Dynamic]): CellMetadataPlan =
    val ordered = variantOrder.filter(needed.contains)
    if ordered.isEmpty then CellMetadataPlan.none
    else
      source match
        case None =>
          CellMetadataPlan(
            ordered.zipWithIndex.map { case (m, i) => m -> (i + 1) }.toMap,
            MetadataOutput.Generated(generate(ordered))
          )
        case Some(xml) =>
          XmlSecurity.parseSafe(xml, defaultPath).toOption match
            // an unreadable part cannot be extended: write the records without `cm`
            case None => CellMetadataPlan.none
            case Some(root) =>
              val existing = index(root).byCm.toVector.sortBy(_._1)
              def firstCm(m: ArrayMode.Dynamic): Option[Int] =
                existing.collectFirst { case (cm, `m`) => cm }
              val missing = ordered.filter(firstCm(_).isEmpty)
              if missing.isEmpty then
                CellMetadataPlan(
                  ordered.flatMap(m => firstCm(m).map(m -> _)).toMap,
                  MetadataOutput.Untouched
                )
              else
                val (appended, newCms) = append(root, missing)
                CellMetadataPlan(
                  (ordered.flatMap(m => firstCm(m).map(m -> _)) ++ newCms).toMap,
                  MetadataOutput.Appended(XmlUtil.compact(appended))
                )

  // ---------------------------------------------------------------------------------------------
  // Generation and append-only rewrite

  private def metaData(pairs: List[(String, String)]): MetaData =
    pairs.foldRight(Null: MetaData) { case ((k, v), next) => new UnprefixedAttribute(k, v, next) }

  private def el(label: String, scope: NamespaceBinding, attrs: List[(String, String)])(
    kids: Node*
  ): Elem =
    Elem(null, label, metaData(attrs), scope, true, kids*)

  /** The XLDAPR future-metadata block of one variant; declares `xda` unless `scope` binds it. */
  private def futureBlock(m: ArrayMode.Dynamic, scope: NamespaceBinding): Elem =
    val xdaScope =
      if scope.getURI("xda") == nsXda then scope else NamespaceBinding("xda", nsXda, scope)
    val props = Elem(
      "xda",
      "dynamicArrayProperties",
      metaData(List("fDynamic" -> "1", "fCollapsed" -> (if m.collapsed then "1" else "0"))),
      xdaScope,
      true
    )
    el("bk", scope, Nil)(el("extLst", scope, Nil)(el("ext", scope, List("uri" -> extUri))(props)))

  private def cellBlock(t: Int, v: Int, scope: NamespaceBinding): Elem =
    el("bk", scope, Nil)(el("rc", scope, List("t" -> t.toString, "v" -> v.toString))())

  /** A fresh part: XLDAPR type 1, one future block and one cell block per variant, in order. */
  private def generate(ordered: Vector[ArrayMode.Dynamic]): String =
    val scope =
      NamespaceBinding(null, XmlUtil.nsSpreadsheetML, NamespaceBinding("xda", nsXda, TopScope))
    val root = el("metadata", scope, Nil)(
      el("metadataTypes", scope, List("count" -> "1"))(
        el("metadataType", scope, xldaprTypeAttrs)()
      ),
      el("futureMetadata", scope, List("name" -> xldapr, "count" -> ordered.size.toString))(
        ordered.map(futureBlock(_, scope))*
      ),
      el("cellMetadata", scope, List("count" -> ordered.size.toString))(
        ordered.indices.map(i => cellBlock(1, i, scope))*
      )
    )
    XmlUtil.compact(root)

  /** Replace (or add) one unprefixed attribute, keeping every other attribute in place. */
  private def withAttr(e: Elem, key: String, value: String): Elem =
    val pairs = e.attributes.toList
    val replaced =
      if pairs.exists { case a: UnprefixedAttribute => a.key == key; case _ => false } then
        pairs.map {
          case a: UnprefixedAttribute if a.key == key => new UnprefixedAttribute(key, value, Null)
          case other => other
        }
      else pairs :+ new UnprefixedAttribute(key, value, Null)
    val md = replaced.foldRight(Null: MetaData) {
      case (a: PrefixedAttribute, next) => new PrefixedAttribute(a.pre, a.key, a.value, next)
      case (a: UnprefixedAttribute, next) => new UnprefixedAttribute(a.key, a.value, next)
      case (_, next) => next
    }
    e.copy(attributes = md)

  private def appendChildren(e: Elem, label: String, extra: Seq[Elem]): Elem =
    val updated = e.copy(child = e.child ++ extra)
    withAttr(updated, "count", children(updated, label).size.toString)

  /** CT_Metadata child order. */
  private val childOrder: Vector[String] = Vector(
    "metadataTypes",
    "metadataStrings",
    "mdxMetadata",
    "futureMetadata",
    "cellMetadata",
    "valueMetadata",
    "extLst"
  )

  private def rank(n: Node): Int = n match
    case e: Elem =>
      childOrder.indexOf(e.label) match
        case -1 => childOrder.size
        case i => i
    case _ => -1

  /** Insert `child` after the last existing element that sorts at or before it. */
  private def insertInOrder(root: Elem, child: Elem): Elem =
    val r = rank(child)
    val kids = root.child
    val at = kids.lastIndexWhere {
      case k: Elem => rank(k) <= r
      case _ => false
    } + 1
    root.copy(child = (kids.take(at) :+ child) ++ kids.drop(at))

  /** Replace the first child element `matches` selects. */
  private def replaceChild(root: Elem, matches: Elem => Boolean, f: Elem => Elem): Elem =
    val idx = root.child.indexWhere {
      case e: Elem => matches(e)
      case _ => false
    }
    root.child.lift(idx) match
      case Some(e: Elem) => root.copy(child = root.child.updated(idx, f(e)))
      case _ => root

  /**
   * Append the blocks for `missing` (in order) and return the new root and each variant's new `cm`.
   * Existing types, blocks and indices are never moved.
   */
  private def append(
    root0: Elem,
    missing: Vector[ArrayMode.Dynamic]
  ): (Elem, Vector[(ArrayMode.Dynamic, Int)]) =
    val scope = root0.scope
    // 1. the XLDAPR metadata type
    val (root1, typeIdx) = xldaprTypeIndices(root0).minOption match
      case Some(t) => (root0, t)
      case None =>
        val typeElem = el("metadataType", scope, xldaprTypeAttrs)()
        children(root0, "metadataTypes").headOption match
          case Some(_) =>
            val r = replaceChild(
              root0,
              _.label == "metadataTypes",
              appendChildren(_, "metadataType", Seq(typeElem))
            )
            (r, children(children(r, "metadataTypes").headOption.getOrElse(r), "metadataType").size)
          case None =>
            (insertInOrder(root0, el("metadataTypes", scope, List("count" -> "1"))(typeElem)), 1)
    // 2. the XLDAPR future-metadata blocks
    val firstV = xldaprBlocks(root1).size
    val blocks = missing.map(futureBlock(_, scope))
    val isXldapr: Elem => Boolean = e =>
      e.label == "futureMetadata" && attr(e, "name").contains(xldapr)
    val root2 =
      if children(root1, "futureMetadata").exists(isXldapr) then
        replaceChild(root1, isXldapr, appendChildren(_, "bk", blocks))
      else
        val fm = el("futureMetadata", scope, List("name" -> xldapr, "count" -> "0"))()
        insertInOrder(root1, appendChildren(fm, "bk", blocks))
    // 3. the cell-metadata blocks pointing at them
    val firstCm = cellBlocks(root2).size + 1
    val cells = missing.indices.map(i => cellBlock(typeIdx, firstV + i, scope))
    val root3 =
      if children(root2, "cellMetadata").nonEmpty then
        replaceChild(root2, _.label == "cellMetadata", appendChildren(_, "bk", cells))
      else
        insertInOrder(
          root2,
          appendChildren(el("cellMetadata", scope, List("count" -> "0"))(), "bk", cells)
        )
    (root3, missing.zipWithIndex.map { case (m, i) => m -> (firstCm + i) })
