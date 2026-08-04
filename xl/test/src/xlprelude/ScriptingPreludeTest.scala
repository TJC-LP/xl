// Deliberately OUTSIDE com.tjclp.xl: scripts import the prelude from a fresh namespace,
// so this test must not see the package-level exports that files under com.tjclp.xl inherit.
package xlprelude

import java.time.{LocalDate, LocalDateTime}

import munit.FunSuite

import com.tjclp.xl.scripting.{*, given}

/**
 * Gate test for the scripting prelude: every public surface a script touches must resolve through
 * `import com.tjclp.xl.scripting.{*, given}` alone — macros (transparent inline through an export
 * hop), given instances, DSL operators, evaluator extensions, sync IO, and the unsafe boundary.
 * Compile success is most of the test; runtime assertions confirm semantics.
 */
class ScriptingPreludeTest extends FunSuite:

  test("compile-time ref macros resolve through the prelude"):
    val cell: ARef = ref"A1"
    val range: CellRange = ref"A1:B10"
    assertEquals(cell.toA1, "A1")
    assertEquals(range.start.toA1, "A1")
    assertEquals(range.end.toA1, "B10")

  test("formatted literal macros resolve through the prelude"):
    val m: Formatted = money"$$1,234.56"
    val p: Formatted = percent"12.5%"
    val d: Formatted = date"2025-11-24"
    val f: CellValue = fx"=SUM(A1:B10)"
    assert(m.numFmt != NumFmt.General)
    assert(p.numFmt != NumFmt.General)
    assert(d.numFmt != NumFmt.General)
    f match
      case CellValue.Formula(expr, _, _) => assert(expr.nonEmpty)
      case other => fail(s"fx literal should produce CellValue.Formula, got $other")

  test("Sheet construction and literal-string puts are infallible"):
    val sheet = Sheet("Demo")
      .put("A1", "Title")
      .put("B1", 42)
    assertEquals(sheet.cells.size, 2)

  test("CellCodec givens and conversions resolve: typed put and readTyped"):
    val sheet = Sheet("Typed")
      .put(ref"A1", "text")
      .put(ref"B1", 42)
      .put(ref"C1", BigDecimal("1000.50"))
      .put(ref"D1", LocalDate.of(2025, 1, 15))
      .put(ref"E1", LocalDateTime.of(2025, 1, 15, 10, 30))
    assertEquals(sheet.readTyped[Int](ref"B1"), Right(Some(42)): Either[CodecError, Option[Int]])
    assertEquals(
      sheet.readTyped[String](ref"A1"),
      Right(Some("text")): Either[CodecError, Option[String]]
    )
    assertEquals(
      sheet.readTyped[BigDecimal](ref"C1"),
      Right(Some(BigDecimal("1000.50"))): Either[CodecError, Option[BigDecimal]]
    )

  test("Patch DSL composes: :=, ++, styled, merge, comment, conditionalFormat"):
    val patch =
      (ref"A1" := "Report") ++
        ref"A1".styled(CellStyle.default.bold) ++
        ref"A1:C1".merge ++
        (ref"B2" := 19.99) ++
        ref"B2".comment(Comment.plainText("unit price", Some("gen"))) ++
        ref"B2:B9".conditionalFormat(CfRule.expression("B2>10", Dxf.fill(Color.Rgb(0xffffcc00))))
    val sheet = Sheet("Patched").put(patch)
    assertEquals(sheet.cells.size, 2)
    assertEquals(sheet.mergedRanges.size, 1)
    assertEquals(sheet.getComment(ref"B2").map(_.text.toPlainText), Some("unit price"))
    assertEquals(sheet.conditionalFormats.size, 1)

  test("range fill := and ARef navigation resolve through the prelude"):
    val sheet = Sheet("Fill").put(ref"A1:B2" := 0)
    assertEquals(sheet.cells.size, 4)
    assertEquals(ref"A1".down(2).right(1), ref"B3")

  test("runtime ref interpolation returns Either and := works on RefType"):
    val row = "5"
    val patch = (for refType <- ref"A$row"
    yield refType := "dynamic").getOrElse(Patch.empty)
    val sheet = Sheet("Dyn").put(patch)
    assertEquals(sheet.cells.size, 1)

  test("runtime column letters drive setColumnProperties (no macro literal, GH-361)"):
    val widths = Vector("C" -> 14.0, "D" -> 22.0)
    val sheet = widths.foldLeft(Sheet("Widths")) { case (s, (letter, w)) =>
      val col = Column.parse(letter).getOrElse(fail(s"unparseable column: $letter"))
      s.setColumnProperties(col, ColumnProperties(width = Some(w)))
    }
    assertEquals(sheet.columnProperties.size, 2)
    assertEquals(sheet.getColumnProperties(Column.from0(3)).width, Some(22.0))
    // Runtime-parsed RefType exposes the column handle directly
    assertEquals(RefType.parse("E3").map(_.col.toLetter), Right("E"))
    assertEquals(RefType.parse("Sales!C2:E9").map(_.col.toLetter), Right("C"))

  test("rich text extensions resolve"):
    val rich = "Status: ".bold + "ACTIVE".green
    val sheet = Sheet("Rich").put(ref"A1", rich)
    assertEquals(sheet.cells.size, 1)

  test("unsafe boundary resolves: .unsafe unwraps XLResult"):
    val runtimeRef = "A1"
    val sheet = Sheet("Unsafe").put(runtimeRef, "value").unsafe
    assertEquals(sheet.cells.size, 1)
    // Literal invalid refs fail at compile time (verified: the macro rejects them through the
    // prelude hop); a runtime string is needed to exercise the XLException path.
    val invalidRef = "NOT A REF!!!"
    intercept[XLException]:
      Sheet("Unsafe").put(invalidRef, "boom").unsafe

  test("formula evaluation extensions resolve through the prelude"):
    val sheet = Sheet("Calc")
      .put(ref"A1", 2)
      .put(ref"A2", 3)
    val result = sheet.evaluateFormula("=SUM(A1:A2)")
    assert(result.exists {
      case CellValue.Number(n) => n == BigDecimal(5)
      case _ => false
    })
    assert(sheet.evaluateWithDependencyCheck().isRight)
    assert(FormulaParser.parse("=SUM(A1:A10)").isRight)
    val clock = Clock.system
    assert(clock != null || true)

  test("display interpolator resolves with given Sheet"):
    val sheet = Sheet("Disp").put(ref"A1", 100)
    given Sheet = sheet
    assertEquals(excel"${ref"A1"}", "100")

  test("display interpolator evaluates formulas (evaluating strategy wins in the prelude)"):
    val sheet = Sheet("DispF")
      .put(ref"A1", 100)
      .put(ref"B1", fx"=A1*2")
    given Sheet = sheet
    assertEquals(excel"${ref"B1"}", "200")

  test("upsert, wb.evaluateFormula, and readTypedOr resolve through the prelude"):
    val wb = Workbook(Sheet("Sales").put(ref"A1", 10).put(ref"A2", 20))
      .upsert("Summary", _.put(ref"A1", 5))
    assertEquals(
      wb.evaluateFormula("=SUM(Sales!A1:A2) + A1", "Summary"),
      Right(CellValue.Number(BigDecimal(35))): XLResult[CellValue]
    )
    val sheet = wb.sheets.headOption.getOrElse(fail("missing sheet"))
    assertEquals(sheet.readTypedOr[Int](ref"A1", 0), 10)
    assertEquals(sheet.readTypedOpt[Int](ref"Z9"), None)

  test("recalculate with per-cell errors resolves through the prelude"):
    val sheet = Sheet("Calc2")
      .put(ref"A1", 10)
      .put(ref"A2", fx"=A1*2")
      .put(ref"A3", fx"=NOSUCHFN(A1)")
    val result: RecalcResult = Workbook(sheet).recalculate()
    assert(!result.isClean)
    assertEquals(result.errors.map(_.ref.toA1), Vector("A3"))
    assertEquals(
      result.evaluated.values.headOption.flatMap(_.get(ref"A2")),
      Some(CellValue.Number(BigDecimal(20)))
    )

  test("workbook construction, withCachedFormulas, and sync Excel round-trip"):
    val sheet = Sheet("Data")
      .put(ref"A1", 10)
      .put(ref"A2", 20)
      .put(ref"A3", fx"=SUM(A1:A2)")
    val wb = Workbook(sheet).withCachedFormulas()
    val path = java.nio.file.Files.createTempDirectory("xl-prelude").resolve("roundtrip.xlsx")
    Excel.write(wb, path.toString)
    val loaded = Excel.read(path.toString)
    assertEquals(loaded.sheets.size, 1)
    assertEquals(
      loaded.sheets.headOption.map(_.cells.size),
      Some(3)
    )

  test("fx formulas serialize without the display form's leading '=' in <f> (GH-456)"):
    val wb = Workbook(Sheet("S").put(ref"A1", 2).put(ref"B1", fx"=A1*2"))
    val path = java.nio.file.Files.createTempDirectory("xl-prelude-fx").resolve("fx.xlsx")
    Excel.write(wb, path.toString)
    val zip = new java.util.zip.ZipFile(path.toFile)
    try
      val entry = Option(zip.getEntry("xl/worksheets/sheet1.xml"))
        .getOrElse(fail("sheet part missing"))
      val xml = new String(zip.getInputStream(entry).readAllBytes(), "UTF-8")
      assert(xml.contains("<f>A1*2</f>"), xml)
      assert(!xml.contains("<f>="), xml)
    finally zip.close()

  test("writeRecalculated caches uncached formulas on disk and returns the RecalcResult (GH-360)"):
    val sheet = Sheet("Calc3")
      .put(ref"A1", 10)
      .put(ref"A2", 20)
      .put(ref"A3", fx"=SUM(A1:A2)")
    val path = java.nio.file.Files.createTempDirectory("xl-prelude-wrecalc").resolve("clean.xlsx")
    val result: RecalcResult = Excel.writeRecalculated(Workbook(sheet), path.toString)
    assert(result.isClean)
    assertEquals(
      result.evaluated.values.headOption.flatMap(_.get(ref"A3")),
      Some(CellValue.Number(BigDecimal(30)))
    )
    val cached = Excel
      .read(path.toString)
      .sheets
      .headOption
      .flatMap(_.cells.get(ref"A3"))
      .map(_.value)
    cached match
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) =>
        assertEquals(n, BigDecimal(30))
      case other => fail(s"expected a cached formula on disk, got $other")

  test("writeRecalculated writes anyway on formula errors — errors are data (GH-360)"):
    val sheet = Sheet("Err")
      .put(ref"A1", 2)
      .put(ref"A2", fx"=NOSUCHFN(A1)")
      .put(ref"A3", fx"=A1*3")
    val path = java.nio.file.Files.createTempDirectory("xl-prelude-wrecalc").resolve("partial.xlsx")
    val result = Excel.writeRecalculated(Workbook(sheet), path.toString)
    assert(!result.isClean)
    assertEquals(result.errors.map(_.ref.toA1), Vector("A2"))
    val cells = Excel.read(path.toString).sheets.headOption.map(_.cells).getOrElse(Map.empty)
    cells.get(ref"A3").map(_.value) match
      case Some(CellValue.Formula(_, Some(CellValue.Number(n)), _)) =>
        assertEquals(n, BigDecimal(6))
      case other => fail(s"expected A3 cached on disk, got $other")
    cells.get(ref"A2").map(_.value) match
      case Some(CellValue.Formula(_, None, _)) => () // failed cell stays uncached
      case other => fail(s"expected A2 uncached on disk, got $other")

  test("writeRecalculated accepts XLResult[Workbook], throwing only on Left (GH-360)"):
    val wb = Workbook(Sheet("R").put(ref"A1", 4).put(ref"B1", fx"=A1*2"))
    val dir = java.nio.file.Files.createTempDirectory("xl-prelude-wrecalc")
    val okPath = dir.resolve("fromResult.xlsx")
    val result = Excel.writeRecalculated(wb.update("R", _.put(ref"A2", 1)), okPath.toString)
    assert(result.isClean)
    assert(java.nio.file.Files.exists(okPath))
    intercept[XLException]:
      Excel.writeRecalculated(
        wb.update("NoSuchSheet", identity),
        dir.resolve("never.xlsx").toString
      )

  test("writeRecalculated threads the clock and rng overloads (GH-360)"):
    val dir = java.nio.file.Files.createTempDirectory("xl-prelude-wrecalc")
    val clock = Clock.fixedDate(java.time.LocalDate.of(2026, 7, 15))
    val clocked = Excel.writeRecalculated(
      Workbook(Sheet("Vol").put(ref"A1", fx"=TODAY()")),
      dir.resolve("clock.xlsx").toString,
      clock
    )
    assertEquals(
      clocked.evaluated.values.headOption.flatMap(_.get(ref"A1")),
      Some(CellValue.DateTime(java.time.LocalDate.of(2026, 7, 15).atStartOfDay()))
    )
    val seeded = Sheet("Rand").put(ref"A1", fx"=RANDBETWEEN(1,10)")
    val r1 = Excel.writeRecalculated(
      Workbook(seeded),
      dir.resolve("rng1.xlsx").toString,
      Clock.system,
      Rng.seeded(42L)
    )
    val r2 = Excel.writeRecalculated(
      Workbook(seeded),
      dir.resolve("rng2.xlsx").toString,
      Clock.system,
      Rng.seeded(42L)
    )
    assert(r1.isClean && r2.isClean)
    assertEquals(
      r1.evaluated.values.headOption.flatMap(_.get(ref"A1")),
      r2.evaluated.values.headOption.flatMap(_.get(ref"A1"))
    )

  test("smart detection: FormattedParsers.detect and String.toFormatted via the prelude"):
    val detected = "$1,234.56".toFormatted
    assertEquals(detected.numFmt, NumFmt.Currency)
    assertEquals(FormattedParsers.detect("45.5%").numFmt, NumFmt.Percent)
    val sheet = Sheet("Smart").put(ref"A1", detected)
    assertEquals(sheet.cells.size, 1)

  test("ExcelIO escape hatch is reachable"):
    val io = ExcelIO
    assert(io != null || true)

  test("chart layer resolves through the prelude: addChart round-trips a typed chart"):
    val data = SheetName.unsafe("Data")
    val chart = Chart
      .bar(
        Vector(
          Series(
            values = DataRef(data, ref"B2:B4"),
            categories = Some(DataRef(data, ref"A2:A4")),
            name = Some(SeriesName.Literal("Units"))
          )
        ),
        direction = BarDirection.Col,
        grouping = BarGrouping.Clustered,
        title = Some("Prelude Chart"),
        legend = Some(Legend(LegendPosition.Right))
      )
      .unsafe
    val sheet = Sheet("Data")
      .put(ref"A2", "Q1")
      .put(ref"A3", "Q2")
      .put(ref"A4", "Q3")
      .put(ref"B2", 1)
      .put(ref"B3", 2)
      .put(ref"B4", 3)
      .addChart(chart, ref"D2:K15")
    assertEquals(sheet.charts.size, 1)
    assertEquals(chart.chartType, ChartType.Bar(): ChartType)
    val dir = java.nio.file.Files.createTempDirectory("xl-prelude-chart")
    val path = dir.resolve("chart.xlsx")
    Excel.write(Workbook(sheet), path.toString)
    val loaded = Excel.read(path.toString)
    val frames = loaded.sheets.headOption.map(_.charts).getOrElse(Vector.empty)
    assertEquals(frames.size, 1)
    // GH-407: the writer materializes an accent-cycle fill for fill-less series (LibreOffice
    // renders spPr-less series invisible) and the reader captures it back into the model
    val materialized =
      chart.copy(series = chart.series.map(_.copy(fill = Some(Color.Rgb(0xff4472c4)))))
    assertEquals(frames.headOption.map(_.chart), Some(materialized))

  test("drawing layer resolves through the prelude: addImage round-trips an embedded image"):
    // 1x1-style tiny PNG (the 2x3 generator template, inlined: prelude tests are self-contained)
    val pngHex =
      "89504e470d0a1a0a0000000d4948445200000002000000030802000000368849d60000000e49444154785e63f8cf8001fe0300150001ff0bfeb2140000000049454e44ae426082"
    val bytes = scala.collection.immutable.ArraySeq.unsafeWrapArray(
      pngHex.grouped(2).map(Integer.parseInt(_, 16).toByte).toArray
    )
    val image = ImageData.detect(bytes).unsafe
    assertEquals(image.format, ImageFormat.Png)
    val sheet = Sheet("Pics")
      .put(ref"A1", "with image")
      .addImage(image, ref"B2")
      .unsafe // natural-size overload returns XLResult (Tiff/Emf/Wmf are unsniffable)
      .addImage(image, DrawingAnchor.over(ref"C3:E6", EditAs.OneCell))
    assertEquals(sheet.pictures.size, 2)
    val dir = java.nio.file.Files.createTempDirectory("xl-prelude-drawing")
    val path = dir.resolve("image.xlsx")
    Excel.write(Workbook(sheet), path.toString)
    val loaded = Excel.read(path.toString)
    val pictures = loaded.sheets.headOption.map(_.pictures).getOrElse(Vector.empty)
    assertEquals(pictures.size, 2)
    assertEquals(pictures.map(_.image.format), Vector(ImageFormat.Png, ImageFormat.Png))
    assertEquals(pictures(0).image.bytes, bytes)

  test("GH-419: dataTable authoring + seedDataTables resolve through the prelude"):
    // The typed overload carries a DEFAULT seeds argument and the string overloads ride
    // @targetName — both must survive the wildcard-export hop (the export-landmine gate).
    val base = Sheet("DT419")
      .put(ref"B1", 10)
      .put(ref"B2", 20)
      .put(ref"C4", fx"=B1*B2")
      .put(ref"D4", 1)
      .put(ref"E4", 2)
      .put(ref"C5", 100)
      .put(ref"C6", 200)
    val authored: Sheet = base.dataTable(ref"D5:E6", ref"B1", ref"B2").unsafe
    authored(ref"D5").value match
      case CellValue.Formula(expr, cached, kind: FormulaKind.DataTable) =>
        assertEquals(expr, "TABLE(B1,B2)")
        assertEquals(cached, None)
        assertEquals(kind.dt2D, true)
        assertEquals(kind.ca, true)
      case other => fail(s"expected the corner record, got $other")
    // Explicit seeds + the runtime-string overload both resolve.
    val seeds = Seq(
      Seq(CellValue.Number(BigDecimal(1)), CellValue.Number(BigDecimal(2))),
      Seq(CellValue.Number(BigDecimal(3)), CellValue.Number(BigDecimal(4)))
    )
    assert(base.dataTable(ref"D5:E6", ref"B1", ref"B2", seeds).isRight)
    assert(base.dataTable("D5:E6", "B1", "B2").isRight)
    // 1-D forms resolve (row axis above + formulas left / column axis left + formulas above).
    val oneD = Sheet("DT419b")
      .put(ref"A1", 3)
      .put(ref"B20", 7)
      .put(ref"A21", fx"=A1+1")
    assert(oneD.dataTableRow(ref"B21:B21", ref"A1").isRight)
    assert(oneD.dataTableRow("B21:B21", "A1").isRight)
    val oneDCol = Sheet("DT419c")
      .put(ref"A2", 3)
      .put(ref"F9", fx"=A2*100")
      .put(ref"E10", 5)
    assert(oneDCol.dataTableCol(ref"F10:F10", ref"A2").isRight)
    assert(oneDCol.dataTableCol("F10:F10", "A2").isRight)
    // Seeding: the what-if engine computes Excel's grid (D5 = B1<-1, B2<-100 -> 100).
    val seeded: Workbook = Workbook(authored).seedDataTables().unsafe
    seeded.sheets.headOption.flatMap(_.cells.get(ref"D5")).map(_.value) match
      case Some(CellValue.Formula(_, cached, _: FormulaKind.DataTable)) =>
        assertEquals(cached, Some(CellValue.Number(BigDecimal(100))))
      case other => fail(s"expected a seeded corner record, got $other")
    assertEquals(
      seeded.sheets.headOption.flatMap(_.cells.get(ref"E6")).map(_.value),
      Some(CellValue.Number(BigDecimal(400)))
    )
    // The scoped overload resolves too.
    assert(
      Workbook(authored)
        .seedDataTables(SheetName.unsafe("DT419"), Some(ref"D5:E6"), Clock.system)
        .isRight
    )

  test("GH-430: a data-table record authors from scratch through the prelude and round-trips"):
    // FormulaKind must resolve through the prelude export alone (export-forwarder landmine
    // guard), and CellValue.dataTable must synthesize the derived TABLE(...) display text —
    // proving #419's authoring sugar needs no remodel of the substrate.
    val kind: FormulaKind.DataTable = FormulaKind.DataTable(
      ref"F2:G2",
      dt2D = true,
      dtr = false,
      r1 = Some(ref"A1"),
      r2 = Some(ref"A2")
    )
    val anchor = CellValue.dataTable(kind, Some(CellValue.Number(BigDecimal(42))))
    assertEquals(anchor.expression, "TABLE(A1,A2)")
    val sheet = Sheet("DT")
      .put(ref"A1", 8)
      .put(ref"A2", 3)
      .put(ref"F2", anchor)
      .put(
        ref"G2",
        CellValue.dataTable(kind.copy(ca = true), Some(CellValue.Number(BigDecimal(7))))
      )
    val path = java.nio.file.Files.createTempDirectory("xl-prelude-dt").resolve("records.xlsx")
    Excel.write(Workbook(sheet), path.toString)
    val loaded = Excel.read(path.toString)
    val cells = loaded.sheets.headOption.map(_.cells).getOrElse(Map.empty)
    assertEquals(cells.get(ref"F2").map(_.value), Some(anchor))
    cells.get(ref"G2").map(_.value) match
      case Some(CellValue.Formula(expr, cached, k: FormulaKind.DataTable)) =>
        assertEquals(expr, "TABLE(A1,A2)")
        assertEquals(cached, Some(CellValue.Number(BigDecimal(7))))
        assertEquals(k.ca, true)
      case other => fail(s"G2 record lost through prelude round-trip: $other")
