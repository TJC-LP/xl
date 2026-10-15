package com.tjclp.xl.formula.perf

import java.nio.file.{Files, Path}

/**
 * Manual benchmark entry point (#767):
 *
 * {{{
 * ./mill xl-evaluator.perf.runMain com.tjclp.xl.formula.perf.LoadsBookMain shape=H10 mode=recalc
 * XL_PERF_HEAP=4g ./mill xl-evaluator.perf.runMain com.tjclp.xl.formula.perf.LoadsBookMain shape=Real mode=checked
 * }}}
 *
 * Arguments are `key=value`:
 *   - `shape`: `Tiny` (default), `H10`, `R20`, `Real`, or `RunTotal` (with `rows`)
 *   - `rows`, `aggs`, `rh`: override the preset's sizes
 *   - `crit`: `helper` or `raw` criteria; `raw` also drops the helper columns (`Real crit=raw` is
 *     the real shape with raw criteria)
 *   - `mode`: `recalc` (default), `parallel`, `uncached`, `checked` (writes the file, timing
 *     `Write`)
 *   - `threads`: worker count for `parallel` (default 2)
 *   - `out`: where `checked` writes (default: a deleted temporary file)
 *   - `oracle`: write the shape's 40 sample cells as `ref<TAB>value` lines to this path
 *   - `index`: `on` (default) or `off`, the criteria index
 *
 * Prints the build time, the counter and phase table, and one machine-readable `RESULT` line.
 */
object LoadsBookMain:

  def main(args: Array[String]): Unit =
    val kv = args.iterator
      .map(_.split("=", 2))
      .collect { case Array(k, v) => k -> v }
      .toMap
    def int(key: String): Option[Int] = kv.get(key).flatMap(_.toIntOption)
    // an unknown shape or mode stops the run: a silent default would print a RESULT line labelled
    // with the requested name for a run of a different book
    def refuse(message: String): Nothing =
      System.err.println(message)
      sys.exit(2)
    val mode = kv.get("mode") match
      case None => LoadsBookHarness.Mode.Recalc
      case Some(m) =>
        LoadsBookHarness.Mode
          .parse(m)
          .getOrElse(refuse(s"unknown mode=$m; one of recalc, parallel, uncached, checked"))
    val threads = int("threads").getOrElse(2)
    val out = kv.get("out").map(Path.of(_))
    val shapeName = kv.getOrElse("shape", "Tiny")
    val criteriaIndex = kv.get("index") match
      case None | Some("on") => true
      case Some("off") => false
      case Some(other) => refuse(s"unknown index=$other; on or off")

    val buildStart = System.nanoTime()
    val (label, workbook, built) =
      if shapeName.equalsIgnoreCase("RunTotal") then
        val rows = int("rows").getOrElse(2000)
        (s"RunTotal rows=$rows", LoadsBook.runTotal(rows), None)
      else
        val preset = LoadsBook
          .preset(shapeName)
          .getOrElse(refuse(s"unknown shape=$shapeName; one of Tiny, H10, R20, Real, RunTotal"))
        val crit = kv.get("crit") match
          case None => preset.crit
          case Some(c) =>
            LoadsBook.Crit.values
              .find(_.toString.equalsIgnoreCase(c))
              .getOrElse(refuse(s"unknown crit=$c; helper or raw"))
        val shape = preset.copy(
          rows = int("rows").getOrElse(preset.rows),
          aggs = int("aggs").getOrElse(preset.aggs),
          rh = int("rh").getOrElse(preset.rh),
          crit = crit,
          helpers = preset.helpers && crit == LoadsBook.Crit.Helper
        )
        val b = LoadsBook.build(shape)
        (
          s"${shape.name} rows=${shape.rows} aggs=${shape.aggs} rh=${shape.rh} " +
            s"helpers=${shape.helpers} crit=${shape.crit} formulas=${b.totalFormulas}",
          b.workbook,
          Some(b)
        )
    val buildSeconds = (System.nanoTime() - buildStart).toDouble / 1e9
    System.err.println(f"built $label in $buildSeconds%.2f s")

    val run =
      LoadsBookHarness.run(
        workbook,
        mode,
        threads = threads,
        out = out,
        criteriaIndex = criteriaIndex
      )
    println(LoadsBookHarness.report(s"$label mode=$mode index=${kv.getOrElse("index", "on")}", run))

    for
      path <- kv.get("oracle")
      b <- built
    do
      val lines = LoadsBookHarness.oracleLines(b, run.result.workbook)
      Files.writeString(Path.of(path), lines.mkString("", "\n", "\n"))
      System.err.println(s"wrote ${lines.size} oracle samples to $path")

    val snap = run.counters
    val fields = Vector(
      s"shape=$shapeName",
      s"mode=$mode",
      s"index=${if criteriaIndex then "on" else "off"}",
      f"build=$buildSeconds%.2f",
      f"wall=${run.wallSeconds}%.2f",
      f"heapPeakGB=${run.heapPeakBytes.toDouble / 1e9}%.2f",
      s"errors=${run.result.errors.size}"
    ) ++ snap.render.linesIterator
    println(("RESULT" +: fields).mkString(" "))
