# Effect System: Ox vs Cats Effect vs ZIO

**Status**: Evaluation, 2026-09-07. Proposes ADR-017 (see `decisions.md`); nothing here is
implemented yet.

**Question**: How well would [Ox](https://github.com/softwaremill/ox) (SoftwareMill's direct-style
concurrency library) fit xl in place of Cats Effect, weighed against ZIO — with the goal of
simplifying xl while keeping what makes it beautiful, and optimizing it for LLM use: *Excel for
agents, by agents, with functional syntax.*

## TL;DR

1. **The effect system is not where xl's beauty lives.** The pure core (`ref"A1"`, `fx"…"`, the
   `Patch` monoid, `Sheet.put`, laws, opaque types) touches Cats in exactly two files, both for
   `Monoid`. Cats Effect and fs2 live only at the IO edge — and that edge is where every piece of
   ceremony an agent has to write comes from (`import cats.effect.unsafe.implicits.global`,
   `.compile.drain.unsafeRunSync()`).
2. **Ox is a poor fit for the published library, an excellent fit for the agent runner.** Ox is
   JVM-only (JDK 21 virtual threads). ADR-016 commits `xl-cats-effect` and the `xl` CLI to Scala
   Native and Scala.js, which Ox cannot follow. `xl-agent` is JVM-only, internal, and the only
   module with real concurrency (`parTraverseN`, `Queue`, `Dispatcher`, fibers) — Ox turns that
   code into plain sequential Scala.
3. **ZIO is a lateral move.** Same monadic shape as Cats Effect, same ceremony at the script edge
   (more, in fact: `Unsafe.unsafe { Runtime.default.unsafe.run(...) }`), a second ecosystem to
   carry, and Scala Native support that ZIO itself labels experimental. Its one real advantage — a
   typed error channel that could carry `XLError` — is already what `XLResult` gives the pure core.
4. **The simplification that keeps the beauty: an effect-free, direct-style published API.**
   `Excel.read`/`Excel.write` implemented directly (no IO runtime underneath), streaming as a
   pull-based row cursor (`Iterator[RowData]` + `AutoCloseable`), errors as `XLResult`, and a
   zero-dependency `boundary`-based block for sequential `XLResult` code. Cats Effect (fs2), Ox
   (Flow) and ZIO (ZStream) then become thin, optional adapter modules over the same cursor —
   users pick their ecosystem, agents never see one.
5. **Do now, regardless of the decision**: take `cats-laws` out of `xl-core`'s compile scope (today
   every consumer of `xl-core` receives ScalaCheck and Discipline on their compile classpath), drop
   the unused `cats-core` dependency from `xl-evaluator`, and stop erasing `XLError` into
   `new Exception(message)` at the `ExcelIO` boundary.

The rest of this document is the evidence.

---

## 1. What xl actually uses today (measured)

Main sources only, `xl*/src`, at commit `a9cdb73`.

| Module | Files / LOC | Cats / CE / fs2 footprint |
|---|---|---|
| `xl-core` | 112 / 16,212 | `import cats.Monoid` in `patch/Patch.scala` and `styles/patch/StylePatch.scala`. **That is all.** Declares `cats-core` **and `cats-laws` in compile scope** (no main source imports `cats.laws`), so the published `xl-core_3` POM drags `scalacheck`, `discipline-core`, `cats-kernel-laws` and `test-interface` onto every consumer's classpath. |
| `xl-ooxml` | 42 / 18,100 | None. |
| `xl-evaluator` | 62 / 21,279 | None. Declares `cats-core` but has zero `cats` imports. `recalculateParallel` (#520) uses `Executors.newFixedThreadPool` + `java.util.concurrent.Future` directly (`WorkbookEvaluator.scala:507`). |
| `xl` (prelude) | 2 / 150 | Re-exports the sync `Excel` facade, `ExcelIO`, `RowData`. |
| `xl-cats-effect` | 13 / 6,833 | The effect layer. `Sync[F].delay` ×53, `Async[F].raiseError` ×11, `Stream.eval` ×10, fs2-data-xml `XmlEvent` ×76 (write path). |
| `xl-cli` | 43 / 17,980 | `CommandIOApp` (decline-effect). `IO.raiseError` ×134, `IO.pure` ×119, `IO.fromEither` ×102, `IO.blocking` ×20, `.compile.drain` ×16, `Ref.of` ×2, `Resource.make` ×2, `parMapN` ×1. No fibers (every `.start` is `range.start`). |
| `xl-agent` | 46 / 9,286 | Real concurrency: `parTraverseN` ×5, `Ref` ×8, `Queue` ×2, `Dispatcher` ×4, `Fiber` ×2, `Clock` ×15, `Resource` ×6, `IO.blocking` ×55. |
| `xl-benchmarks` | 7 / 1,096 | `import cats.effect.IO` only to call `ExcelIO`. |

### 1.1 The pure core is already effect-free

`Monoid[Patch]` and `Monoid[StylePatch]` are the entire Cats surface of the 55k-LOC pure core.
Both types also expose `++` (CLAUDE.md documents the type-ascription gotcha that makes the Cats
`|+|` syntax awkward). `xl-evaluator` runs its parallel recalculation on a plain JDK thread pool
with deterministic folding — the core has already chosen "no effect system" for its own
concurrency.

### 1.2 The effect layer is mostly a wrapper

`ExcelIO` is `class ExcelIO[F[_]: Async]`, but `F` is instantiated **27 times in the repo, every
time as `IO`**; there is no second interpreter. Of its ~1,400 lines:

- `read`/`write`/`readMetadata`/`readDimension`/`loadStyles` are `Sync[F].delay(pureFunction(…))`
  followed by `raiseError(new Exception(s"Failed to …: ${err.message}"))` — **six places where the
  structured `XLError` is flattened into an exception message** at the exact boundary the purity
  charter cares about. (`ExcelR` keeps the `XLResult`, but nothing in the repo calls it.)
- The row-stream *read* path is a SAX parser on a dedicated thread pushing 1,024-row chunks into an
  `ArrayBlockingQueue`, drained by `Sync[F].interruptible(queue.take())` into an fs2 `Stream`
  (`SaxStreamingReader.scala:64-96`). `docs/plan/scala-native.md` already schedules deleting this
  bridge for a pull-based `WorksheetRowCursor`.
- The row-stream *write* path emits fs2-data-xml `XmlEvent`s that are rendered with
  `xml.render.raw()` and `fs2.text.utf8.encode` into a `ZipOutputStream` inside `Sync[F].delay`
  (`ExcelIO.scala:769-774`). The same plan replaces fs2-data-xml with an in-house `XmlTextWriter`.
- `StreamingTransform.scala` (1,355 LOC) and `StylePatcher.scala` (630 LOC) contain **no effect
  code at all**; `ZipTransformer` wraps them in four `Sync[F].delay` blocks.
- `PreservedPartStore` is hard-wired to `IO`/`Resource[IO, …]`, not `F`-polymorphic.
- The sync `Excel` facade that scripts use runs `unsafeRunSync()` on the global Cats Effect
  runtime to call what is, underneath, a pure `XlsxReader.read` — a full effect runtime is
  started so that a direct-style call can be made.

### 1.3 The CLI pays for a runtime it barely uses

`xl-cli` has 355 sites whose only job is lifting `Either`/values into `IO` (`IO.fromEither`,
`IO.pure`, `IO.raiseError`), one `parMapN` (`Main.scala:1846`), two `Ref`s (the unused REPL
`SessionRef` scaffolding), and 442 `unsafeRunSync()` calls in its tests. It also carries a
runtime-config override (`Main.scala:105-109`, GH-519) because the default Cats Effect shutdown
hook let a `SIGTERM`ed `xl` keep computing a recalculation at full CPU: for a batch compute
process the effect runtime was an obstacle, not a feature. `decline-effect` couples argument
parsing to that runtime.

### 1.4 The agent runner is the one real consumer

`xl-agent` schedules work units with `parTraverseN`, bridges the Anthropic Java SDK's blocking
`createStreaming` callbacks into `IO` through `Dispatcher.sequential` and a `Queue`
(`CodeExecution.scala:112-159`), runs a bounded console-printer fiber (`StreamingConsole`), tracks
per-turn state in `Ref`s, and uses `Resource` for uploaded files and skill ids. This is exactly the
workload structured concurrency exists for — and it is JVM-only by construction (Anthropic Java
SDK, OkHttp).

### 1.5 What agents are actually told to write

The LLM-facing surface is already direct-style, with one leak:

| Surface | Sync `Excel.*` | `ExcelIO` / `unsafeRunSync` |
|---|---|---|
| `plugin/skills/xl-scripting/SKILL.md` | 19 | 2 |
| `plugin/skills/xl-scripting/reference/RECIPES.md` | 12 | 1 (recipe 5, streaming) |
| `plugin/skills/xl-scripting/reference/API.md` | 4 | 2 ("Run effects at the script edge with `.unsafeRunSync()`") |
| `docs/QUICK-START.md` | 1 | 8 + 7 |
| `examples/*.sc` | 14 of 20 scripts mention `IO`, `ExcelIO` or `unsafeRunSync` | |

Recipe 5 is the tell. To filter a large file an agent must write four lines of incantation
(`import cats.effect.IO`, `import cats.effect.unsafe.implicits.global`, `.compile.drain`,
`.unsafeRunSync()`) that have nothing to do with Excel, and any mistake in them produces implicit-
resolution errors (`no given instance of type cats.effect.Async[F]`) that burn agent turns. The
README still lists "Cats Effect integration" under *Why XL?* — as a feature.

---

## 2. Constraints that decide the question

| Constraint | Source | Consequence |
|---|---|---|
| Purity charter: pure core, effects only in interpreters, `XLResult` everywhere | `purity-charter.md`, `style-guide.md` | Any choice must keep the core effect-free. All three candidates can; only the *edge* is in play. |
| **ADR-016: cross-compile `xl-core`, `xl-evaluator`, `xl-ooxml`, `xl-cats-effect`, `xl` to Scala Native 0.5 and Scala.js; the `xl` CLI to Scala Native** | `decisions.md`, `plan/scala-native.md` (locked goals, Wave 0 verified 2026-08-31) | Ox requires JDK 21 virtual threads and targets the JVM only. It cannot be the IO substrate of any module on the cross-build list. |
| `Excel[F]`/`ExcelIO` public API frozen during the streaming rework | `plan/scala-native.md` §Architecture | A new direct-style API must be **additive**; the `F[_]` API can be kept as an adapter and deprecated on its own schedule. |
| Scala 3.9 LTS, JDK 25 pinned | `build.mill`, `.mill-jvm-version` | Ox's JDK 21+ requirement is met on the JVM. JDK 24+ (JEP 491) also removed the `synchronized` pinning problem virtual threads had. |
| GraalVM native-image CLI today, Scala Native CLI later | `xl-cli/package.mill`, ADR-016 goal 1 | GraalVM for JDK 21+ supports virtual threads in native images, so Ox *would* work in the GraalVM CLI — but not in the Scala Native one the roadmap is heading to. |
| Optimize for LLM authorship | this request | Fewest concepts, stdlib shapes, plain compiler errors, no runtime plumbing. See §4. |

---

## 3. The candidates

### 3.1 Ox (1.0.x)

**What it is.** Direct-style structured concurrency, streaming and resiliency for Scala 3 on the
JVM, on top of JDK 21 virtual threads. Version checked: 1.0.6. Core pieces: `supervised` scopes with
`fork`/`join`, `par`/`parLimit`/`parEither`/`race`/`timeout`, `either:` blocks with `.ok()` (built
on `scala.util.boundary`), `Flow` (lazy, pull/push streaming with `map`/`filter`/`buffer`/
`runForeach`/`runToList`, `Flow.usingEmit`, `Flow.fromIterator`), `Channel`s for integrating
callback APIs, `retry`/`repeat` with schedules, `useInScope`/`useCloseableInScope` for resources.

**Fit, layer by layer.**

| Layer | Fit | Why |
|---|---|---|
| Pure core | n/a | Nothing to replace; Ox would add nothing. |
| Published IO layer (`xl-cats-effect` → cross-built) | **Poor** | JVM-only; conflicts with ADR-016. Also the wrong place to introduce a small-corpus dialect (§4). |
| CLI | **Unnecessary** | One `parMapN` and zero fibers. Plain direct-style Scala with `scala.util.Using` needs no library at all and stays Scala Native-portable. |
| `xl-agent` | **Excellent** | `parTraverseN(n)` → `parLimit(n)`; `Queue` + `Dispatcher.sequential` → a `Channel` fed from the SDK's blocking stream inside a `fork`; `StreamingConsole`'s fiber → a `fork` in a `supervised` scope; `Resource` → `useInScope`; API calls → `retry` with backoff. Every `IO[…]` signature (174 of them) becomes a plain return type. |
| Users who want streaming with backpressure | Good, as an **adapter** (`xl-ox`: `Flow.fromIterator(cursor)`) | ~100 LOC over the effect-free cursor. |

**Strengths.** Code reads top-to-bottom; real stack traces; blocking is free (virtual threads);
`either:`/`.ok()` is exactly the `XLResult` ergonomics scripts want; structured concurrency
guarantees no leaked threads; 1.0 API stability.

**Weaknesses.** JVM-only, JDK 21+; smaller ecosystem and much smaller training corpus than Cats
Effect/ZIO (agents will invent `ox.*` APIs unless the library keeps Ox out of the surface they
write against); `boundary`/`break` inside lambdas is a subtle footgun (`.ok()` in a `map` body is
a non-local jump); fs2 users lose their integration unless an fs2 adapter remains.

### 3.2 ZIO (2.1.x)

**What it is.** A monadic effect system with typed errors (`ZIO[R, E, A]`), fibers, `ZStream`,
`ZLayer`. Cross-compiles to Scala.js and Scala Native (ZIO's platform page marks Native as
experimental; recent releases track Scala Native 0.5.x).

**Fit.** ZIO would replace `Excel[F]`/fs2 with `ZIO[Any, XLError, Workbook]`/`ZStream[Any,
XLError, RowData]` — a *better* shape than today's `F[Workbook]`-plus-`Exception`, because the
error channel would finally carry `XLError`. But:

- It is the **same architecture** as Cats Effect: monadic edge, runtime, `unsafe.run` at the
  script boundary. Nothing gets simpler for agents; recipe 5 becomes
  `Unsafe.unsafe { implicit u => Runtime.default.unsafe.run(stream.run(sink)).getOrThrow() }`.
- It is a full rewrite of `xl-cats-effect`, the CLI and the agent for the *same* shape — cost
  without simplification — and existing Cats Effect users lose interop.
- ZIO's idioms (`ZLayer`, `R` environments, `provide`) are exactly what LLMs over-apply when they
  see `ZIO` in scope; a library that needs none of them invites noise.
- Typed errors can be had without ZIO: `XLResult` already is the typed error channel, and the
  direct-style API below keeps it.

**Verdict:** a lateral move. Only worth it if the team already lives in ZIO — it does not.

### 3.3 Cats Effect 3.7.x + fs2 (status quo)

**Fit.** Mature, cross-platform (Cats Effect 3.7 brought full multithreading to Scala Native 0.5;
the Wave 0 spike ran `unsafeRunSync` and `IO.blocking` on Native), the fs2 ecosystem, and it is
what exists.

**Costs, as measured in §1.** Generality (`F[_]: Async`) that has exactly one instance; structured
errors erased at the boundary; a sync facade that boots a runtime to call a pure function; a CLI
whose runtime needed tuning to honour `SIGTERM`; the ceremony that every streaming recipe hands to
an agent; and the dependency train below, pulled by everyone who depends on the aggregate `xl`
artifact. Note that ADR-016 itself forces a dependency bump wave (CE 3.7.1, fs2 3.13, fs2-data
1.14.1, munit-cats-effect 2.2.0, decline-effect 2.6.2) *because* these libraries sit in the
cross-built modules.

Transitive closure of the published `com.tjclp:xl_3:0.19.3` (compile/runtime scope, resolved from
Maven Central POMs; 30 artifacts, of which 4 are xl's own modules and 2 the Scala runtime):

| Brought in by | Third-party jars | Which |
|---|---|---|
| `cats-laws` in `xl-core` compile scope | 5 | cats-laws, cats-kernel-laws, discipline-core, scalacheck, test-interface |
| `cats-effect` | 3 | cats-effect, cats-effect-kernel, cats-effect-std |
| `fs2-core` / `fs2-io` | 5 | fs2-core, fs2-io, ip4s-core, scodec-bits, literally |
| `fs2-data-xml` / `fs2-data-xml-scala` | 8 | fs2-data-xml, fs2-data-xml-scala, fs2-data-text, fs2-data-finite-state, cats-collections-core, algebra, scala-collection-compat, portable-scala-reflect |
| `cats-core` (the two `Monoid` givens) | 2 | cats-core, cats-kernel |
| `scala-xml` (the OOXML DOM, stays) | 1 | scala-xml |
| **Total** | **24** | |

The effect-free proposal in §5 leaves `scala-xml` plus `cats-kernel` (or `cats-core` +
`cats-kernel`): **2–3 third-party jars instead of 24**, and `xl-core` alone goes from 7 to 1–2.

### 3.4 Gears (for completeness)

Scala Center's experimental direct-style library. It is the only direct-style option with a
Scala Native story (delimited continuations on SN 0.5), which makes it the one to **watch** for the
cross-built modules. It is explicitly "not recommended for production" today. Do not adopt; the
effect-free API below is designed so that a Gears adapter would be another ~100 LOC if it matures.

---

## 4. Optimizing for LLM authorship

What an agent writes against xl today: the `xl` CLI (stateless, JSON batch ops) and scala-cli
scripts through the `com.tjclp.xl.scripting` prelude. The design rules that matter for that
audience, ranked by how often they bite in traces:

1. **Fewest concepts.** Every extra concept — runtime, `compile`, `unsafeRunSync`, an implicit
   global import, `F[_]` — is another place to be wrong. The prelude already reduced this to one
   import; streaming is the last leak.
2. **Stdlib shapes.** `Either`, `Option`, `Iterator`, `Vector`, `Using` are the shapes every model
   knows cold. `fs2.Stream`, `ZStream`, `ox.flow.Flow` are all dialects; of the three, Ox has by
   far the smallest corpus, which is the argument for keeping Ox *inside* xl rather than in front
   of it.
3. **Plain compiler errors.** Direct-style code fails with "type mismatch: found `XLResult[Sheet]`,
   required `Sheet`" — actionable in one turn. Tagless-final code fails with implicit-search
   errors that point nowhere.
4. **Errors as values, unwrapped once.** `XLResult` + `.unsafe` at the edge is already right. What
   is missing is a way to write *sequential* `XLResult` code without nesting `for`/`flatMap`;
   Ox's `either:`/`.ok()` is the right idea and it is 30 lines of `scala.util.boundary` (xl already
   uses `boundary` in eight core files). xl can ship its own, zero-dependency, cross-platform.
5. **Real stack traces and top-to-bottom reading** help an agent debug its own script as much as a
   human.

By these rules the ranking for *what agents see* is: effect-free direct style with `XLResult` >
Ox > Cats Effect ≈ ZIO. The ranking for *what xl uses internally where concurrency is real*
(`xl-agent`) is: Ox > Cats Effect ≈ ZIO, because that module is JVM-only and its readers are the
maintainers.

---

## 5. Proposal: effect-agnostic core, direct-style edge

### 5.1 Shape

```
xl-core / xl-ooxml / xl-evaluator        pure, effect-free (unchanged)
xl-io   (new, cross-built)               direct-style IO: Excel facade, RowCursor, RowWriter,
                                          XLResult everywhere, java.nio + scala.util.Using
xl-cats-effect (kept, cross-built)        thin adapter: Excel[F]/ExcelIO over xl-io + fs2 Stream
xl-ox   (optional, JVM)                   thin adapter: Flow.fromIterator(cursor), Flow → RowWriter
xl-zio  (optional, cross-built)           thin adapter: ZStream.fromIterator(cursor)
xl      (aggregate + scripting prelude)   exports xl-io's Excel; no runtime underneath
xl-cli  (internal)                        effect-free direct style; decline (not decline-effect)
xl-agent (internal, JVM)                  Ox
```

Whether `xl-io` is a new module or the JVM leg of `xl-ooxml`'s platform split is a Wave A3/A4
detail; the `WorksheetRowCursor` and `XmlTextWriter` that ADR-016 already plans are its
streaming half.

### 5.2 The direct-style surface (illustrative; not compiled)

```scala
// Reading and writing: total, XLResult, no runtime. Scripts keep the throwing shorthand.
val wb: XLResult[Workbook] = Excel.readResult("in.xlsx")
val wb2: Workbook          = Excel.read("in.xlsx")            // unchanged for scripts: throws XLException
Excel.write(wb2, "out.xlsx")

// Streaming: a loan-pattern cursor; O(1) memory, nothing to compile or run.
val total = Excel.rows("huge.xlsx", sheet = "Data") { rows =>   // rows: Iterator[RowData]
  rows.flatMap(_.cells.get(2)).collect { case CellValue.Number(n) => n }.sum
}
Excel.writeRows("out.xlsx", "Filtered")(rows.filter(isLarge))     // rows: Iterator[RowData]

// Sequential XLResult code without nesting — a boundary-based block, zero dependencies.
val report: XLResult[Workbook] = XLResult.block {
  val sales   = wb("Sales").ok
  val summary = sales.readTyped[BigDecimal](ref"C1").ok
  wb.update("Summary", _.put(ref"A1", summary)).ok
}
```

Recipe 5 becomes three lines of Excel and zero lines of runtime. The same code compiles on Scala
Native and Scala.js (the cursor is `Iterator` + `AutoCloseable`; the ZIP/XML engines are the ones
ADR-016 is building anyway).

### 5.3 Adapters (illustrative)

```scala
// xl-cats-effect: Excel[F] survives as a wrapper, honouring the API freeze.
def readStream(path: Path): Stream[F, RowData] =
  Stream.resource(Resource.fromAutoCloseable(Sync[F].blocking(RowCursor.open(path))))
    .flatMap(c => Stream.fromBlockingIterator[F](c.iterator, chunkSize = 1024))

// xl-ox
def rows(path: Path): Flow[RowData] = Flow.usingEmit { emit =>
  Using.resource(RowCursor.open(path))(_.iterator.foreach(emit.apply))
}
```

### 5.4 `xl-agent` on Ox (illustrative)

```scala
// today: workUnits.parTraverseN(config.parallelism.max(1))(executeWorkUnit(_, …))
val unitResults = parLimit(config.parallelism.max(1))(workUnits.map(u => () => executeWorkUnit(u, …)))

// today: Dispatcher.sequential[IO].use { d => sdk.createStreaming(params).stream().forEach(ev => d.unsafeRunSync(processor.process(ev))) }
supervised {
  val events = Channel.bufferedDefault[AgentEvent]
  fork { Using.resource(sdk.createStreaming(params))(_.stream().forEach(ev => processor.process(ev, events))); events.done() }
  events.foreach(consoleAndTrace)
}
```

`ProcessorState` in a `Ref` becomes a plain `var` guarded by the single processing thread, or an
`AtomicReference`; `Clock[IO].realTime` becomes `Instant.now()`; `Resource.make(upload)(delete)`
becomes `useInScope(upload)(delete)`.

### 5.5 Monoid

Keep `Monoid[Patch]`/`Monoid[StylePatch]` as instances, but decide what they are instances *of*:

- **Keep `cats-kernel` only** (not `cats-core`): `Monoid` lives in `cats-kernel`, a small artifact
  with no further Scala dependencies; Cats users keep `|+|` and `combineAll`. Cheapest, keeps
  interop.
- **Own `Monoid` + optional Cats instance** in a `com.tjclp.xl.cats` package of the adapter module:
  zero core dependencies; Cats users import one given. More work, purer.

Either way, `cats-laws` leaves compile scope (it belongs to the test module, alongside the
existing `Generators` and law suites — today it puts ScalaCheck and Discipline on every
`xl-core` consumer's compile classpath) and `xl-evaluator` drops its unused `cats-core`.

### 5.6 Decision matrix

| Criterion | Effect-free API + Ox in agent (proposed) | Ox everywhere | ZIO | Cats Effect (status quo) |
|---|---|---|---|---|
| Purity charter (pure core, errors as values) | ✅ `XLResult` end-to-end | ✅ | ✅ typed `E` | ⚠️ `XLError` erased at edge |
| Scala Native + Scala.js (ADR-016) | ✅ stdlib only in cross-built modules | ❌ JVM-only | ⚠️ Native experimental | ✅ (CE 3.7) |
| LLM authorship (scripts, recipes) | ✅ stdlib shapes, no runtime | ⚠️ small corpus, `boundary` footguns | ❌ runtime + `Unsafe` ceremony | ❌ runtime + implicit global |
| Structured concurrency where needed (`xl-agent`) | ✅ Ox | ✅ Ox | ✅ fibers | ✅ fibers |
| Streaming for library users | ✅ cursor + adapters (fs2/Ox/ZIO) | Flow only | ZStream only | fs2 only |
| Dependency weight of `com.tjclp::xl` (third-party jars) | ✅ 2–3 (scala-xml + cats-kernel) | Ox core + scala-xml (JVM only) | ZIO + zio-streams + scala-xml | 24 today (§3.3), 19 after phase-0 hygiene |
| Ecosystem interop for existing users | ✅ `Excel[F]` kept as adapter | ❌ fs2 users cut | ❌ fs2 users cut | ✅ |
| Migration cost | Medium: new module + mechanical CLI rewrite + agent rewrite; adapters small | High: everything + platform conflict | High: everything, for the same shape | None |
| Risk | Additive API; freeze honoured; Ox contained to an internal module | Blocks ADR-016 | Second runtime; experimental Native | Carries today's costs into the cross-build |

---

## 6. Migration plan and cost

Phased so that every step is independently shippable and aligned with the ADR-016 waves.

| Phase | Work | Size | Notes |
|---|---|---|---|
| 0 — hygiene (no API change) | Move `cats-laws` to `xl-core.test`; remove `cats-core` from `xl-evaluator`; make `ExcelIO.read`/`write` raise `XLException(err)` instead of `new Exception(err.message)`; add `XLResult.block` (`boundary`-based) to `xl-core`; rewrite recipe 5 and the QUICK-START streaming snippets once §1 lands. | Small | Ships in the next patch. CHANGELOG entries under Unreleased. |
| 1 — direct-style IO (with Wave A3/A4) | `xl-io`: sync `Excel` implemented directly (no runtime), `RowCursor` (the planned `WorksheetRowCursor`), `RowWriter` over the planned `XmlTextWriter`, `Excel.rows`/`writeRows`. Prelude exports the new facade; `ExcelIO` stays. | ~2–3k LOC new, largely moved from `ExcelIO`/`StreamingXmlWriter`/`SaxStreamingReader` | The SAX-thread bridge and fs2-data-xml leave the hot path, as the native plan already intends. |
| 2 — adapter + CLI | `xl-cats-effect` becomes a wrapper over `xl-io` (API frozen, behaviour identical, parity specs already exist). `xl-cli` drops `decline-effect` for `decline`, `IO` for direct calls, `Ref` scaffolding removed; tests lose 442 `unsafeRunSync`. | CLI: 43 files, mechanical (355 lift sites); tests 17 suites | Also removes the GH-519 runtime override — process exit is the JVM's again. Prerequisite for the Scala Native CLI. |
| 3 — agent on Ox | `xl-agent` to Ox: `parLimit`, `Channel`, `supervised`, `retry`, `useInScope`. | 46 files; ~2–3k LOC of real concurrency, rest mechanical; 12 CE test suites | JVM-only module; JDK 25 satisfies Ox. Biggest readability win per line. |
| 4 — optional adapters | `xl-ox`, `xl-zio` if users ask. | ~100–200 LOC each | Each is a `fromIterator` plus resource handling. |

**Risks and mitigations.** Public-API expectations of Cats Effect users → `Excel[F]` remains, as an
adapter, for as long as the module exists; deprecate only after two minor releases with the
direct-style API. Ox API churn → Ox is 1.0 and confined to an internal module. Behavioural drift in
streaming → the existing streaming/in-memory parity specs and the real-file fixture corpus gate
phase 1, exactly as `plan/scala-native.md` gates the pull-parser swap.

---

## 7. Recommendation (proposed ADR-017)

- **Do not** replace Cats Effect with ZIO: same shape, same ceremony, a second ecosystem.
- **Do not** put Ox into any cross-built module: it is JVM-only and would block ADR-016.
- **Do** make the published API effect-free and direct-style (`XLResult`, `Iterator`-based
  streaming, a `boundary`-based block), which is what the scripting prelude already pretends to be
  and what agents write best against.
- **Do** keep `xl-cats-effect` as a thin adapter honouring the frozen `Excel[F]` API; add `xl-ox`
  and `xl-zio` adapters on demand.
- **Do** adopt Ox in `xl-agent`, the one module with real structured-concurrency needs, and keep
  the CLI free of any effect library.
- **Do now**: the phase-0 hygiene items.

The result keeps everything that makes xl beautiful — the pure, law-governed core and the
compile-time DSL — and moves the part agents stumble over out of their way.

---

## Sources

- Ox README and documentation: https://github.com/softwaremill/ox (version 1.0.6; "safe
  direct-style streaming, concurrency and resiliency for Scala on the JVM"; requires JDK 21+ and
  Scala 3), https://ox.softwaremill.com/latest/
- Cats Effect releases: https://github.com/typelevel/cats-effect/releases (3.7.x adds Scala Native
  0.5 multithreading; 3.6 blocks in place on virtual threads)
- ZIO platforms: https://zio.dev/overview/platforms/ (JVM, Scala.js, Scala Native — Native
  experimental); releases: https://github.com/zio/zio/releases
- Gears: https://lampepfl.github.io/gears/ (JVM 21+ and Scala Native 0.5 via delimited
  continuations; experimental)
- GraalVM for JDK 21 — virtual threads in Native Image: https://blogs.oracle.com/graal/oracle-graalvm-for-jdk-21
- Scala Native 0.5 changelog (system threads, delimited continuations):
  https://scala-native.org/en/stable/changelog/0.5.x/0.5.0.html
- In-repo: `docs/design/purity-charter.md`, `docs/design/io-modes.md`, `docs/design/decisions.md`
  (ADR-001, ADR-007, ADR-016), `docs/plan/scala-native.md`, `xl-cats-effect/src/com/tjclp/xl/io/*`,
  `xl-cli/src/com/tjclp/xl/cli/Main.scala`, `xl-agent/src/com/tjclp/xl/agent/**`,
  `plugin/skills/xl-scripting/**`.

Library versions were checked on 2026-09-07. The dependency closure in §3.3 was resolved from the
published POMs on Maven Central (compile/runtime scope, optional dependencies skipped, highest
version wins, as Coursier does); `./mill xl.showMvnDepsTree` reproduces it locally with exact
version reconciliation.
