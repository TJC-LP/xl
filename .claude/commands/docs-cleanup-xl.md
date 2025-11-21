# XL Docs Cleanup & Plan Triage

You are working inside the `xl` repository. Your job is to **audit and clean up documentation**, with a special focus on:

1. **Plan files** – deciding which `docs/plan/*.md` are still needed vs ready to archive.
2. **API / structure changes** – finding docs that no longer match the current XL API and implementation and updating them.

You MUST use **codebase context** (Scala sources, tests, examples) and the existing docs hierarchy, not just timestamps.

---

## 0. Mental model of the docs layout (treat this as given)

- Canonical “truth” docs:
  - Root: `README.md`
  - Core status: `docs/STATUS.md`, `docs/LIMITATIONS.md`
  - Architecture: `docs/design/architecture.md`
- Plan docs:
  - Active: `docs/plan/` (current and future plans)
  - Archived / completed: `docs/archive/plan/` (P0–P8, P31, string interpolation, unified put API, etc.)
  - Index: `docs/plan/roadmap.md` (authoritative plan index and status tracker)
- Design docs: `docs/design/…`
- Reference docs: `docs/reference/…`
- Root misc: `docs/FAQ.md`, `docs/QUICK-START.md`, `docs/CONTRIBUTING.md`, `docs/reviews/*.md`, etc.

**Rule of thumb:** When in doubt about what XL “really does” today, trust:

1. The Scala sources (`xl-core`, `xl-ooxml`, `xl-cats-effect`),
2. The tests,
3. `docs/STATUS.md` + `docs/LIMITATIONS.md` + `docs/plan/roadmap.md`.

Plan docs and older deep dives are *secondary* and may be historical.

---

## 1. Build a plan index and normalize statuses

1. Read:
   - `docs/plan/README.md`
   - `docs/plan/roadmap.md`
   - `docs/STATUS.md`
   - `docs/LIMITATIONS.md`

2. Enumerate all plan files:

   - Active: every `*.md` in `docs/plan/` excluding `README.md` and `roadmap.md`.
   - Archived: everything under `docs/archive/plan/`.

3. For each active plan file in `docs/plan/`:

   - Parse its header for:
     - Phase / feature (e.g. “P6.6: Fix Streaming Reader”, “Type Class Consolidation for Easy Mode `put()` API”).
     - Status (look for lines like `Status: …`, “Status: Partially Complete …”, “Deferred indefinitely”, etc.).
     - “Last Updated” / “Date” / “Documented” metadata if present.

   - Cross-check with:
     - Its entry in `docs/plan/roadmap.md` (if present).
     - Relevant sections in `docs/STATUS.md` and `docs/LIMITATIONS.md` (do they claim this feature is already complete? partially done? deferred?).
     - The code and tests referenced in that plan (files called out in the plan itself).

4. Derive a **normalized status** for each plan:

   - `Complete` – The work described is implemented, tested, and reflected in `docs/STATUS.md` and/or design/reference docs.
   - `Active` – There is still meaningful implementation work planned that is not present in code/tests, and `roadmap.md` treats it as current.
   - `Deferred/Rejected` – The plan itself, or another doc (e.g. lazy evaluation review), explicitly says the plan is not going to be implemented as originally written, or has been replaced by a different approach.

Keep a small table in your scratchpad for each plan:

- Path
- Phase / feature name
- Status (plan header)
- Status (roadmap)
- Status (what code + tests actually show)
- Normalized status (Complete / Active / Deferred)

---

## 2. Use codebase context to classify plan files

For each active plan doc in `docs/plan/`, use the codebase as ground truth:

### 2.1 Identify the “owner” modules

From the plan contents, list the key modules, types, and functions it talks about.

Examples:

- **Streaming I/O (`docs/plan/streaming-improvements.md`):**
  - `xl-cats-effect/src/com/tjclp/xl/io/ExcelIO.scala`
  - `StreamingXmlReader.scala`
  - `StreamingXmlWriter.scala`
  - Streaming-related tests in `xl-cats-effect/test/src/com/tjclp/xl/io/*.scala`

- **Type-class `put()` plan (`docs/plan/type-class-put.md`):**
  - `xl-core/src/com/tjclp/xl/extensions.scala`
  - `xl-core/src/com/tjclp/xl/codec/CellWriter.scala`
  - `ExtensionsSpec.scala` (tests)
  - Easy Mode design doc: `docs/design/easy-mode-api.md`
  - Examples like `examples/easy-mode-demo.sc`

- **Lazy evaluation / builder pattern (`docs/plan/lazy-evaluation.md`, builder sections):**
  - Look for any `SheetBuilder` or related types in `xl-core` / `xl-evaluator` (if present).
  - Check whether the full optimizer described in the plan was *actually implemented* or explicitly deferred in the doc.

### 2.2 Decide if the plan is “Complete”

Treat a plan as **Complete** if most of the following are true:

- The plan’s main tasks are clearly implemented in the referenced code modules.
- Tests exist that match what the plan describes (e.g. round-trip tests, streaming memory tests, error-path tests, etc.).
- `docs/STATUS.md` and/or `docs/plan/roadmap.md` already describe this work as done or “✅ Complete”.
- There is no longer any to-do section that matters beyond small polish.

If so:

1. **Mark the plan for archival:**
   - Create or reuse a folder under `docs/archive/plan/` for that phase or feature (e.g. `docs/archive/plan/p6-streaming-improvements/` or similar).
   - Move the plan file there.
   - At the top of the moved file, add a short “Completed” banner:

     > **Status**: ✅ Completed – This plan has been fully implemented.
     > This document is retained as a historical design/implementation record.

2. In the original location (`docs/plan/`), either:
   - Remove the file, **or**
   - Replace it with a tiny stub that links to the archived version and to the relevant design/reference docs.

3. Update `docs/plan/roadmap.md`:
   - Move this plan into a “Completed (see archive)” section, linking to its new path.
   - Ensure it uses the same style as existing “Recently Completed (See Archive)” entries.

4. If `docs/STATUS.md` still references this plan as “future work”, update that section so that:
   - It describes the current behavior and points to design/reference docs, **not** the plan.

### 2.3 Keep “Active” plans lean and forward-looking

For a plan classified as **Active**:

1. Split the document into:
   - “What’s already implemented” – with links to code and tests.
   - “What remains” – concrete remaining tasks.

2. Remove or drastically shorten any sections that duplicate information in:
   - `docs/STATUS.md`
   - `docs/design/architecture.md`, `docs/design/io-modes.md`, etc.
   Use those “truth” docs as canonical; keep the plan focused on *future work*.

3. Ensure `docs/plan/roadmap.md` has:
   - The correct status (partially complete / in progress).
   - A short, accurate one-line summary and link.

### 2.4 Handle “Deferred/Rejected” plans

For a plan that is **Deferred/Rejected**:

1. Add a clear note at the top:

   > **Status**: ⏸ Deferred – not planned for implementation.
   > See `<link>` for the updated approach.

2. Move it under `docs/archive/plan/` (e.g. `docs/archive/plan/deferred/lazy-evaluation.md`) to keep `docs/plan/` focused on current work.

3. Update `docs/plan/roadmap.md` to move it into a “Deferred” section rather than active roadmap.

---

## 3. API- and structure-driven docs cleanup

Now focus on docs that describe the XL API and behavior and bring them in line with current code.

### 3.1 Use `docs/design/quick-wins.md` as your backlog

1. Read `docs/design/quick-wins.md` and fully apply its “Documentation Wins” section:
   - Top-level truth: keep `docs/STATUS.md`, `docs/LIMITATIONS.md`, `docs/design/architecture.md`, and root `README.md` authoritative.
   - Deprecate `cell"..."/putMixed` in docs; prefer `ref"A1"`, `ref"A1:B10"`, and batch `Sheet.put(ref -> value, ...)`.
   - Rewrite `docs/QUICK-START.md` and `docs/reference/migration-from-poi.md` to use the current Easy Mode API and macros.
   - Clarify streaming trade-offs consistently using `docs/design/io-modes.md` + `docs/reference/performance-guide.md` as canonical.

2. For each bullet, identify the concrete files to update and implement the changes directly.

### 3.2 Detect API drift by searching code vs docs

For each doc in:

- `docs/QUICK-START.md`
- `docs/STATUS.md`
- `docs/LIMITATIONS.md`
- `docs/reference/*.md`
- `docs/design/easy-mode-api.md`
- `docs/FAQ.md`
- `docs/plan/*` that describe APIs (not just roadmap)
- `examples/*.sc`

do the following:

1. **Extract all public API identifiers and patterns** mentioned in the doc:
   - Functions, methods, and objects: e.g. `Sheet("Name")`, `Workbook.empty`, `Workbook.put`, `ExcelIO.read`, `Excel.write`, `Sheet.put("A1", 42)`, `Excel.modify`, etc.
   - Macros: `ref"A1"`, `ref"A1:B10"`, `fx"=SUM(...)"`, `money"$1,234.56"`, `percent"15%"`.
   - Old/possibly obsolete names: `cell"A1"`, `range"A1:B10"`, `putMixed`, older streaming APIs, outdated package names.

2. **Search the code** (sources + tests) for these identifiers:
   - Confirm the symbol exists in the expected package and with the semantics the doc claims.
   - If a symbol used in the docs **does not exist** anymore, or exists with clearly different semantics, mark that doc section as **needing update**.

3. **Update the doc** to match the current API:

   - Replace old terminology with new:
     - `cell"..."/range"..."/putMixed` → `ref"..."/Sheet.put(ref -> value, ...)`, or Easy Mode `.put("A1", value)`.
   - Ensure code snippets compile *conceptually* against the current modules:
     - Imports: `import com.tjclp.xl.*` and/or `import com.tjclp.xl.easy.*` as appropriate.
     - Types and functions match what exists in `xl-core` and `xl-cats-effect`.

4. When you’re unsure of exact signatures:
   - Prefer examples drawn from actual tests (`ExtensionsSpec.scala`, `ExcelIOSpec.scala`, etc.) or from `examples/*.sc` scripts.
   - Preserve those examples as the single source of truth and align docs to them.

### 3.3 Easy Mode & type-class `put()` alignment

1. Open:
   - `docs/design/easy-mode-api.md`
   - `docs/QUICK-START.md`
   - `docs/reference/migration-from-poi.md`
   - `examples/easy-mode-demo.sc`
   - `xl-cats-effect/src/com/tjclp/xl/easy.scala`
   - `xl-cats-effect/src/com/tjclp/xl/io/EasyExcel.scala`
   - `xl-core/src/com/tjclp/xl/extensions.scala` and its tests.

2. Ensure the docs accurately reflect the **three-tier model** already described in Easy Mode doc:

   - Tier 1: Pure API (`com.tjclp.xl.*`) – `Either`-based, no IO.
   - Tier 2: Easy Mode extensions (string refs, throws `XLException` on invalid input).
   - Tier 3: IO boundary (`com.tjclp.xl.easy.*` → `Excel.read/write/modify`).

3. For the `type-class-put` plan:

   - If the refactor **is now implemented** in `extensions.scala` and covered by tests:
     - Update `docs/plan/type-class-put.md` to clearly say it has been implemented, and either:
       - Archive it as a completed plan (see §2.2), or
       - Convert it into a “design note” and move to `docs/design/` if it’s long-lived architecture.
     - Make sure Easy Mode docs and Quick Start mention the behavior this refactor provides: auto NumFmt inference, type-class extensibility, etc.

   - If it’s **not implemented yet**:
     - Keep it in `docs/plan/` but update the docs (Easy Mode, Quick Start, etc.) so they describe the current, still-duplicated overloads correctly and don’t promise type-class behavior that doesn’t yet exist.

### 3.4 Streaming docs & performance

1. Use `docs/plan/streaming-improvements.md`, `docs/design/io-modes.md`, `docs/reference/performance-guide.md`, and `docs/STATUS.md` together:

   - Make sure **all** descriptions of streaming read/write:
     - Acknowledge that streaming read has been fixed (uses fs2-based streaming, constant memory).
     - Clearly state current limitations (e.g. SST/styles support only in certain modes, if that’s the case).
   - Remove all remaining references to the old “broken streaming reader” behavior.

2. When describing performance targets or test counts, keep them internally consistent within the docs.
   - If you can’t confidently derive exact numbers from tests, prefer phrasing like “over N tests” instead of stale exact counts.

---

## 4. Generic docs cleanup (age, stubs, naming)

After plan triage and API alignment, do a pass over **all files under `docs/`** (including `docs/archive/`):

1. For each `.md` file, evaluate it as a **candidate for archival** if ANY of the following:

   - It’s clearly a **stub**:
     - Very short (e.g. < 50–100 words) and contains no links, examples, or concrete guidance.
   - The filename or title includes obviously “temporary” words:
     - `_old`, `_draft`, `backup`, `archive`, “scratch”, “notes”, etc.
   - It describes behavior that:
     - Directly contradicts `docs/STATUS.md`, and
     - Has been superseded by newer design or plan docs, and
     - Is not worth maintaining as a design history.

2. Before archiving such a file:

   - Search the code and other docs for references to it (by filename and by main heading).
   - If it is still referenced from canonical docs (`README.md`, `STATUS.md`, `LIMITATIONS.md`, core design docs), either:
     - Update it instead of archiving, or
     - Update the referencing docs to point somewhere else, then archive it.

3. For each file you decide to **archive**:

   - Move it under `docs/archive/`:
     - For plan-related docs, prefer `docs/archive/plan/...`.
     - For other docs, group under a reasonable subdirectory (e.g., `docs/archive/reference/` or `docs/archive/design/`).
   - Add or update `docs/ARCHIVE_LIST.md` (create it if it doesn’t exist) with a row:

     - original path
     - new path
     - reason (e.g., “stub”, “superseded by STATUS.md + easy-mode-api.md”, “completed plan”)
     - timestamp (today’s date in ISO format).

---

## 5. Keep the docs index authoritative

After all moves and edits:

1. Update `docs/plan/roadmap.md`:
   - Ensure every plan that still lives in `docs/plan/` is listed with an accurate status and concise scope.
   - Ensure every archived plan is either:
     - Listed in a “Completed (see archive)” section, or
     - Explicitly marked as “Deferred” with reasoning.

2. Update `docs/STATUS.md` and `docs/LIMITATIONS.md`:
   - Confirm they match the *actual* feature set and behavior implied by code + tests.
   - Link to:
     - `docs/design/io-modes.md` and `docs/reference/performance-guide.md` for streaming.
     - Easy Mode docs for the high-level API structure.
     - Any newly important design docs.

3. If needed, add a short “Documentation Overview” section to `docs/README.md` that explains:

   - Where to find:
     - Status & limitations,
     - Design docs,
     - Reference docs,
     - Plan docs vs archived plans.

---

## 6. Final output

At the end of a `/docs-cleanup-xl` run, output a human-readable summary like:

- Number of docs scanned.
- Plan docs:
  - How many moved from `docs/plan/` → `docs/archive/plan/` (list them).
  - How many updated in-place as Active.
  - How many marked as Deferred.
- Docs with API drift:
  - Which files had significant API example updates.
  - A short note for each (“Updated Quick Start to use Easy Mode macros and new `put` API”, etc.).
- Misc archival:
  - Which miscellaneous docs were archived and why (from `docs/ARCHIVE_LIST.md`).

Where useful, include small diff-style summaries (in natural language) for major docs you changed so a human maintainer can skim and decide whether to commit or tweak the edits.
