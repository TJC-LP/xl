# Diff-triggered checklist

Read the PR's diff, then run every item whose trigger matches as an extra flow in xl-verify. Each
item comes from a real bug; add one whenever a bug gets past review (step 5 of the skill). Keep
the trigger concrete (file, function or command), and the flow short.

| # | Trigger (the diff touches...) | Extra flow | From |
|---|---|---|---|
| 1 | Hyperlink reading or writing: `XlsxReader` hyperlink code, `OoxmlWorksheet` hyperlinks, `XlsxWriter.planSheetRels`, `unmodeledHyperlinksElem` | Round-trip one sheet that has a cell link, a range link (`ref="D2:E3"`) and a relationship id that collides with a generated one; make an unrelated `put`, then check every `<hyperlink>` and its `.rels` entry survives | #895, #912 |
| 2 | Hyperlinks plus any row or column edit (`insert-rows`, `delete-rows`, `insert-cols`, `delete-cols`, `sort`) | Insert a row above a kept range link and delete a column through it; the link's `ref` must move or shrink with the cells | #914 review |
| 3 | Setting a value, link or comment on a cell that sits inside a kept range-level object | Set a new link on one cell inside a range link; the file must carry one link for that cell, and the rest of the range keeps the old one | #914 review |
| 4 | Criteria evaluation: `CriteriaIndex`, `PassCache`, `FunctionSpecsAggregate`, `*IFS` functions | Date-window `SUMIFS`/`COUNTIFS` book (made-up generator), base vs branch, 5 timed runs; values identical, plus one equality-only column to show it keeps its index | #920, #922 |
| 5 | Any memory budget or cache size (`maxIndexBytes`, cache caps) | A column of about 1,000,000 distinct numbers read by a `>`/`<` criterion next to equality-only columns; equality calls must not fall back to scanning | #922 review |
| 6 | Recalc or audit paths (`recalc`, `audit`, `WorkbookEvaluator`, dependency or SCC code) | `xl recalc` and `xl audit --json` on a circular made-up book and on a 50k-row book: results identical to base, timings reported, and to the installed release when the PR claims a speedup | #891, #920 |
| 7 | `clear`, `unmerge` or style removal | Clear and unmerge styled, merged and row/column-styled ranges; formats the user did not ask to remove stay, and styles.xml has no orphan or duplicated entries | #886, #894, #897, #908, #909 |
| 8 | Sheet adoption or cross-workbook copy | Adopt a sheet whose name differs only by case from a hidden one, and one whose name a preserved non-worksheet tab already uses | #526 |
| 9 | Output path handling (`--output`, `--in-place`, path normalization) | Write to a relative path with `..` segments, to an existing read-only file, and in place; the file mode and location are what a user expects | #868, #870 |
| 10 | Any writer change | Write the same workbook twice; the bytes must be identical (determinism), and `lint --strict` stays clean | CLAUDE.md non-negotiable 3 |
