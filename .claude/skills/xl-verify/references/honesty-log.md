# Honesty log

Seeded-bug checks (step 5 of the skill):
1. On a scratch branch, revert a merged fix, keeping its tests out.
2. Run xl-verify as if it were a new PR.
3. Record here whether it caught the bug.

| Date | Fix reverted (PR) | Caught? | Which flow caught it, or why it missed | Change made |
|---|---|---|---|---|
| 2026-11-23 | #914 (main 43808312 vs 2b8f849d) | Yes, both bugs | `remove-last-link` (#895): main kept B7's link and `<hyperlinks>`; #914 removed both. `range-link-survives-edit` (#912): main dropped the D2:E3 range link and its relationship; #914 kept both. Checklist item 2 also caught a case still open after #914: after `insert-rows 1`, the kept range link stays at D2:E3 instead of moving to D3:E4. | First run. `compare.sh` gained absolute output paths and `<OUT>` path normalization. |
